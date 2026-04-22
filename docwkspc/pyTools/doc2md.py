from __future__ import annotations

import argparse
import base64
import hashlib
import mimetypes
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, unquote_to_bytes, urlparse
from urllib.request import Request, urlopen

try:
	from markitdown import MarkItDown
except ImportError as exc:
	MarkItDown = None
	MARKITDOWN_IMPORT_ERROR = exc
else:
	MARKITDOWN_IMPORT_ERROR = None

try:
	import tkinter as tk
	from tkinter import filedialog, messagebox, ttk
	from tkinter.scrolledtext import ScrolledText
except ImportError:
	tk = None
	filedialog = None
	messagebox = None
	ttk = None
	ScrolledText = None


WORD_FILE_EXTENSIONS = {'.doc', '.docx'}
GUI_POLL_INTERVAL_MS = 150
USER_AGENT = 'Mozilla/5.0 (compatible; WordToMarkdownExporter/1.0)'
MARKDOWN_IMAGE_PATTERN = re.compile(
	r'!\[(?P<alt>[^\]]*)\]\((?P<url><[^>]+>|[^)\s]+)(?:\s+"(?P<title>[^"]*)")?\)'
)
HTML_IMAGE_PATTERN = re.compile(
	r'(?P<prefix><img\b[^>]*?\bsrc=["\'])(?P<url>[^"\']+)(?P<suffix>["\'])',
	re.IGNORECASE,
)
WINDOWS_DRIVE_PATTERN = re.compile(r'^[A-Za-z]:[\\/]')
POWERSHELL_DOC_TO_DOCX = r"""
$ErrorActionPreference = 'Stop'
$inputPath = $env:DOC2MD_INPUT_PATH
$outputPath = $env:DOC2MD_OUTPUT_PATH
$word = $null
$document = $null
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $false
	$word.DisplayAlerts = 0
	$document = $word.Documents.Open($inputPath)
	$wdFormatDocumentDefault = 16
	$document.SaveAs([ref]$outputPath, [ref]$wdFormatDocumentDefault)
}
finally {
	if ($document -ne $null) {
		$document.Close([ref]$false) | Out-Null
	}
	if ($word -ne $null) {
		$word.Quit() | Out-Null
	}
}
""".strip()


@dataclass(slots=True)
class ConversionResult:
	input_path: Path
	output_path: Path
	asset_dir: Path
	asset_path_prefix: str
	image_count: int
	converted_legacy_doc: bool
	warnings: list[str]


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description='Convert a Word document into Markdown with image files saved beside the output.'
	)
	parser.add_argument(
		'input',
		nargs='?',
		help='Word source file (.docx or .doc). If omitted, the GUI starts.',
	)
	parser.add_argument(
		'-o',
		'--output',
		help='Output Markdown path. Defaults to <input>.md.',
	)
	parser.add_argument(
		'--asset-dir',
		help='Directory used to store extracted image assets. Defaults to <output_stem>_assets beside the Markdown file.',
	)
	parser.add_argument(
		'--asset-path',
		help='Path prefix written into the Markdown for image assets. Defaults to the relative path from the Markdown file to --asset-dir.',
	)
	parser.add_argument(
		'--gui',
		action='store_true',
		help='Launch the graphical interface.',
	)
	return parser.parse_args()


def require_markitdown() -> None:
	if MarkItDown is None:
		raise RuntimeError(
			'markitdown is not installed in this Python environment. '
			'Install it with the docx extra, for example: pip install "markitdown[docx]"'
		) from MARKITDOWN_IMPORT_ERROR


def resolve_output_path(input_path: Path, output_arg: str | None) -> Path:
	if output_arg:
		output_path = Path(output_arg)
	else:
		output_path = input_path.with_suffix('.md')

	if not output_path.suffix:
		output_path = output_path.with_suffix('.md')

	return output_path.resolve()


def resolve_asset_dir(output_path: Path, asset_dir_arg: str | None) -> Path:
	if asset_dir_arg:
		asset_dir = Path(asset_dir_arg)
	else:
		asset_dir = output_path.parent / f'{output_path.stem}_assets'
	return asset_dir.resolve()


def resolve_asset_path_prefix(output_path: Path, asset_dir: Path, asset_path_arg: str | None) -> str:
	if asset_path_arg:
		asset_path = asset_path_arg.replace('\\', '/').strip()
		if asset_path in {'', '.', './'}:
			return ''
		return asset_path.rstrip('/')

	relative_path = Path(os.path.relpath(asset_dir, output_path.parent)).as_posix()
	return '' if relative_path == '.' else relative_path


def unwrap_markdown_url(url: str) -> str:
	stripped = url.strip()
	if stripped.startswith('<') and stripped.endswith('>'):
		return stripped[1:-1].strip()
	return stripped


def should_process_image(url: str) -> bool:
	lowered = url.lower()
	return bool(url) and not lowered.startswith(('mailto:', 'javascript:', '#', '//'))


def shorten_image_reference(url: str, max_length: int = 80) -> str:
	if not url.startswith('data:') or len(url) <= max_length:
		return url
	prefix = url.split(',', 1)[0]
	return f'{prefix},...'


def build_output_filename(source_key: str, original_name: str, content_type: str | None = None) -> str:
	safe_name = re.sub(r'[^0-9A-Za-z._-]+', '-', original_name).strip('-')

	guessed_ext = ''
	if safe_name:
		guessed_ext = Path(safe_name).suffix
	if not guessed_ext and content_type:
		guessed_ext = mimetypes.guess_extension(content_type.split(';', 1)[0].strip()) or ''
	if guessed_ext == '.jpe':
		guessed_ext = '.jpg'
	if not guessed_ext:
		guessed_ext = '.bin'

	if not safe_name:
		safe_name = f'image{guessed_ext}'
	elif not Path(safe_name).suffix:
		safe_name = f'{safe_name}{guessed_ext}'

	digest = hashlib.sha1(source_key.encode('utf-8')).hexdigest()[:10]
	return f'{digest}-{safe_name}'


def download_image(url: str, asset_dir: Path) -> Path:
	asset_dir.mkdir(parents=True, exist_ok=True)

	request = Request(url, headers={'User-Agent': USER_AGENT})
	with urlopen(request, timeout=30) as response:
		payload = response.read()
		content_type = response.headers.get('Content-Type')

	parsed = urlparse(url)
	original_name = Path(unquote(parsed.path)).name
	output_path = asset_dir / build_output_filename(url, original_name, content_type)
	if not output_path.exists():
		output_path.write_bytes(payload)
	return output_path


def resolve_local_image_source(url: str, source_dir: Path) -> Path:
	unwrapped_url = unquote(url)
	if WINDOWS_DRIVE_PATTERN.match(unwrapped_url) or unwrapped_url.startswith('\\\\'):
		return Path(unwrapped_url).resolve()

	parsed = urlparse(unwrapped_url)
	if parsed.scheme == 'file':
		path_text = unquote(parsed.path)
		if os.name == 'nt' and WINDOWS_DRIVE_PATTERN.match(path_text.lstrip('/')):
			path_text = path_text.lstrip('/')
		return Path(path_text).resolve()

	candidate = Path(unquote(parsed.path or unwrapped_url))
	if not candidate.is_absolute():
		candidate = (source_dir / candidate).resolve()
	return candidate


def copy_local_image(url: str, source_dir: Path, asset_dir: Path) -> Path:
	asset_dir.mkdir(parents=True, exist_ok=True)

	source_path = resolve_local_image_source(url, source_dir)
	if not source_path.exists() or not source_path.is_file():
		raise FileNotFoundError(f'Local image not found: {source_path}')

	output_path = asset_dir / build_output_filename(str(source_path), source_path.name)
	if not output_path.exists():
		shutil.copy2(source_path, output_path)
	return output_path


def parse_data_uri(data_uri: str) -> tuple[str, bytes]:
	if not data_uri.startswith('data:'):
		raise ValueError('Only data URIs are supported.')

	try:
		header, encoded_data = data_uri[5:].split(',', 1)
	except ValueError as exc:
		raise ValueError('Malformed data URI.') from exc

	header_parts = [part.strip() for part in header.split(';') if part.strip()]
	mime_type = header_parts[0] if header_parts and '/' in header_parts[0] else 'application/octet-stream'
	is_base64 = any(part.lower() == 'base64' for part in header_parts[1:])

	if is_base64:
		payload = base64.b64decode(encoded_data)
	else:
		payload = unquote_to_bytes(encoded_data)

	return mime_type, payload


def save_data_uri_image(data_uri: str, asset_dir: Path) -> Path:
	asset_dir.mkdir(parents=True, exist_ok=True)

	content_type, payload = parse_data_uri(data_uri)
	output_path = asset_dir / build_output_filename(data_uri, 'embedded-image', content_type)
	if not output_path.exists():
		output_path.write_bytes(payload)
	return output_path


def build_asset_reference(asset_filename: str, asset_path_prefix: str) -> str:
	if not asset_path_prefix:
		return asset_filename
	return f"{asset_path_prefix.rstrip('/')}/{asset_filename}"


def materialize_image(
	url: str,
	source_dir: Path,
	asset_dir: Path,
	asset_path_prefix: str,
	warnings: list[str],
	cache: dict[str, str],
	saved_assets: set[str],
) -> str:
	normalized_url = unwrap_markdown_url(url)
	if not should_process_image(normalized_url):
		return normalized_url

	if normalized_url not in cache:
		try:
			parsed = urlparse(normalized_url)
			if normalized_url.startswith('data:'):
				image_path = save_data_uri_image(normalized_url, asset_dir)
			elif parsed.scheme in {'http', 'https'}:
				image_path = download_image(normalized_url, asset_dir)
			else:
				image_path = copy_local_image(normalized_url, source_dir, asset_dir)
			cache[normalized_url] = build_asset_reference(image_path.name, asset_path_prefix)
			saved_assets.add(image_path.name)
		except Exception as exc:  # noqa: BLE001
			warnings.append(f'Failed to materialize {shorten_image_reference(normalized_url)}: {exc}')
			cache[normalized_url] = normalized_url

	return cache[normalized_url]


def localize_images(
	markdown_text: str,
	source_dir: Path,
	asset_dir: Path,
	asset_path_prefix: str,
) -> tuple[str, int, list[str]]:
	warnings: list[str] = []
	cache: dict[str, str] = {}
	saved_assets: set[str] = set()

	def replace_markdown_image(match: re.Match[str]) -> str:
		alt_text = match.group('alt')
		title_text = match.group('title')
		url = match.group('url')
		localized_url = materialize_image(
			url,
			source_dir,
			asset_dir,
			asset_path_prefix,
			warnings,
			cache,
			saved_assets,
		)
		title_suffix = f' "{title_text}"' if title_text else ''
		return f'![{alt_text}]({localized_url}{title_suffix})'

	def replace_html_image(match: re.Match[str]) -> str:
		url = match.group('url')
		localized_url = materialize_image(
			url,
			source_dir,
			asset_dir,
			asset_path_prefix,
			warnings,
			cache,
			saved_assets,
		)
		return f"{match.group('prefix')}{localized_url}{match.group('suffix')}"

	localized = MARKDOWN_IMAGE_PATTERN.sub(replace_markdown_image, markdown_text)
	localized = HTML_IMAGE_PATTERN.sub(replace_html_image, localized)
	return localized, len(saved_assets), warnings


def finalize_markdown(markdown_text: str) -> str:
	normalized = markdown_text.replace('\r\n', '\n')
	if normalized and not normalized.endswith('\n'):
		normalized += '\n'
	return normalized


def convert_doc_to_docx(input_path: Path, temp_dir: Path) -> Path:
	if os.name != 'nt':
		raise RuntimeError('Legacy .doc conversion requires Windows and Microsoft Word.')

	output_path = temp_dir / f'{input_path.stem}.docx'
	env = os.environ.copy()
	env['DOC2MD_INPUT_PATH'] = str(input_path)
	env['DOC2MD_OUTPUT_PATH'] = str(output_path)

	completed = subprocess.run(
		['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', POWERSHELL_DOC_TO_DOCX],
		capture_output=True,
		text=True,
		env=env,
		check=False,
	)
	if completed.returncode != 0 or not output_path.exists():
		detail = (completed.stderr or completed.stdout or '').strip()
		if not detail:
			detail = 'Microsoft Word automation did not produce a .docx file.'
		raise RuntimeError(f'Failed to convert legacy .doc to .docx: {detail}')

	return output_path


def prepare_input_for_conversion(input_path: Path) -> tuple[Path, tempfile.TemporaryDirectory[str] | None, bool]:
	suffix = input_path.suffix.lower()
	if suffix not in WORD_FILE_EXTENSIONS:
		raise ValueError('Input file must be a .docx or .doc file.')

	if suffix == '.docx':
		return input_path, None, False

	temp_dir = tempfile.TemporaryDirectory(prefix='doc2md-')
	temp_docx_path = convert_doc_to_docx(input_path, Path(temp_dir.name))
	return temp_docx_path, temp_dir, True


def convert_word_document(
	input_path: str | Path,
	output_arg: str | None = None,
	asset_dir_arg: str | None = None,
	asset_path_arg: str | None = None,
) -> ConversionResult:
	require_markitdown()

	resolved_input_path = Path(input_path).resolve()
	if not resolved_input_path.exists():
		raise FileNotFoundError(f'Input Word document not found: {resolved_input_path}')

	output_path = resolve_output_path(resolved_input_path, output_arg)
	asset_dir = resolve_asset_dir(output_path, asset_dir_arg)
	asset_path_prefix = resolve_asset_path_prefix(output_path, asset_dir, asset_path_arg)

	temporary_directory: tempfile.TemporaryDirectory[str] | None = None
	converted_legacy_doc = False
	try:
		prepared_input_path, temporary_directory, converted_legacy_doc = prepare_input_for_conversion(resolved_input_path)

		converter = MarkItDown(enable_plugins=False)
		result = converter.convert(prepared_input_path, keep_data_uris=True)
		markdown_text = finalize_markdown(result.text_content or '')
		localized_markdown, image_count, warnings = localize_images(
			markdown_text,
			resolved_input_path.parent,
			asset_dir,
			asset_path_prefix,
		)

		output_path.parent.mkdir(parents=True, exist_ok=True)
		output_path.write_text(finalize_markdown(localized_markdown), encoding='utf-8')

		return ConversionResult(
			input_path=resolved_input_path,
			output_path=output_path,
			asset_dir=asset_dir,
			asset_path_prefix=asset_path_prefix,
			image_count=image_count,
			converted_legacy_doc=converted_legacy_doc,
			warnings=warnings,
		)
	finally:
		if temporary_directory is not None:
			temporary_directory.cleanup()


def print_conversion_summary(result: ConversionResult) -> None:
	print(f'Markdown written to: {result.output_path}')
	print(f'Image paths used in Markdown: {result.asset_path_prefix or "."}')
	if result.converted_legacy_doc:
		print('Legacy .doc input was converted to a temporary .docx via Microsoft Word.')
	if result.image_count:
		print(f'Extracted images stored in: {result.asset_dir}')
		print(f'Extracted image count: {result.image_count}')
	else:
		print('No images were materialized.')

	if result.warnings:
		print('\nWarnings:', file=sys.stderr)
		for item in result.warnings:
			print(f'- {item}', file=sys.stderr)


class WordToMarkdownApp:
	def __init__(self, root: Any, initial_args: argparse.Namespace | None = None) -> None:
		self.root = root
		self.root.title('Word to Markdown Exporter')
		self.root.geometry('920x680')
		self.root.minsize(820, 560)

		self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
		self.worker: threading.Thread | None = None
		self.controls: list[Any] = []

		self.input_var = tk.StringVar(value=(getattr(initial_args, 'input', '') or ''))
		self.output_var = tk.StringVar(value=(getattr(initial_args, 'output', '') or ''))
		self.asset_dir_var = tk.StringVar(value=(getattr(initial_args, 'asset_dir', '') or ''))
		self.asset_path_var = tk.StringVar(value=(getattr(initial_args, 'asset_path', '') or ''))
		self.status_var = tk.StringVar(value='Select a Word document to start.')

		self._configure_style()
		self._build_ui()
		self._apply_default_paths()
		self.root.after(GUI_POLL_INTERVAL_MS, self._poll_worker)


	def _configure_style(self) -> None:
		style = ttk.Style(self.root)
		try:
			style.theme_use('clam')
		except tk.TclError:
			pass
		style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'))


	def _build_ui(self) -> None:
		self.root.columnconfigure(0, weight=1)
		self.root.rowconfigure(0, weight=1)

		main_frame = ttk.Frame(self.root, padding=18)
		main_frame.grid(row=0, column=0, sticky='nsew')
		main_frame.columnconfigure(0, weight=1)
		main_frame.rowconfigure(2, weight=1)

		ttk.Label(main_frame, text='Word to Markdown Exporter', style='Title.TLabel').grid(
			row=0,
			column=0,
			sticky='w',
		)
		ttk.Label(
			main_frame,
			text='Leave Output, Asset Directory, or Asset Path blank to let the script derive sensible defaults. Legacy .doc input requires Microsoft Word on Windows so it can be saved to a temporary .docx first.',
			wraplength=820,
		).grid(row=1, column=0, sticky='w', pady=(6, 14))

		form_frame = ttk.Frame(main_frame)
		form_frame.grid(row=2, column=0, sticky='nsew')
		form_frame.columnconfigure(1, weight=1)
		main_frame.rowconfigure(2, weight=0)

		self._add_path_row(form_frame, 0, 'Word Input', self.input_var, self._browse_input_file)
		self._add_path_row(form_frame, 1, 'Markdown Output', self.output_var, self._browse_output_file)
		self._add_path_row(form_frame, 2, 'Asset Directory', self.asset_dir_var, self._browse_asset_directory)
		self._add_entry_row(form_frame, 3, 'Markdown Asset Path', self.asset_path_var)

		button_frame = ttk.Frame(main_frame)
		button_frame.grid(row=3, column=0, sticky='w', pady=(16, 12))

		auto_button = ttk.Button(button_frame, text='Fill Default Paths', command=lambda: self._apply_default_paths(force=True))
		auto_button.grid(row=0, column=0, padx=(0, 10))
		self.controls.append(auto_button)

		convert_button = ttk.Button(button_frame, text='Convert', command=self._start_conversion)
		convert_button.grid(row=0, column=1)
		self.controls.append(convert_button)

		self.log_widget = ScrolledText(main_frame, height=16, wrap='word', font=('Consolas', 10))
		self.log_widget.grid(row=4, column=0, sticky='nsew', pady=(0, 10))
		self.log_widget.configure(state='disabled')
		main_frame.rowconfigure(4, weight=1)

		status_label = ttk.Label(main_frame, textvariable=self.status_var)
		status_label.grid(row=5, column=0, sticky='w')


	def _add_path_row(self, parent: Any, row: int, label: str, variable: Any, browse_command) -> None:
		ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=6, padx=(0, 12))
		ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky='ew', pady=6)
		browse_button = ttk.Button(parent, text='Browse...', command=browse_command)
		browse_button.grid(row=row, column=2, sticky='ew', pady=6, padx=(10, 0))
		self.controls.append(browse_button)


	def _add_entry_row(self, parent: Any, row: int, label: str, variable: Any) -> None:
		ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=6, padx=(0, 12))
		ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky='ew', pady=6)


	def _browse_input_file(self) -> None:
		selected = filedialog.askopenfilename(
			title='Select Word File',
			filetypes=[('Word Files', '*.docx *.doc'), ('All Files', '*.*')],
		)
		if selected:
			self.input_var.set(selected)
			self._apply_default_paths()


	def _browse_output_file(self) -> None:
		initial_name = ''
		input_value = self.input_var.get().strip()
		if input_value:
			initial_name = Path(input_value).with_suffix('.md').name
		selected = filedialog.asksaveasfilename(
			title='Select Markdown Output',
			defaultextension='.md',
			initialfile=initial_name,
			filetypes=[('Markdown Files', '*.md'), ('All Files', '*.*')],
		)
		if selected:
			self.output_var.set(selected)
			self._apply_default_paths()


	def _browse_asset_directory(self) -> None:
		selected = filedialog.askdirectory(title='Select Asset Directory')
		if selected:
			self.asset_dir_var.set(selected)
			self._apply_default_paths()


	def _apply_default_paths(self, force: bool = False) -> None:
		input_value = self.input_var.get().strip()
		if not input_value:
			return

		try:
			input_path = Path(input_value).resolve()
		except (OSError, RuntimeError):
			return

		output_value = self.output_var.get().strip()
		output_path = resolve_output_path(input_path, output_value or None)
		if force or not output_value:
			self.output_var.set(str(output_path))

		asset_dir_value = self.asset_dir_var.get().strip()
		asset_dir = resolve_asset_dir(output_path, asset_dir_value or None)
		if force or not asset_dir_value:
			self.asset_dir_var.set(str(asset_dir))

		asset_path_value = self.asset_path_var.get().strip()
		asset_path = resolve_asset_path_prefix(output_path, asset_dir, asset_path_value or None)
		if force or not asset_path_value:
			self.asset_path_var.set(asset_path)


	def _start_conversion(self) -> None:
		if self.worker and self.worker.is_alive():
			return

		input_value = self.input_var.get().strip()
		if not input_value:
			messagebox.showerror('Missing Input', 'Please select a Word file first.')
			return

		self._set_busy(True)
		self.status_var.set('Converting...')
		self._append_log('Starting conversion...')

		worker_kwargs = {
			'input_path': input_value,
			'output_arg': self.output_var.get().strip() or None,
			'asset_dir_arg': self.asset_dir_var.get().strip() or None,
			'asset_path_arg': self.asset_path_var.get().strip() or None,
		}

		self.worker = threading.Thread(target=self._run_conversion_worker, kwargs=worker_kwargs, daemon=True)
		self.worker.start()


	def _run_conversion_worker(self, **kwargs) -> None:
		try:
			result = convert_word_document(**kwargs)
			self.result_queue.put(('success', result))
		except Exception:
			self.result_queue.put(('error', traceback.format_exc()))


	def _poll_worker(self) -> None:
		try:
			while True:
				event_name, payload = self.result_queue.get_nowait()
				if event_name == 'success':
					self._handle_success(payload)
				else:
					self._handle_error(payload)
		except queue.Empty:
			pass
		finally:
			self.root.after(GUI_POLL_INTERVAL_MS, self._poll_worker)


	def _handle_success(self, result: ConversionResult) -> None:
		self.output_var.set(str(result.output_path))
		self.asset_dir_var.set(str(result.asset_dir))
		self.asset_path_var.set(result.asset_path_prefix)

		self._append_log(f'Markdown written to: {result.output_path}')
		self._append_log(f'Image paths used in Markdown: {result.asset_path_prefix or "."}')
		if result.converted_legacy_doc:
			self._append_log('Legacy .doc input was converted to a temporary .docx via Microsoft Word.')
		if result.image_count:
			self._append_log(f'Extracted images stored in: {result.asset_dir}')
			self._append_log(f'Extracted image count: {result.image_count}')
		else:
			self._append_log('No images were materialized.')
		if result.warnings:
			for warning in result.warnings:
				self._append_log(f'Warning: {warning}')
		else:
			self._append_log('Finished without warnings.')

		self.status_var.set('Conversion completed.')
		self._set_busy(False)
		messagebox.showinfo('Done', f'Markdown generated successfully:\n{result.output_path}')


	def _handle_error(self, trace_text: str) -> None:
		self._append_log('Conversion failed.')
		self._append_log(trace_text.rstrip())
		self.status_var.set('Conversion failed.')
		self._set_busy(False)
		messagebox.showerror('Conversion Failed', trace_text.rstrip().splitlines()[-1])


	def _set_busy(self, is_busy: bool) -> None:
		new_state = 'disabled' if is_busy else 'normal'
		for control in self.controls:
			control.configure(state=new_state)


	def _append_log(self, message: str) -> None:
		self.log_widget.configure(state='normal')
		self.log_widget.insert('end', f'{message}\n')
		self.log_widget.see('end')
		self.log_widget.configure(state='disabled')


def launch_gui(initial_args: argparse.Namespace | None = None) -> int:
	if tk is None or ttk is None or filedialog is None or messagebox is None or ScrolledText is None:
		print('Tkinter is not available in this Python environment.', file=sys.stderr)
		return 1

	try:
		root = tk.Tk()
	except tk.TclError as exc:
		print(f'Unable to start GUI: {exc}', file=sys.stderr)
		return 1

	WordToMarkdownApp(root, initial_args)
	root.mainloop()
	return 0


def main() -> int:
	args = parse_args()

	if args.gui or not args.input:
		return launch_gui(args)

	result = convert_word_document(
		input_path=args.input,
		output_arg=args.output,
		asset_dir_arg=args.asset_dir,
		asset_path_arg=args.asset_path,
	)
	print_conversion_summary(result)
	return 0


if __name__ == '__main__':
	raise SystemExit(main())
