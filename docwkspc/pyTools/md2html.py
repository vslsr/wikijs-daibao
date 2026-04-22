# CLI example:
# python.exe content/export_basic_html.py input.md --output out/page.html --asset-dir out/files --asset-path assets
# GUI example:
# python.exe content/export_basic_html.py --gui

from __future__ import annotations

import argparse
import hashlib
import html
import mimetypes
import os
import queue
import re
import shutil
import sys
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

import markdown

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


MARKDOWN_IMAGE_PATTERN = re.compile(
    r'!\[(?P<alt>[^\]]*)\]\((?P<url><[^>]+>|[^)\s]+)(?:\s+"(?P<title>[^"]*)")?\)'
)
HTML_IMAGE_PATTERN = re.compile(
    r'(?P<prefix><img\b[^>]*?\bsrc=["\'])(?P<url>[^"\']+)(?P<suffix>["\'])',
    re.IGNORECASE,
)
TITLE_PATTERN = re.compile(r'^\s*#\s+(.+?)\s*$', re.MULTILINE)
USER_AGENT = 'Mozilla/5.0 (compatible; MarkdownToHtmlExporter/1.0)'
GUI_POLL_INTERVAL_MS = 150


@dataclass(slots=True)
class ConversionResult:
    input_path: Path
    output_path: Path
    asset_dir: Path
    asset_path_prefix: str
    title: str
    warnings: list[str]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='Convert a Markdown document into a standalone HTML file and materialize its images.'
    )
    parser.add_argument(
        'input',
        nargs='?',
        help='Markdown source file. If omitted, the GUI starts.',
    )
    parser.add_argument(
        '-o',
        '--output',
        help='Output HTML path. Defaults to <input>.html.',
    )
    parser.add_argument(
        '--asset-dir',
        help='Directory used to store image assets. Defaults to <output_stem>_assets beside the HTML file.',
    )
    parser.add_argument(
        '--asset-path',
        help='Path prefix written into the generated HTML for image assets. Defaults to the relative path from the HTML file to --asset-dir.',
    )
    parser.add_argument(
        '--title',
        help='Override the generated HTML <title>. Defaults to the first Markdown H1 or the input filename.',
    )
    parser.add_argument(
        '--gui',
        action='store_true',
        help='Launch the graphical interface.',
    )
    return parser.parse_args()


def extract_title(markdown_text: str, fallback: str) -> str:
    match = TITLE_PATTERN.search(markdown_text)
    if not match:
        return fallback
    title = match.group(1).strip()
    title = re.sub(r'`([^`]+)`', r'\1', title)
    title = re.sub(r'<[^>]+>', '', title)
    return title or fallback


def resolve_output_path(input_path: Path, output_arg: str | None) -> Path:
    if output_arg:
        output_path = Path(output_arg)
    else:
        output_path = input_path.with_suffix('.html')

    if not output_path.suffix:
        output_path = output_path.with_suffix('.html')

    return output_path.resolve()


def resolve_asset_dir(output_path: Path, asset_dir_arg: str | None) -> Path:
    if asset_dir_arg:
        asset_dir = Path(asset_dir_arg)
        if not asset_dir.is_absolute():
            asset_dir = output_path.parent / asset_dir
    else:
        asset_dir = output_path.parent / f'{output_path.stem}_assets'
    return asset_dir.resolve()


def normalize_asset_path(value: str) -> str:
    asset_path = value.replace('\\', '/').strip()
    if asset_path in {'', '.', './'}:
        return ''
    return asset_path.rstrip('/')


def resolve_asset_path_prefix(
    output_path: Path,
    asset_dir: Path,
    asset_path_arg: str | None,
    asset_dir_arg: str | None = None,
) -> str:
    if asset_path_arg:
        return normalize_asset_path(asset_path_arg)

    if asset_dir_arg:
        raw_asset_dir = Path(asset_dir_arg)
        if not raw_asset_dir.is_absolute():
            return normalize_asset_path(asset_dir_arg)

    relative_path = Path(os.path.relpath(asset_dir, output_path.parent)).as_posix()
    if relative_path == '.':
        return ''
    return relative_path if relative_path.startswith('.') else f'./{relative_path}'


def unwrap_markdown_url(url: str) -> str:
    stripped = url.strip()
    if stripped.startswith('<') and stripped.endswith('>'):
        return stripped[1:-1].strip()
    return stripped


def should_process_image(url: str) -> bool:
    lowered = url.lower()
    return bool(url) and not lowered.startswith(('data:', 'mailto:', 'javascript:', '#', '//'))


def build_output_filename(source_key: str, original_name: str, content_type: str | None = None) -> str:
    safe_name = re.sub(r'[^0-9A-Za-z._-]+', '-', original_name).strip('-')

    guessed_ext = ''
    if safe_name:
        guessed_ext = Path(safe_name).suffix
    if not guessed_ext and content_type:
        guessed_ext = mimetypes.guess_extension(content_type.split(';', 1)[0].strip()) or ''
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


def resolve_local_image_source(url: str, markdown_dir: Path) -> Path:
    parsed = urlparse(url)
    if parsed.scheme == 'file':
        path_text = unquote(parsed.path)
        if os.name == 'nt' and re.match(r'^/[A-Za-z]:/', path_text):
            path_text = path_text[1:]
        return Path(path_text).resolve()

    candidate = Path(unquote(parsed.path))
    if not candidate.is_absolute():
        candidate = (markdown_dir / candidate).resolve()
    return candidate


def copy_local_image(url: str, markdown_dir: Path, asset_dir: Path) -> Path:
    asset_dir.mkdir(parents=True, exist_ok=True)

    source_path = resolve_local_image_source(url, markdown_dir)
    if not source_path.exists() or not source_path.is_file():
        raise FileNotFoundError(f'Local image not found: {source_path}')

    output_path = asset_dir / build_output_filename(str(source_path), source_path.name)
    if not output_path.exists():
        shutil.copy2(source_path, output_path)
    return output_path


def build_html_asset_reference(asset_filename: str, asset_path_prefix: str) -> str:
    if not asset_path_prefix:
        return asset_filename
    return f"{asset_path_prefix.rstrip('/')}/{asset_filename}"


def materialize_image(
    url: str,
    markdown_dir: Path,
    asset_dir: Path,
    asset_path_prefix: str,
    warnings: list[str],
    cache: dict[str, str],
) -> str:
    normalized_url = unwrap_markdown_url(url)
    if not should_process_image(normalized_url):
        return normalized_url

    if normalized_url not in cache:
        try:
            parsed = urlparse(normalized_url)
            if parsed.scheme in {'http', 'https'}:
                image_path = download_image(normalized_url, asset_dir)
            else:
                image_path = copy_local_image(normalized_url, markdown_dir, asset_dir)
            cache[normalized_url] = build_html_asset_reference(image_path.name, asset_path_prefix)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f'Failed to materialize {normalized_url}: {exc}')
            cache[normalized_url] = normalized_url
    return cache[normalized_url]


def localize_images(
    markdown_text: str,
    markdown_dir: Path,
    asset_dir: Path,
    asset_path_prefix: str,
) -> tuple[str, list[str]]:
    warnings: list[str] = []
    cache: dict[str, str] = {}

    def replace_markdown_image(match: re.Match[str]) -> str:
        alt_text = match.group('alt')
        title_text = match.group('title')
        url = match.group('url')
        localized_url = materialize_image(url, markdown_dir, asset_dir, asset_path_prefix, warnings, cache)
        title_suffix = f' "{title_text}"' if title_text else ''
        return f'![{alt_text}]({localized_url}{title_suffix})'

    def replace_html_image(match: re.Match[str]) -> str:
        url = match.group('url')
        localized_url = materialize_image(url, markdown_dir, asset_dir, asset_path_prefix, warnings, cache)
        return f"{match.group('prefix')}{localized_url}{match.group('suffix')}"

    localized = MARKDOWN_IMAGE_PATTERN.sub(replace_markdown_image, markdown_text)
    localized = HTML_IMAGE_PATTERN.sub(replace_html_image, localized)
    return localized, warnings


def render_markdown(markdown_text: str) -> str:
    renderer = markdown.Markdown(
        extensions=[
            'extra',
            'sane_lists',
            'toc',
        ],
        output_format='html5',
    )
    return renderer.convert(markdown_text)


def wrap_html_document(title: str, body_html: str) -> str:
    safe_title = html.escape(title)
    return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{safe_title}</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f4f6fb;
      --surface: #ffffff;
      --text: #18212f;
      --muted: #5a6575;
      --border: #d8dfeb;
      --accent: #0b5fff;
      --code-bg: #101828;
      --code-text: #e5edf8;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      background: linear-gradient(180deg, #eef3ff 0%, var(--bg) 100%);
      color: var(--text);
      line-height: 1.75;
    }}
    main {{
      max-width: 1100px;
      margin: 40px auto;
      padding: 40px 48px;
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 18px;
      box-shadow: 0 18px 50px rgba(15, 23, 42, 0.08);
    }}
    h1, h2, h3, h4, h5, h6 {{
      line-height: 1.3;
      margin-top: 1.6em;
      margin-bottom: 0.6em;
    }}
    h1 {{ font-size: 2.2rem; margin-top: 0; }}
    h2 {{ font-size: 1.65rem; border-bottom: 1px solid var(--border); padding-bottom: 0.25em; }}
    h3 {{ font-size: 1.3rem; }}
    p, li {{ font-size: 1rem; }}
    code {{
      font-family: "Cascadia Code", "Consolas", monospace;
      background: #eef2ff;
      color: #1e3a8a;
      padding: 0.15em 0.4em;
      border-radius: 6px;
      font-size: 0.95em;
    }}
    pre {{
      background: var(--code-bg);
      color: var(--code-text);
      padding: 18px 20px;
      border-radius: 12px;
      overflow-x: auto;
      border: 1px solid rgba(255, 255, 255, 0.08);
    }}
    pre code {{
      background: transparent;
      color: inherit;
      padding: 0;
    }}
    blockquote {{
      margin: 1.2em 0;
      padding: 0.8em 1.2em;
      border-left: 4px solid var(--accent);
      background: #eff5ff;
      color: var(--muted);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin: 1.5em 0;
      overflow: hidden;
      border-radius: 12px;
      border: 1px solid var(--border);
    }}
    th, td {{
      padding: 12px 14px;
      border-bottom: 1px solid var(--border);
      text-align: left;
      vertical-align: top;
    }}
    th {{
      background: #f7f9fc;
      font-weight: 700;
    }}
    img {{
      display: block;
      max-width: 100%;
      height: auto;
      margin: 18px auto;
      border: 1px solid var(--border);
      border-radius: 14px;
      box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08);
    }}
    a {{ color: var(--accent); }}
    hr {{ border: 0; border-top: 1px solid var(--border); margin: 2em 0; }}
    @media (max-width: 900px) {{
      main {{ margin: 16px; padding: 24px 18px; }}
      h1 {{ font-size: 1.8rem; }}
    }}
  </style>
</head>
<body>
  <main>
{body_html}
  </main>
</body>
</html>
'''


def convert_markdown_document(
    input_path: str | Path,
    output_arg: str | None = None,
    asset_dir_arg: str | None = None,
    asset_path_arg: str | None = None,
    title_arg: str | None = None,
) -> ConversionResult:
    resolved_input_path = Path(input_path).resolve()
    if not resolved_input_path.exists():
        raise FileNotFoundError(f'Input markdown file not found: {resolved_input_path}')

    output_path = resolve_output_path(resolved_input_path, output_arg)
    asset_dir = resolve_asset_dir(output_path, asset_dir_arg)
    asset_path_prefix = resolve_asset_path_prefix(output_path, asset_dir, asset_path_arg, asset_dir_arg)

    markdown_text = resolved_input_path.read_text(encoding='utf-8')
    title = title_arg.strip() if title_arg and title_arg.strip() else extract_title(markdown_text, resolved_input_path.stem)

    localized_markdown, warnings = localize_images(markdown_text, resolved_input_path.parent, asset_dir, asset_path_prefix)
    html_body = render_markdown(localized_markdown)
    html_document = wrap_html_document(title, html_body)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_document, encoding='utf-8')

    return ConversionResult(
        input_path=resolved_input_path,
        output_path=output_path,
        asset_dir=asset_dir,
        asset_path_prefix=asset_path_prefix,
        title=title,
        warnings=warnings,
    )


def print_conversion_summary(result: ConversionResult) -> None:
    print(f'HTML written to: {result.output_path}')
    print(f'Image assets stored in: {result.asset_dir}')
    print(f'Image paths used in HTML: {result.asset_path_prefix or "."}')

    if result.warnings:
        print('\nWarnings:', file=sys.stderr)
        for item in result.warnings:
            print(f'- {item}', file=sys.stderr)


class MarkdownExporterApp:
    def __init__(self, root: Any, initial_args: argparse.Namespace | None = None) -> None:
        self.root = root
        self.root.title('Markdown to HTML Exporter')
        self.root.geometry('920x680')
        self.root.minsize(820, 560)

        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.controls: list[Any] = []

        self.input_var = tk.StringVar(value=(getattr(initial_args, 'input', '') or ''))
        self.output_var = tk.StringVar(value=(getattr(initial_args, 'output', '') or ''))
        self.asset_dir_var = tk.StringVar(value=(getattr(initial_args, 'asset_dir', '') or ''))
        self.asset_path_var = tk.StringVar(value=(getattr(initial_args, 'asset_path', '') or ''))
        self.title_var = tk.StringVar(value=(getattr(initial_args, 'title', '') or ''))
        self._asset_path_syncing = False
        self._last_auto_asset_path = ''
        self._asset_path_is_auto = not self.asset_path_var.get().strip()
        self.asset_path_var.trace_add('write', self._on_asset_path_var_changed)
        self.status_var = tk.StringVar(value='Select a Markdown file to start.')

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

        ttk.Label(main_frame, text='Markdown to HTML Exporter', style='Title.TLabel').grid(
            row=0,
            column=0,
            sticky='w',
        )
        ttk.Label(
            main_frame,
            text='Leave Output, Asset Directory, or Asset Path blank to let the script derive sensible defaults.',
            wraplength=820,
        ).grid(row=1, column=0, sticky='w', pady=(6, 14))

        form_frame = ttk.Frame(main_frame)
        form_frame.grid(row=2, column=0, sticky='nsew')
        form_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=0)

        self._add_path_row(form_frame, 0, 'Markdown Input', self.input_var, self._browse_input_file)
        self._add_path_row(form_frame, 1, 'HTML Output', self.output_var, self._browse_output_file)
        self._add_path_row(form_frame, 2, 'Asset Directory', self.asset_dir_var, self._browse_asset_directory)
        self._add_entry_row(form_frame, 3, 'HTML Asset Path', self.asset_path_var)
        self._add_entry_row(form_frame, 4, 'Document Title', self.title_var)

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


    def _on_asset_path_var_changed(self, *_args) -> None:
        if self._asset_path_syncing:
            return

        current_value = self.asset_path_var.get().strip()
        self._asset_path_is_auto = current_value in {'', self._last_auto_asset_path}


    def _set_auto_asset_path(self, value: str) -> None:
        self._asset_path_syncing = True
        try:
            self.asset_path_var.set(value)
        finally:
            self._asset_path_syncing = False
        self._last_auto_asset_path = value
        self._asset_path_is_auto = True


    def _browse_input_file(self) -> None:
        selected = filedialog.askopenfilename(
            title='Select Markdown File',
            filetypes=[('Markdown Files', '*.md *.markdown *.mdown *.mkd'), ('All Files', '*.*')],
        )
        if selected:
            self.input_var.set(selected)
            self._apply_default_paths()


    def _browse_output_file(self) -> None:
        initial_name = ''
        input_value = self.input_var.get().strip()
        if input_value:
            initial_name = Path(input_value).with_suffix('.html').name
        selected = filedialog.asksaveasfilename(
            title='Select HTML Output',
            defaultextension='.html',
            initialfile=initial_name,
            filetypes=[('HTML Files', '*.html'), ('All Files', '*.*')],
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
        asset_path_arg = asset_path_value or None
        if force or not asset_path_value or self._asset_path_is_auto:
            asset_path_arg = None
        asset_path = resolve_asset_path_prefix(output_path, asset_dir, asset_path_arg, asset_dir_value or None)
        if force or not asset_path_value or self._asset_path_is_auto:
            self._set_auto_asset_path(asset_path)


    def _start_conversion(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        input_value = self.input_var.get().strip()
        if not input_value:
            messagebox.showerror('Missing Input', 'Please select a Markdown file first.')
            return

        self._apply_default_paths(force=False)
        self._set_busy(True)
        self.status_var.set('Converting...')
        self._append_log('Starting conversion...')

        worker_kwargs = {
            'input_path': input_value,
            'output_arg': self.output_var.get().strip() or None,
            'asset_dir_arg': self.asset_dir_var.get().strip() or None,
            'asset_path_arg': self.asset_path_var.get().strip() or None,
            'title_arg': self.title_var.get().strip() or None,
        }

        self.worker = threading.Thread(target=self._run_conversion_worker, kwargs=worker_kwargs, daemon=True)
        self.worker.start()


    def _run_conversion_worker(self, **kwargs) -> None:
        try:
            result = convert_markdown_document(**kwargs)
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
        if not self.title_var.get().strip():
            self.title_var.set(result.title)

        self._append_log(f'HTML written to: {result.output_path}')
        self._append_log(f'Image assets stored in: {result.asset_dir}')
        self._append_log(f'Image paths used in HTML: {result.asset_path_prefix or "."}')
        if result.warnings:
            for warning in result.warnings:
                self._append_log(f'Warning: {warning}')
        else:
            self._append_log('Finished without warnings.')

        self.status_var.set('Conversion completed.')
        self._set_busy(False)
        messagebox.showinfo('Done', f'HTML generated successfully:\n{result.output_path}')


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

    MarkdownExporterApp(root, initial_args)
    root.mainloop()
    return 0


def main() -> int:
    args = parse_args()

    if args.gui or not args.input:
        return launch_gui(args)

    result = convert_markdown_document(
        input_path=args.input,
        output_arg=args.output,
        asset_dir_arg=args.asset_dir,
        asset_path_arg=args.asset_path,
        title_arg=args.title,
    )
    print_conversion_summary(result)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
