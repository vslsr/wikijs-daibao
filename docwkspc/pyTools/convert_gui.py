from __future__ import annotations

import importlib
import os
import queue
import sys
import tempfile
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any

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


GUI_POLL_INTERVAL_MS = 150
WORD_SUFFIXES = {'.doc', '.docx'}
MARKDOWN_SUFFIXES = {'.md', '.markdown', '.mdown', '.mkd'}
INPUT_TYPE_WORD = 'word'
INPUT_TYPE_MARKDOWN = 'markdown'
TARGET_FORMAT_MARKDOWN = 'markdown'
TARGET_FORMAT_HTML = 'html'
TARGET_LABELS = {
    TARGET_FORMAT_MARKDOWN: 'Markdown (.md)',
    TARGET_FORMAT_HTML: 'HTML (.html)',
}
TARGET_LABEL_TO_KEY = {label: key for key, label in TARGET_LABELS.items()}
INPUT_FILETYPES = [
    ('Supported Files', '*.docx *.doc *.md *.markdown *.mdown *.mkd'),
    ('Word Files', '*.docx *.doc'),
    ('Markdown Files', '*.md *.markdown *.mdown *.mkd'),
    ('All Files', '*.*'),
]


@dataclass(slots=True)
class ConversionSummary:
    input_type: str
    target_format: str
    output_path: Path
    asset_dir: Path
    asset_path_prefix: str
    warnings: list[str]
    log_lines: list[str]
    title: str | None = None


def get_target_format(target_label: str) -> str:
    try:
        return TARGET_LABEL_TO_KEY[target_label]
    except KeyError as exc:
        raise ValueError(f'Unsupported target format: {target_label}') from exc


def resolve_output_path(input_path: Path, output_arg: str | None, output_suffix: str) -> Path:
    if output_arg:
        output_path = Path(output_arg)
    else:
        output_path = input_path.with_suffix(output_suffix)

    if not output_path.suffix:
        output_path = output_path.with_suffix(output_suffix)

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


def finalize_markdown(markdown_text: str) -> str:
    normalized = markdown_text.replace('\r\n', '\n')
    if normalized and not normalized.endswith('\n'):
        normalized += '\n'
    return normalized


def load_converter_module(module_name: str):
    try:
        return importlib.import_module(module_name)
    except ImportError as exc:
        raise RuntimeError(f'Unable to import {module_name}. {exc}') from exc


def detect_input_type(input_path: Path) -> str:
    suffix = input_path.suffix.lower()
    if suffix in WORD_SUFFIXES:
        return INPUT_TYPE_WORD
    if suffix in MARKDOWN_SUFFIXES:
        return INPUT_TYPE_MARKDOWN
    raise ValueError('Input file must be a supported Markdown or Word document.')


def validate_conversion_route(input_type: str, target_format: str) -> None:
    if input_type in {INPUT_TYPE_WORD, INPUT_TYPE_MARKDOWN} and target_format in {
        TARGET_FORMAT_MARKDOWN,
        TARGET_FORMAT_HTML,
    }:
        return
    raise ValueError(f'Unsupported conversion route: {input_type} -> {target_format}')


def build_description(input_type: str | None, target_format: str) -> str:
    target_label = TARGET_LABELS[target_format]
    if input_type == INPUT_TYPE_WORD:
        if target_format == TARGET_FORMAT_MARKDOWN:
            return 'Use markitdown to convert Word documents to Markdown and save extracted images.'
        return 'Convert Word documents to HTML by first generating temporary Markdown, then rendering standalone HTML with copied assets.'
    if input_type == INPUT_TYPE_MARKDOWN:
        if target_format == TARGET_FORMAT_MARKDOWN:
            return 'Rewrite a Markdown file to a new Markdown output and materialize its image references into the selected asset directory.'
        return 'Convert Markdown into standalone HTML and materialize referenced images into an asset directory.'
    return f'Select a supported input file, then convert it to {target_label}.'


def output_suffix_for_target(target_format: str) -> str:
    if target_format == TARGET_FORMAT_MARKDOWN:
        return '.md'
    if target_format == TARGET_FORMAT_HTML:
        return '.html'
    raise ValueError(f'Unsupported target format: {target_format}')


def output_label_for_target(target_format: str) -> str:
    if target_format == TARGET_FORMAT_MARKDOWN:
        return 'Markdown Output'
    if target_format == TARGET_FORMAT_HTML:
        return 'HTML Output'
    raise ValueError(f'Unsupported target format: {target_format}')


def asset_path_label_for_target(target_format: str) -> str:
    if target_format == TARGET_FORMAT_MARKDOWN:
        return 'Markdown Asset Path'
    if target_format == TARGET_FORMAT_HTML:
        return 'HTML Asset Path'
    raise ValueError(f'Unsupported target format: {target_format}')


def output_filetypes_for_target(target_format: str) -> list[tuple[str, str]]:
    if target_format == TARGET_FORMAT_MARKDOWN:
        return [('Markdown Files', '*.md'), ('All Files', '*.*')]
    if target_format == TARGET_FORMAT_HTML:
        return [('HTML Files', '*.html'), ('All Files', '*.*')]
    raise ValueError(f'Unsupported target format: {target_format}')


def title_supported_for_target(target_format: str) -> bool:
    return target_format == TARGET_FORMAT_HTML


def convert_markdown_to_markdown(
    input_path: Path,
    output_arg: str | None = None,
    asset_dir_arg: str | None = None,
    asset_path_arg: str | None = None,
) -> ConversionSummary:
    md2html = load_converter_module('md2html')

    output_path = resolve_output_path(input_path, output_arg, '.md')
    asset_dir = resolve_asset_dir(output_path, asset_dir_arg)
    asset_path_prefix = resolve_asset_path_prefix(output_path, asset_dir, asset_path_arg, asset_dir_arg)

    markdown_text = input_path.read_text(encoding='utf-8')
    localized_markdown, warnings = md2html.localize_images(markdown_text, input_path.parent, asset_dir, asset_path_prefix)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(finalize_markdown(localized_markdown), encoding='utf-8')

    log_lines = [
        f'Markdown written to: {output_path}',
        f'Image paths used in Markdown: {asset_path_prefix or "."}',
    ]
    if asset_dir.exists() and any(asset_dir.iterdir()):
        log_lines.append(f'Image assets stored in: {asset_dir}')
    else:
        log_lines.append('No images were materialized.')

    return ConversionSummary(
        input_type=INPUT_TYPE_MARKDOWN,
        target_format=TARGET_FORMAT_MARKDOWN,
        output_path=output_path,
        asset_dir=asset_dir,
        asset_path_prefix=asset_path_prefix,
        warnings=warnings,
        log_lines=log_lines,
    )


def convert_word_to_html(
    input_path: Path,
    output_arg: str | None = None,
    asset_dir_arg: str | None = None,
    asset_path_arg: str | None = None,
    title_arg: str | None = None,
) -> ConversionSummary:
    doc2md = load_converter_module('doc2md')
    md2html = load_converter_module('md2html')

    with tempfile.TemporaryDirectory(prefix='convert-gui-') as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        intermediate_md = temp_dir / f'{input_path.stem}.md'
        intermediate_assets = temp_dir / f'{input_path.stem}_assets'
        intermediate_asset_path = intermediate_assets.name

        word_result = doc2md.convert_word_document(
            input_path=input_path,
            output_arg=str(intermediate_md),
            asset_dir_arg=str(intermediate_assets),
            asset_path_arg=intermediate_asset_path,
        )
        html_result = md2html.convert_markdown_document(
            input_path=intermediate_md,
            output_arg=output_arg,
            asset_dir_arg=asset_dir_arg,
            asset_path_arg=asset_path_arg,
            title_arg=title_arg,
        )

    log_lines = [
        f'HTML written to: {html_result.output_path}',
        f'Image assets stored in: {html_result.asset_dir}',
        f'Image paths used in HTML: {html_result.asset_path_prefix or "."}',
        f'Document title: {html_result.title}',
        'Word input was converted via an intermediate Markdown file.',
    ]
    if word_result.converted_legacy_doc:
        log_lines.append('Legacy .doc input was converted to a temporary .docx via Microsoft Word.')

    return ConversionSummary(
        input_type=INPUT_TYPE_WORD,
        target_format=TARGET_FORMAT_HTML,
        output_path=html_result.output_path,
        asset_dir=html_result.asset_dir,
        asset_path_prefix=html_result.asset_path_prefix,
        warnings=[*word_result.warnings, *html_result.warnings],
        log_lines=log_lines,
        title=html_result.title,
    )


def convert_selected_document(
    input_path: str | Path,
    target_format: str,
    output_arg: str | None = None,
    asset_dir_arg: str | None = None,
    asset_path_arg: str | None = None,
    title_arg: str | None = None,
) -> ConversionSummary:
    resolved_input_path = Path(input_path).resolve()
    if not resolved_input_path.exists():
        raise FileNotFoundError(f'Input file not found: {resolved_input_path}')

    input_type = detect_input_type(resolved_input_path)
    validate_conversion_route(input_type, target_format)

    if input_type == INPUT_TYPE_WORD and target_format == TARGET_FORMAT_MARKDOWN:
        doc2md = load_converter_module('doc2md')
        result = doc2md.convert_word_document(
            input_path=resolved_input_path,
            output_arg=output_arg,
            asset_dir_arg=asset_dir_arg,
            asset_path_arg=asset_path_arg,
        )
        log_lines = [
            f'Markdown written to: {result.output_path}',
            f'Image paths used in Markdown: {result.asset_path_prefix or "."}',
        ]
        if result.converted_legacy_doc:
            log_lines.append('Legacy .doc input was converted to a temporary .docx via Microsoft Word.')
        if result.image_count:
            log_lines.append(f'Extracted images stored in: {result.asset_dir}')
            log_lines.append(f'Extracted image count: {result.image_count}')
        else:
            log_lines.append('No images were materialized.')
        return ConversionSummary(
            input_type=input_type,
            target_format=target_format,
            output_path=result.output_path,
            asset_dir=result.asset_dir,
            asset_path_prefix=result.asset_path_prefix,
            warnings=result.warnings,
            log_lines=log_lines,
        )

    if input_type == INPUT_TYPE_WORD and target_format == TARGET_FORMAT_HTML:
        return convert_word_to_html(
            input_path=resolved_input_path,
            output_arg=output_arg,
            asset_dir_arg=asset_dir_arg,
            asset_path_arg=asset_path_arg,
            title_arg=title_arg,
        )

    if input_type == INPUT_TYPE_MARKDOWN and target_format == TARGET_FORMAT_MARKDOWN:
        return convert_markdown_to_markdown(
            input_path=resolved_input_path,
            output_arg=output_arg,
            asset_dir_arg=asset_dir_arg,
            asset_path_arg=asset_path_arg,
        )

    md2html = load_converter_module('md2html')
    result = md2html.convert_markdown_document(
        input_path=resolved_input_path,
        output_arg=output_arg,
        asset_dir_arg=asset_dir_arg,
        asset_path_arg=asset_path_arg,
        title_arg=title_arg,
    )
    log_lines = [
        f'HTML written to: {result.output_path}',
        f'Image assets stored in: {result.asset_dir}',
        f'Image paths used in HTML: {result.asset_path_prefix or "."}',
        f'Document title: {result.title}',
    ]
    return ConversionSummary(
        input_type=input_type,
        target_format=target_format,
        output_path=result.output_path,
        asset_dir=result.asset_dir,
        asset_path_prefix=result.asset_path_prefix,
        warnings=result.warnings,
        log_lines=log_lines,
        title=result.title,
    )


class ConverterHubApp:
    def __init__(self, root: Any) -> None:
        self.root = root
        self.root.title('Document Conversion Hub')
        self.root.geometry('960x720')
        self.root.minsize(860, 600)

        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.controls: list[Any] = []

        self.target_var = tk.StringVar(value=TARGET_LABELS[TARGET_FORMAT_MARKDOWN])
        self.mode_description_var = tk.StringVar(value=build_description(None, TARGET_FORMAT_MARKDOWN))
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.asset_dir_var = tk.StringVar()
        self.asset_path_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self._asset_path_syncing = False
        self._last_auto_asset_path = ''
        self._asset_path_is_auto = True
        self.asset_path_var.trace_add('write', self._on_asset_path_var_changed)
        self.input_label_var = tk.StringVar(value='Input File')
        self.output_label_var = tk.StringVar(value=output_label_for_target(TARGET_FORMAT_MARKDOWN))
        self.asset_dir_label_var = tk.StringVar(value='Asset Directory')
        self.asset_path_label_var = tk.StringVar(value=asset_path_label_for_target(TARGET_FORMAT_MARKDOWN))
        self.source_type_var = tk.StringVar(value='Source type: not selected')
        self.status_var = tk.StringVar(value='Select a Markdown or Word file, then choose the target format.')

        self._configure_style()
        self._build_ui()
        self._sync_ui_state(force_defaults=False)
        self.root.after(GUI_POLL_INTERVAL_MS, self._poll_worker)


    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass
        style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'))
        style.configure('Target.TCombobox', padding=4)


    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=18)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        ttk.Label(main_frame, text='Document Conversion Hub', style='Title.TLabel').grid(
            row=0,
            column=0,
            sticky='w',
        )
        ttk.Label(
            main_frame,
            text='Choose a Markdown or Word input file, then select whether the result should be Markdown or HTML. The tool will route through the existing converters and keep asset handling consistent.',
            wraplength=860,
        ).grid(row=1, column=0, sticky='w', pady=(6, 12))

        target_frame = ttk.Frame(main_frame)
        target_frame.grid(row=2, column=0, sticky='ew', pady=(0, 14))
        target_frame.columnconfigure(1, weight=1)

        ttk.Label(target_frame, text='Target Format').grid(row=0, column=0, sticky='w', padx=(0, 12))
        self.target_combo = ttk.Combobox(
            target_frame,
            textvariable=self.target_var,
            values=[TARGET_LABELS[key] for key in (TARGET_FORMAT_MARKDOWN, TARGET_FORMAT_HTML)],
            state='readonly',
            style='Target.TCombobox',
        )
        self.target_combo.grid(row=0, column=1, sticky='ew')
        self.target_combo.bind('<<ComboboxSelected>>', self._on_target_changed)
        self.controls.append(self.target_combo)

        ttk.Label(target_frame, textvariable=self.source_type_var).grid(row=1, column=0, columnspan=2, sticky='w', pady=(8, 0))

        self.mode_description_label = ttk.Label(
            target_frame,
            textvariable=self.mode_description_var,
            wraplength=760,
        )
        self.mode_description_label.grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky='w', pady=(0, 12))

        auto_button = ttk.Button(button_frame, text='Fill Default Paths', command=lambda: self._apply_default_paths(force=True))
        auto_button.grid(row=0, column=0, padx=(0, 10))
        self.controls.append(auto_button)

        convert_button = ttk.Button(button_frame, text='Convert', command=self._start_conversion)
        convert_button.grid(row=0, column=1)
        self.controls.append(convert_button)

        form_frame = ttk.Frame(main_frame)
        form_frame.grid(row=4, column=0, sticky='nsew')
        form_frame.columnconfigure(1, weight=1)

        self._add_path_row(form_frame, 0, self.input_label_var, self.input_var, self._browse_input_file)
        self._add_path_row(form_frame, 1, self.output_label_var, self.output_var, self._browse_output_file)
        self._add_path_row(form_frame, 2, self.asset_dir_label_var, self.asset_dir_var, self._browse_asset_directory)
        self._add_entry_row(form_frame, 3, self.asset_path_label_var, self.asset_path_var)

        self.title_label = ttk.Label(form_frame, text='Document Title')
        self.title_label.grid(row=4, column=0, sticky='w', pady=6, padx=(0, 12))
        self.title_entry = ttk.Entry(form_frame, textvariable=self.title_var)
        self.title_entry.grid(row=4, column=1, columnspan=2, sticky='ew', pady=6)

        self.log_widget = ScrolledText(main_frame, height=16, wrap='word', font=('Consolas', 10))
        self.log_widget.grid(row=5, column=0, sticky='nsew', pady=(0, 10))
        self.log_widget.configure(state='disabled')

        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=6, column=0, sticky='w')


    def _add_path_row(self, parent: Any, row: int, label_var: Any, value_var: Any, browse_command) -> None:
        ttk.Label(parent, textvariable=label_var).grid(row=row, column=0, sticky='w', pady=6, padx=(0, 12))
        ttk.Entry(parent, textvariable=value_var).grid(row=row, column=1, sticky='ew', pady=6)
        browse_button = ttk.Button(parent, text='Browse...', command=browse_command)
        browse_button.grid(row=row, column=2, sticky='ew', pady=6, padx=(10, 0))
        self.controls.append(browse_button)


    def _add_entry_row(self, parent: Any, row: int, label_var: Any, value_var: Any) -> None:
        ttk.Label(parent, textvariable=label_var).grid(row=row, column=0, sticky='w', pady=6, padx=(0, 12))
        ttk.Entry(parent, textvariable=value_var).grid(row=row, column=1, columnspan=2, sticky='ew', pady=6)


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


    def _current_target_format(self) -> str:
        return get_target_format(self.target_var.get())


    def _current_input_type(self) -> str | None:
        input_value = self.input_var.get().strip()
        if not input_value:
            return None
        try:
            return detect_input_type(Path(input_value))
        except ValueError:
            return None


    def _on_target_changed(self, _event: Any = None) -> None:
        self._sync_ui_state(force_defaults=True)


    def _sync_ui_state(self, force_defaults: bool) -> None:
        input_type = self._current_input_type()
        target_format = self._current_target_format()

        if input_type == INPUT_TYPE_WORD:
            self.source_type_var.set('Source type: Word document')
        elif input_type == INPUT_TYPE_MARKDOWN:
            self.source_type_var.set('Source type: Markdown document')
        elif self.input_var.get().strip():
            self.source_type_var.set('Source type: unsupported file')
        else:
            self.source_type_var.set('Source type: not selected')

        self.mode_description_var.set(build_description(input_type, target_format))
        self.output_label_var.set(output_label_for_target(target_format))
        self.asset_dir_label_var.set('Asset Directory')
        self.asset_path_label_var.set(asset_path_label_for_target(target_format))

        if title_supported_for_target(target_format):
            self.title_label.grid()
            self.title_entry.grid()
        else:
            self.title_label.grid_remove()
            self.title_entry.grid_remove()

        self._apply_default_paths(force=force_defaults)


    def _browse_input_file(self) -> None:
        selected = filedialog.askopenfilename(
            title='Select Input File',
            filetypes=INPUT_FILETYPES,
        )
        if selected:
            self.input_var.set(selected)
            self._sync_ui_state(force_defaults=True)


    def _browse_output_file(self) -> None:
        input_value = self.input_var.get().strip()
        initial_name = ''
        target_format = self._current_target_format()
        if input_value:
            initial_name = Path(input_value).with_suffix(output_suffix_for_target(target_format)).name
        selected = filedialog.asksaveasfilename(
            title=f'Select {output_label_for_target(target_format)}',
            defaultextension=output_suffix_for_target(target_format),
            initialfile=initial_name,
            filetypes=output_filetypes_for_target(target_format),
        )
        if selected:
            self.output_var.set(selected)
            self._apply_default_paths(force=False)


    def _browse_asset_directory(self) -> None:
        selected = filedialog.askdirectory(title='Select Asset Directory')
        if selected:
            self.asset_dir_var.set(selected)
            self._apply_default_paths(force=False)


    def _apply_default_paths(self, force: bool = False) -> None:
        input_value = self.input_var.get().strip()
        if not input_value:
            return

        try:
            input_path = Path(input_value).resolve()
        except (OSError, RuntimeError):
            return

        target_format = self._current_target_format()
        output_value = self.output_var.get().strip()
        output_path = resolve_output_path(input_path, output_value or None, output_suffix_for_target(target_format))
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
            messagebox.showerror('Missing Input', 'Please select an input file first.')
            return

        self._apply_default_paths(force=False)
        target_format = self._current_target_format()
        self._set_busy(True)
        self.status_var.set('Converting...')
        self._append_log(f'Starting conversion to: {TARGET_LABELS[target_format]}')

        worker_kwargs = {
            'input_path': input_value,
            'target_format': target_format,
            'output_arg': self.output_var.get().strip() or None,
            'asset_dir_arg': self.asset_dir_var.get().strip() or None,
            'asset_path_arg': self.asset_path_var.get().strip() or None,
            'title_arg': self.title_var.get().strip() or None,
        }
        self.worker = threading.Thread(target=self._run_conversion_worker, kwargs=worker_kwargs, daemon=True)
        self.worker.start()


    def _run_conversion_worker(self, **kwargs) -> None:
        try:
            result = convert_selected_document(**kwargs)
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


    def _handle_success(self, result: ConversionSummary) -> None:
        self.output_var.set(str(result.output_path))
        self.asset_dir_var.set(str(result.asset_dir))
        self.asset_path_var.set(result.asset_path_prefix)
        if result.title and not self.title_var.get().strip():
            self.title_var.set(result.title)

        for line in result.log_lines:
            self._append_log(line)
        if result.warnings:
            for warning in result.warnings:
                self._append_log(f'Warning: {warning}')
        else:
            self._append_log('Finished without warnings.')

        self.status_var.set('Conversion completed.')
        self._set_busy(False)
        messagebox.showinfo('Done', f'Output generated successfully:\n{result.output_path}')


    def _handle_error(self, trace_text: str) -> None:
        self._append_log('Conversion failed.')
        self._append_log(trace_text.rstrip())
        self.status_var.set('Conversion failed.')
        self._set_busy(False)
        messagebox.showerror('Conversion Failed', trace_text.rstrip().splitlines()[-1])


    def _set_busy(self, is_busy: bool) -> None:
        new_state = 'disabled' if is_busy else 'readonly'
        self.target_combo.configure(state=new_state)
        for control in self.controls:
            if control is self.target_combo:
                continue
            control.configure(state='disabled' if is_busy else 'normal')


    def _append_log(self, message: str) -> None:
        self.log_widget.configure(state='normal')
        self.log_widget.insert('end', f'{message}\n')
        self.log_widget.see('end')
        self.log_widget.configure(state='disabled')


def launch_gui() -> int:
    if tk is None or ttk is None or filedialog is None or messagebox is None or ScrolledText is None:
        print('Tkinter is not available in this Python environment.', file=sys.stderr)
        return 1

    try:
        root = tk.Tk()
    except tk.TclError as exc:
        print(f'Unable to start GUI: {exc}', file=sys.stderr)
        return 1

    ConverterHubApp(root)
    root.mainloop()
    return 0


def main() -> int:
    return launch_gui()


if __name__ == '__main__':
    raise SystemExit(main())
