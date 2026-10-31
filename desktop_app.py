import logging
import queue
import re
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from html import escape
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import unquote, urlsplit
from urllib.request import url2pathname

from openpyxl import load_workbook


def get_application_paths():
    frozen = getattr(sys, "frozen", False)
    executable_dir = (
        Path(sys.executable).resolve().parent
        if frozen
        else Path(__file__).resolve().parent
    )
    bundle_dir = Path(
        getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    ).resolve()
    external_resources = executable_dir / "templates", executable_dir / "assets"
    resource_dir = (
        executable_dir
        if frozen and all(path.is_dir() for path in external_resources)
        else bundle_dir
    )
    data_dir = (
        executable_dir
        if frozen and (executable_dir / "data" / "students.xlsx").is_file()
        else bundle_dir
    )
    return executable_dir, resource_dir, data_dir


BASE_DIR, RESOURCE_DIR, DATA_DIR = get_application_paths()
DEFAULT_TEMPLATE_PATH = RESOURCE_DIR / "templates" / "certificate_template.html"
ASSETS_DIR = RESOURCE_DIR / "assets"
USER_DATA_DIR = Path.home() / "Documents" / "CertificateGenerator"
DEFAULT_EXCEL_PATH = (
    USER_DATA_DIR / "data" / "students.xlsx"
    if getattr(sys, "frozen", False)
    else DATA_DIR / "data" / "students.xlsx"
)
OUTPUT_DIR = USER_DATA_DIR / "output"
REQUIRED_COLUMNS = ("student_name", "course_name", "score", "issue_date")
IMAGE_COLUMNS = ("signature_path", "stamp_path")
PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")


def _format_cell_value(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _format_date_text(value):
    value = _format_cell_value(value)
    if not value:
        return value
    for date_format in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, date_format).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return value


class MissingColumnsError(ValueError):
    def __init__(self, missing_columns):
        self.missing_columns = tuple(missing_columns)
        super().__init__(
            "В Excel-таблице отсутствуют обязательные колонки: "
            + ", ".join(self.missing_columns)
        )


def validate_and_load_excel(file_path: Path):
    file_path = Path(file_path).expanduser()
    if file_path.suffix.lower() != ".xlsx":
        raise ValueError("Выберите Excel-файл с расширением .xlsx.")
    if not file_path.is_file():
        raise FileNotFoundError(f"Excel-файл не найден: {file_path}")

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    try:
        rows = workbook.active.iter_rows(values_only=True)
        headers_row = next(rows, None)
        if headers_row is None:
            raise ValueError("Excel-файл не содержит строки заголовков.")

        headers = [_format_cell_value(value) for value in headers_row]
        named_headers = [header for header in headers if header]
        if len(named_headers) != len(set(named_headers)):
            raise ValueError("В Excel-файле обнаружены повторяющиеся заголовки.")

        missing_columns = [
            column for column in REQUIRED_COLUMNS if column not in headers
        ]
        if missing_columns:
            raise MissingColumnsError(missing_columns)

        column_indexes = {
            header: index for index, header in enumerate(headers) if header
        }
        students = []
        for values_row in rows:
            values = [_format_cell_value(value) for value in values_row]
            if not any(values):
                continue
            student = {
                column: values[index] if index < len(values) else ""
                for column, index in column_indexes.items()
            }
            student["issue_date"] = _format_date_text(
                student.get("issue_date", "")
            )
            students.append(student)
    finally:
        workbook.close()

    if not students:
        raise ValueError("В Excel-файле нет строк с данными.")
    return students


def _resolve_asset_path(image_value, column):
    image_path = Path(image_value).expanduser()
    if not image_path.is_absolute():
        image_path = ASSETS_DIR / image_path
    image_path = image_path.resolve()
    assets_dir = ASSETS_DIR.resolve()
    try:
        image_path.relative_to(assets_dir)
    except ValueError as error:
        raise ValueError(
            f"Путь изображения из колонки {column} должен находиться "
            f"в папке ресурсов: {assets_dir}"
        ) from error
    if not image_path.is_file():
        raise FileNotFoundError(
            f"Изображение из колонки {column} не найдено: {image_path}"
        )
    return image_path


def render_certificate(template, student):
    for column in IMAGE_COLUMNS:
        image_value = student.get(column, "").strip()
        image_pattern = re.compile(
            r"<img\b(?=[^>]*\bsrc\s*=\s*['\"]\{\{\s*"
            + re.escape(column)
            + r"\s*\}\}['\"])[^>]*>",
            flags=re.IGNORECASE,
        )
        if not image_value:
            template = image_pattern.sub("", template)
            template = re.sub(
                r"\{\{\s*" + re.escape(column) + r"\s*\}\}",
                "",
                template,
            )
            continue
        image_uri = _resolve_asset_path(image_value, column).as_uri()
        template = re.sub(
            r"\{\{\s*" + re.escape(column) + r"\s*\}\}",
            escape(image_uri, quote=True),
            template,
        )

    def replace_placeholder(match):
        column = match.group(1).strip()
        value = student.get(column, "")
        return escape(str(value), quote=True)

    return PLACEHOLDER_PATTERN.sub(replace_placeholder, template)


def _safe_resource_fetcher(allowed_directories):
    from weasyprint.urls import URLFetcher

    directories = tuple(Path(path).resolve() for path in allowed_directories)
    fetcher = URLFetcher(allowed_protocols={"file", "data"})

    def fetch(url):
        parsed = urlsplit(url)
        if parsed.scheme.lower() == "data":
            return fetcher.fetch(url)
        if parsed.scheme.lower() != "file" or parsed.netloc not in ("", "localhost"):
            raise ValueError("Разрешены только локальные ресурсы HTML-шаблона.")

        resource_path = Path(url2pathname(unquote(parsed.path))).resolve()
        if not resource_path.is_file() or not any(
            resource_path.is_relative_to(directory) for directory in directories
        ):
            raise ValueError(
                f"Шаблону запрещён доступ к ресурсу: {resource_path}"
            )
        return fetcher.fetch(url)

    return fetch


def _safe_filename(value):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    name = re.sub(r"\s+", "_", name)
    return (name[:100] or "student")


def generate_certificates(
    excel_path,
    template_path,
    output_dir=OUTPUT_DIR,
    progress_callback=None,
):
    from weasyprint import HTML

    excel_path = Path(excel_path).expanduser().resolve()
    template_path = Path(template_path).expanduser().resolve()
    output_dir = Path(output_dir).expanduser()
    if not template_path.is_file():
        raise FileNotFoundError(f"HTML-шаблон не найден: {template_path}")

    students = validate_and_load_excel(excel_path)
    template = template_path.read_text(encoding="utf-8")
    placeholders = {
        match.group(1).strip()
        for match in PLACEHOLDER_PATTERN.finditer(template)
    }
    image_placeholders = set(IMAGE_COLUMNS)
    missing_placeholders = sorted(
        placeholders - set(students[0]) - image_placeholders
    )
    if missing_placeholders:
        raise ValueError(
            "В Excel-файле отсутствуют колонки для плейсхолдеров шаблона: "
            + ", ".join(missing_placeholders)
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    resource_directories = (
        RESOURCE_DIR,
        ASSETS_DIR,
        template_path.parent,
        excel_path.parent,
    )
    generated_paths = []
    total = len(students)
    for current, student in enumerate(students, start=1):
        student_name = student.get("student_name", "") or f"student_{current}"
        filename = f"certificate_{current:03d}_{_safe_filename(student_name)}.pdf"
        output_path = output_dir / filename
        suffix = 1
        while output_path.exists():
            output_path = output_dir / (
                f"certificate_{current:03d}_"
                f"{_safe_filename(student_name)}_{suffix}.pdf"
            )
            suffix += 1

        rendered_html = render_certificate(template, student)
        try:
            HTML(
                string=rendered_html,
                base_url=str(template_path.parent),
                url_fetcher=_safe_resource_fetcher(resource_directories),
            ).write_pdf(target=str(output_path))
        except Exception as error:
            raise RuntimeError(
                f"Не удалось создать PDF для {student_name}: {error}"
            ) from error

        generated_paths.append(output_path)
        if progress_callback is not None:
            progress_callback(
                current,
                total,
                f"Создан документ {current} из {total}: {student_name}",
            )
    return generated_paths


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор PDF-сертификатов")
        self.geometry("600x320")
        self.resizable(False, False)
        self.configure(bg="#f4f6fb")

        self._messages = queue.Queue()
        self._worker = None
        self.excel_path = tk.StringVar(
            value=str(DEFAULT_EXCEL_PATH)
            if DEFAULT_EXCEL_PATH.is_file()
            else ""
        )
        self.template_path = tk.StringVar(
            value=str(DEFAULT_TEMPLATE_PATH)
            if DEFAULT_TEMPLATE_PATH.is_file()
            else ""
        )
        self.lbl_status = tk.StringVar(value="Выберите Excel-файл и HTML-шаблон.")

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("App.TFrame", background="#f4f6fb")
        style.configure(
            "Title.TLabel",
            background="#f4f6fb",
            foreground="#172554",
            font=("Segoe UI", 18, "bold"),
        )
        style.configure(
            "Status.TLabel",
            background="#f4f6fb",
            foreground="#475569",
            font=("Segoe UI", 9),
        )
        style.configure(
            "Run.TButton",
            padding=(12, 9),
            background="#3157d5",
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
        )

        frame = ttk.Frame(self, padding=24, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text="Пакетная генерация PDF",
            style="Title.TLabel",
        ).pack(anchor="w", pady=(0, 18))

        self._add_file_row(
            frame,
            "Excel-файл",
            self.excel_path,
            self._browse_excel,
        )
        self._add_file_row(
            frame,
            "HTML-шаблон",
            self.template_path,
            self._browse_template,
        )

        self.btn_generate = ttk.Button(
            frame,
            text="Создать сертификаты",
            command=self._start_generation,
            style="Run.TButton",
        )
        self.btn_generate.pack(fill="x", pady=(6, 12))

        ttk.Label(
            frame,
            textvariable=self.lbl_status,
            style="Status.TLabel",
            wraplength=540,
        ).pack(anchor="w", pady=(0, 7))
        self.progress = ttk.Progressbar(
            frame,
            mode="determinate",
            maximum=100,
        )
        self.progress.pack(fill="x")

    @staticmethod
    def _add_file_row(parent, label, variable, command):
        row = ttk.Frame(parent, style="App.TFrame")
        row.pack(fill="x", pady=(0, 12))
        ttk.Label(row, text=label, width=13).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
        )
        ttk.Button(row, text="Обзор...", command=command).pack(side="right")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=(("Excel", "*.xlsx"),),
        )
        if path:
            self.excel_path.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Выберите HTML-шаблон",
            filetypes=(
                ("HTML", "*.html *.htm"),
                ("Все файлы", "*.*"),
            ),
        )
        if path:
            self.template_path.set(path)

    def _start_generation(self):
        if self._worker is not None and self._worker.is_alive():
            return

        excel_path = Path(self.excel_path.get()).expanduser()
        template_path = Path(self.template_path.get()).expanduser()
        if not excel_path.is_file():
            messagebox.showerror("Файл не найден", "Выберите существующий Excel-файл.")
            return
        if not template_path.is_file():
            messagebox.showerror(
                "Файл не найден",
                "Выберите существующий HTML-шаблон.",
            )
            return

        self.btn_generate.configure(state="disabled")
        self.progress.configure(value=0, maximum=100)
        self.lbl_status.set("Подготовка документов...")
        self._worker = threading.Thread(
            target=self._run_generation,
            args=(excel_path, template_path),
            name="certificate-generator",
            daemon=True,
        )
        self._worker.start()

    def _run_generation(self, excel_path, template_path):
        try:
            paths = generate_certificates(
                excel_path,
                template_path,
                output_dir=OUTPUT_DIR,
                progress_callback=lambda current, total, message: self._messages.put(
                    ("progress", current, total, message)
                ),
            )
        except Exception as error:
            logging.exception("PDF certificate generation failed")
            self._messages.put(("error", str(error)))
        else:
            self._messages.put(("success", paths))

    def _poll_queue(self):
        try:
            while True:
                message = self._messages.get_nowait()
                event = message[0]
                if event == "progress":
                    _, current, total, status = message
                    self.progress.configure(
                        maximum=max(total, 1),
                        value=current,
                    )
                    self.lbl_status.set(status)
                elif event == "error":
                    self._finish()
                    self.lbl_status.set("Не удалось создать документы.")
                    messagebox.showerror("Ошибка генерации", message[1])
                elif event == "success":
                    paths = message[1]
                    self._finish()
                    self.progress.configure(
                        maximum=max(len(paths), 1),
                        value=len(paths),
                    )
                    self.lbl_status.set(
                        f"Готово: создано документов — {len(paths)}."
                    )
                    messagebox.showinfo(
                        "Генерация завершена",
                        f"Создано PDF-файлов: {len(paths)}\n\n"
                        f"Папка: {OUTPUT_DIR}",
                    )
        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _finish(self):
        self.btn_generate.configure(state="normal")
        self._worker = None


if __name__ == "__main__":
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        filename=USER_DATA_DIR / "desktop_app.log",
        encoding="utf-8",
    )
    app = Application()
    app.mainloop()
