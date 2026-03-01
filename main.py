from __future__ import annotations

import argparse
import calendar
import datetime as dt
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional

try:
    from docx import Document
except ImportError as exc:  # pragma: no cover - import guard
    raise SystemExit(
        "Missing dependency 'python-docx'. Install with: pip install python-docx"
    ) from exc


DEFAULT_INVOICE_DIR = Path(
    r"C:\Users\ver0016\OneDrive - Hoppers Crossing Secondary College\Desktop\Study Work\Invoices"
)


@dataclass(frozen=True)
class InvoiceConfig:
    prefix: str
    weekdays: tuple[int, ...]


INVOICE_RULES: dict[str, InvoiceConfig] = {
    "AFLO": InvoiceConfig(prefix="AF", weekdays=(calendar.MONDAY,)),
    "Bensons": InvoiceConfig(
        prefix="BE", weekdays=(calendar.WEDNESDAY, calendar.FRIDAY)
    ),
    "Adeval": InvoiceConfig(prefix="AD", weekdays=(calendar.FRIDAY,)),
    "Rodpak": InvoiceConfig(prefix="RO", weekdays=(calendar.SUNDAY,)),
}

MONTH_TOKEN = "Feb"


def first_weekday_of_month(year: int, month: int, target_weekday: int) -> dt.date:
    for day in range(1, 8):
        candidate = dt.date(year, month, day)
        if candidate.weekday() == target_weekday:
            return candidate
    raise ValueError("Unable to compute first weekday of month")


def all_weekdays_in_month(year: int, month: int, weekdays: Iterable[int]) -> list[dt.date]:
    weekday_set = set(weekdays)
    dates: list[dt.date] = []
    _, days_in_month = calendar.monthrange(year, month)
    for day in range(1, days_in_month + 1):
        current = dt.date(year, month, day)
        if current.weekday() in weekday_set:
            dates.append(current)
    return dates


def add_months(date_value: dt.date, months: int) -> dt.date:
    month = date_value.month - 1 + months
    year = date_value.year + month // 12
    month = month % 12 + 1
    day = min(date_value.day, calendar.monthrange(year, month)[1])
    return dt.date(year, month, day)


def parse_money(value: str) -> float:
    cleaned = value.strip().replace("$", "").replace(",", "")
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def format_money(value: float) -> str:
    if value.is_integer():
        return f"${int(value):,}"
    return f"${value:,.2f}"


def replace_first_word_with_month(text: str, month_name: str) -> str:
    if not text.strip():
        return text
    parts = text.split(maxsplit=1)
    if len(parts) == 1:
        return month_name
    return f"{month_name} {parts[1]}"


def replace_text_in_runs(paragraph, old: str, new: str) -> bool:
    full_text = "".join(run.text for run in paragraph.runs)
    if old not in full_text:
        return False
    replaced = full_text.replace(old, new)

    cursor = 0
    for run in paragraph.runs:
        run_len = len(run.text)
        run.text = replaced[cursor : cursor + run_len]
        cursor += run_len

    if cursor < len(replaced):
        paragraph.runs[-1].text += replaced[cursor:]
    return True


def iter_all_paragraphs(document: Document):
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def find_customer_key(filename: str) -> Optional[str]:
    stem = Path(filename).stem
    for key in INVOICE_RULES:
        if stem.lower().startswith(key.lower()):
            return key
    return None


def get_due_date(invoice_date: dt.date) -> dt.date:
    return add_months(invoice_date, 1) + dt.timedelta(days=1)


def find_line_with_label(document: Document, label: str):
    lower_label = label.lower()
    for paragraph in iter_all_paragraphs(document):
        if lower_label in paragraph.text.lower():
            return paragraph
    return None


def update_invoice_number(document: Document) -> None:
    paragraph = find_line_with_label(document, "invoice no")
    if not paragraph:
        return

    match = re.search(r"([A-Za-z]+)(\d+)", paragraph.text)
    if not match:
        return

    prefix, number = match.group(1), match.group(2)
    incremented = f"{prefix}{int(number) + 1:0{len(number)}d}"
    old = match.group(0)
    replace_text_in_runs(paragraph, old, incremented)


def update_labelled_date(document: Document, label: str, new_date: dt.date) -> None:
    paragraph = find_line_with_label(document, label)
    if not paragraph:
        return
    new_text = new_date.strftime("%d/%m/%y")

    match = re.search(r"\d{1,2}/\d{1,2}/\d{2}", paragraph.text)
    if match:
        replace_text_in_runs(paragraph, match.group(0), new_text)
        return

    # Fallback if date is glued to label, e.g. Date:01/03/26
    base = paragraph.text
    pattern = re.compile(r"(:\s*)([^\s]+)$")
    m = pattern.search(base)
    if m:
        replace_text_in_runs(paragraph, m.group(2), new_text)


def update_description(document: Document, month_name: str) -> None:
    paragraph = find_line_with_label(document, "description")
    if not paragraph:
        return
    text = paragraph.text
    idx = text.lower().find("description")
    if idx == -1:
        return

    content_start = text.find(":", idx)
    if content_start == -1:
        return
    content_start += 1
    original_desc = text[content_start:].strip()
    updated_desc = replace_first_word_with_month(original_desc, month_name)
    if original_desc and original_desc != updated_desc:
        replace_text_in_runs(paragraph, original_desc, updated_desc)


def find_service_table(document: Document):
    for table in document.tables:
        if not table.rows:
            continue
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]
        if len(headers) >= 2 and "date" in headers[0] and "amount" in headers[1]:
            return table
    return None


def duplicate_row(table, row_idx: int):
    row = table.rows[row_idx]
    new_tr = row._tr.clone()  # pylint: disable=protected-access
    table._tbl.append(new_tr)  # pylint: disable=protected-access


def set_service_dates(table, service_dates: List[dt.date]) -> float:
    existing_data_rows = table.rows[1:]

    while len(existing_data_rows) < len(service_dates):
        duplicate_row(table, len(table.rows) - 1)
        existing_data_rows = table.rows[1:]

    while len(existing_data_rows) > len(service_dates):
        table._tbl.remove(existing_data_rows[-1]._tr)  # pylint: disable=protected-access
        existing_data_rows = table.rows[1:]

    total = 0.0
    for row, service_date in zip(existing_data_rows, service_dates):
        if len(row.cells) < 2:
            continue
        row.cells[0].text = service_date.strftime("%d/%m")
        total += parse_money(row.cells[1].text)

    return total


def update_gst_and_total(document: Document, subtotal: float) -> None:
    gst_value = 0.0
    total_value = subtotal + gst_value

    for table in document.tables:
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                label = cell.text.strip().lower()
                if label == "gst" and idx + 1 < len(row.cells):
                    row.cells[idx + 1].text = "$0"
                if label == "total" and idx + 1 < len(row.cells):
                    row.cells[idx + 1].text = format_money(total_value)


def convert_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    try:
        from docx2pdf import convert

        convert(str(docx_path), str(pdf_path))
        return
    except ImportError:
        pass

    # Windows-only fallback using Word COM automation if pywin32 is present.
    try:
        import win32com.client  # type: ignore

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        document = word.Documents.Open(str(docx_path))
        document.SaveAs(str(pdf_path), FileFormat=17)
        document.Close()
        word.Quit()
        return
    except Exception as exc:  # pragma: no cover - runtime env specific
        raise RuntimeError(
            "Could not convert DOCX to PDF. Install 'docx2pdf' or ensure Word COM automation works."
        ) from exc


def target_names_for_month(source_path: Path, month_abbrev: str) -> tuple[Path, Path]:
    new_stem = source_path.stem.replace(MONTH_TOKEN, month_abbrev)
    docx_target = source_path.with_name(f"{new_stem}.docx")
    pdf_target = source_path.with_name(f"{new_stem}.pdf")
    return docx_target, pdf_target


def resolve_invoice_files(base_dir: Path) -> list[Path]:
    requested = [
        "AFLO Feb.docx",
        "Bensons Feb.docx",
        "Adeval Feb.docx",
        "Rodpak Feb.docx",
        "Rodpak Feb.doxc",
    ]
    files: list[Path] = []
    for name in requested:
        candidate = base_dir / name
        if candidate.exists() and candidate.suffix.lower() == ".docx":
            files.append(candidate)
    # Deduplicate while preserving order.
    unique: list[Path] = []
    seen: set[str] = set()
    for file in files:
        key = file.resolve().as_posix()
        if key not in seen:
            unique.append(file)
            seen.add(key)
    return unique


def process_invoice(source_doc: Path, year: int, month: int, dry_run: bool = False) -> None:
    customer_key = find_customer_key(source_doc.name)
    if not customer_key:
        print(f"Skipping {source_doc.name}: unknown customer key")
        return

    config = INVOICE_RULES[customer_key]
    invoice_date = first_weekday_of_month(year, month, calendar.SUNDAY)
    due_date = get_due_date(invoice_date)
    service_dates = all_weekdays_in_month(year, month, config.weekdays)

    month_abbrev = dt.date(year, month, 1).strftime("%b")
    month_name = dt.date(year, month, 1).strftime("%B")

    docx_target, pdf_target = target_names_for_month(source_doc, month_abbrev)

    if dry_run:
        print(f"[DRY RUN] Would generate {docx_target.name} and {pdf_target.name}")
        return

    working_copy = source_doc.with_name(f"{source_doc.stem}__tmp_working.docx")
    shutil.copy2(source_doc, working_copy)

    document = Document(str(working_copy))
    update_labelled_date(document, "date", invoice_date)
    update_labelled_date(document, "due date", due_date)
    update_invoice_number(document)
    update_description(document, month_name)

    service_table = find_service_table(document)
    subtotal = 0.0
    if service_table:
        subtotal = set_service_dates(service_table, service_dates)
    update_gst_and_total(document, subtotal)

    document.save(str(docx_target))
    convert_to_pdf(docx_target, pdf_target)

    working_copy.unlink(missing_ok=True)
    print(f"Generated: {docx_target.name} and {pdf_target.name}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate monthly invoice DOCX and PDF files from February templates."
    )
    parser.add_argument(
        "--invoice-dir",
        type=Path,
        default=DEFAULT_INVOICE_DIR,
        help="Directory containing source invoice DOCX files.",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=dt.date.today().year,
        help="Target year (defaults to current year).",
    )
    parser.add_argument(
        "--month",
        type=int,
        default=dt.date.today().month,
        help="Target month number 1-12 (defaults to current month).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned output without writing files.",
    )

    args = parser.parse_args()

    if not 1 <= args.month <= 12:
        raise SystemExit("Month must be between 1 and 12")

    invoice_files = resolve_invoice_files(args.invoice_dir)
    if not invoice_files:
        raise SystemExit(
            f"No matching source DOCX files found in {args.invoice_dir}. "
            "Expected files such as 'AFLO Feb.docx'."
        )

    for source_doc in invoice_files:
        process_invoice(source_doc, args.year, args.month, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
