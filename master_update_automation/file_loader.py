"""Load Excel and CSV files with header extraction and validation."""

from __future__ import annotations

import csv
import logging
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, Iterator, List, Optional, Sequence, Tuple, Union

import openpyxl
import pandas as pd

from .config import (
    BUSINESS_KEY_SOURCE_CANDIDATES,
    CSV_CHUNK_SIZE,
    CSV_EXTENSIONS,
    CSV_SAMPLE_BYTES,
    EXCEL_EXTENSIONS,
    HIERARCHY_EXTENSIONS,
    MAX_UPLOAD_SIZE_BYTES,
    MAX_UPLOAD_SIZE_LABEL,
    REFERENCE_SHEET_NAME,
    SOURCE_EXTENSIONS,
    TABLE_PREVIEW_ROWS,
    TEMPLATE_EXTENSIONS,
    UPLOAD_COPY_CHUNK_BYTES,
)

logger = logging.getLogger(__name__)

CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
CSV_DELIMITERS = (",", ";", "\t", "|")


class FileLoadError(ValueError):
    """User-facing file loading or validation failure."""


@dataclass
class TabularInspection:
    headers: List[str]
    row_count: int
    preview: pd.DataFrame
    encoding: str = ""
    delimiter: str = ""
    is_csv: bool = False
    sheet_name: Optional[str] = None
    duplicate_key_count: int = 0
    key_column: Optional[str] = None
    warnings: List[str] = field(default_factory=list)


def _business_key_column(headers: Sequence[str]) -> Optional[str]:
    lower_map = {str(h).strip().lower(): h for h in headers if h}
    for candidate in BUSINESS_KEY_SOURCE_CANDIDATES:
        if candidate in lower_map:
            return lower_map[candidate]
    return None


def is_csv_file(filepath: str) -> bool:
    return Path(filepath).suffix.lower() in CSV_EXTENSIONS


def is_excel_file(filepath: str) -> bool:
    return Path(filepath).suffix.lower() in EXCEL_EXTENSIONS


def format_bytes(size: Optional[int]) -> str:
    if size is None or size < 0:
        return "unknown size"
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    if size < 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    return f"{size / (1024 * 1024 * 1024):.2f} GB"


def validate_upload_size(filename: str, size: Optional[int]) -> None:
    if size is None:
        return
    if size > MAX_UPLOAD_SIZE_BYTES:
        raise FileLoadError(
            f'The file "{filename}" is {format_bytes(size)}, which exceeds the '
            f"maximum allowed size of {MAX_UPLOAD_SIZE_LABEL} per file. "
            "Split the file or reduce its size, then upload it again."
        )


def validate_extension(filepath: str, allowed: Sequence[str], role: str) -> None:
    ext = Path(filepath).suffix.lower()
    if ext not in {a.lower() for a in allowed}:
        allowed_text = ", ".join(allowed)
        raise FileLoadError(
            f'The file "{Path(filepath).name}" uses unsupported extension "{ext or "(none)"}" '
            f"for the {role}. Accepted formats: {allowed_text}."
        )


def ensure_writable_temp_space(needed_bytes: int, directory: Optional[str] = None) -> None:
    target = directory or tempfile.gettempdir()
    try:
        free = shutil.disk_usage(target).free
    except OSError as exc:
        raise FileLoadError(
            f"Could not check available disk space in '{target}'. {exc}"
        ) from exc
    if needed_bytes > 0 and free < needed_bytes:
        raise FileLoadError(
            f"Insufficient disk space to store the upload. "
            f"Need about {format_bytes(needed_bytes)} free in '{target}', "
            f"but only {format_bytes(free)} is available. Free disk space and try again."
        )


def remove_temp_file(path: Optional[str]) -> None:
    if not path:
        return
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        logger.warning("Could not remove temporary file: %s", path)


def save_upload_stream(
    stream: Union[BinaryIO, object],
    suffix: str,
    filename: str = "",
    size: Optional[int] = None,
) -> str:
    """Copy an uploaded file to a temporary path using a streaming buffer."""
    validate_upload_size(filename or "uploaded file", size)
    needed = (size or 0) + UPLOAD_COPY_CHUNK_BYTES
    ensure_writable_temp_space(needed)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        if hasattr(stream, "seek"):
            try:
                stream.seek(0)
            except Exception:
                pass
        shutil.copyfileobj(stream, tmp, length=UPLOAD_COPY_CHUNK_BYTES)
    except OSError as exc:
        tmp.close()
        remove_temp_file(tmp.name)
        lowered = str(exc).lower()
        if "space" in lowered or "no space" in lowered:
            raise FileLoadError(
                "Insufficient disk space while saving the uploaded file. "
                "Free disk space and upload it again."
            ) from exc
        raise FileLoadError(f"Could not save the uploaded file: {exc}") from exc
    finally:
        tmp.close()
    return tmp.name


def load_excel_workbook(filepath: str) -> openpyxl.Workbook:
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    if path.suffix.lower() not in EXCEL_EXTENSIONS:
        raise FileLoadError(f"Expected an Excel file (.xlsx or .xls), got: {path.suffix}")
    if path.stat().st_size == 0:
        raise FileLoadError(f'The Excel file "{path.name}" is empty.')
    try:
        wb = openpyxl.load_workbook(filepath, data_only=False)
        logger.info("Loaded workbook: %s with sheets: %s", filepath, wb.sheetnames)
        return wb
    except Exception as exc:
        raise FileLoadError(f"Failed to load Excel file '{filepath}': {exc}") from exc


def get_sheet_names(wb: openpyxl.Workbook) -> List[str]:
    return list(wb.sheetnames)


def get_sheet_headers(wb: openpyxl.Workbook, sheet_name: str) -> List[str]:
    ws = wb[sheet_name]
    headers: List[str] = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        headers.append(str(val).strip() if val is not None else "")
    while headers and headers[-1] == "":
        headers.pop()
    return headers


def get_type_hints(wb: openpyxl.Workbook, sheet_name: str) -> dict:
    ws = wb[sheet_name]
    headers = get_sheet_headers(wb, sheet_name)
    type_hints = {}
    for col_idx, header in enumerate(headers, 1):
        if header:
            val = ws.cell(row=2, column=col_idx).value
            type_hints[header] = str(val).strip() if val is not None else ""
    return type_hints


def read_excel_data(
    filepath: str,
    sheet_name: str,
    skip_type_hint_row: bool = False,
) -> pd.DataFrame:
    if skip_type_hint_row:
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=0,
            skiprows=[1],
            dtype=str,
            engine="openpyxl",
        )
    else:
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            header=0,
            dtype=str,
            engine="openpyxl",
        )
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    logger.info("Read %s rows from '%s' in '%s'", len(df), sheet_name, filepath)
    return df


def _read_sample_bytes(filepath: str, nbytes: int = CSV_SAMPLE_BYTES) -> bytes:
    with open(filepath, "rb") as handle:
        return handle.read(nbytes)


def _decode_sample(raw: bytes) -> Tuple[str, str]:
    last_error: Optional[Exception] = None
    for enc in CSV_ENCODINGS:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    raise FileLoadError(
        "The CSV encoding could not be determined. Save the file as UTF-8 CSV and upload it again."
    ) from last_error


def _detect_delimiter(sample_text: str) -> str:
    header_line = sample_text.splitlines()[0] if sample_text.strip() else ""
    header_line = header_line.replace("\ufeff", "")
    if not header_line.strip():
        raise FileLoadError("The CSV file has no header row.")
    try:
        dialect = csv.Sniffer().sniff(sample_text, delimiters="".join(CSV_DELIMITERS))
        if dialect.delimiter in CSV_DELIMITERS:
            return dialect.delimiter
    except csv.Error:
        pass
    counts = {delim: header_line.count(delim) for delim in CSV_DELIMITERS}
    best = max(counts, key=counts.get)
    if counts[best] == 0:
        if "," not in header_line and ";" not in header_line and "\t" not in header_line:
            return ","
        raise FileLoadError(
            "The CSV delimiter could not be detected. Save the file as a comma-separated "
            "UTF-8 CSV (or use semicolon, tab, or pipe consistently) and upload it again."
        )
    return best


def _normalise_csv_frame(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    if any(str(c).startswith("Unnamed") for c in df.columns) and all(
        str(c).startswith("Unnamed") for c in df.columns
    ):
        raise FileLoadError("The CSV file is missing a header row.")
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def read_csv_data(filepath: str, encoding: str = "utf-8-sig") -> pd.DataFrame:
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    if path.stat().st_size == 0:
        raise FileLoadError(f'The CSV file "{path.name}" is empty.')
    raw = _read_sample_bytes(filepath)
    if not raw.strip():
        raise FileLoadError(f'The CSV file "{path.name}" is empty.')
    sample_text, detected_encoding = _decode_sample(raw)
    encodings_to_try = [encoding, detected_encoding, *CSV_ENCODINGS]
    seen = []
    last_error: Optional[Exception] = None
    delimiter = _detect_delimiter(sample_text)
    for enc in encodings_to_try:
        if not enc or enc in seen:
            continue
        seen.append(enc)
        try:
            df = pd.read_csv(
                filepath,
                dtype=str,
                encoding=enc,
                sep=delimiter,
                keep_default_na=False,
                na_filter=False,
                quoting=csv.QUOTE_MINIMAL,
            )
            logger.info("Read CSV with encoding=%s delimiter=%r", enc, delimiter)
            df = _normalise_csv_frame(df)
            if len(df.columns) == 0:
                raise FileLoadError(f'The CSV file "{path.name}" is missing a header row.')
            logger.info("Read %s rows from CSV: %s", len(df), filepath)
            return df
        except FileLoadError:
            raise
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
        except Exception as exc:
            last_error = exc
            continue
    raise FileLoadError(
        f"Could not read CSV '{path.name}'. Save it as UTF-8 with a consistent delimiter and try again."
    ) from last_error


def _csv_read_settings(filepath: str) -> Tuple[str, str]:
    raw = _read_sample_bytes(filepath)
    if not raw.strip():
        raise FileLoadError(f'The CSV file "{Path(filepath).name}" is empty.')
    sample_text, encoding = _decode_sample(raw)
    return encoding, _detect_delimiter(sample_text)


def read_csv_headers(filepath: str) -> List[str]:
    """Read only the CSV header row, without loading consumer records."""
    encoding, delimiter = _csv_read_settings(filepath)
    columns = pd.read_csv(
        filepath,
        dtype=str,
        encoding=encoding,
        sep=delimiter,
        keep_default_na=False,
        na_filter=False,
        quoting=csv.QUOTE_MINIMAL,
        nrows=0,
    ).columns
    headers = [str(column).replace("\ufeff", "").strip() for column in columns]
    if not headers or all(not header for header in headers):
        raise FileLoadError(f'The CSV file "{Path(filepath).name}" is missing a header row.')
    return headers


def iter_csv_data_chunks(
    filepath: str,
    chunksize: int = CSV_CHUNK_SIZE,
) -> Iterator[pd.DataFrame]:
    """Yield normalized CSV data chunks to bound source-file memory use."""
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    if path.stat().st_size == 0:
        raise FileLoadError(f'The CSV file "{path.name}" is empty.')
    encoding, delimiter = _csv_read_settings(filepath)
    reader = pd.read_csv(
        filepath,
        dtype=str,
        encoding=encoding,
        sep=delimiter,
        keep_default_na=False,
        na_filter=False,
        quoting=csv.QUOTE_MINIMAL,
        chunksize=chunksize,
    )
    try:
        for chunk in reader:
            yield _normalise_csv_frame(chunk)
    finally:
        if hasattr(reader, "close"):
            reader.close()


def inspect_csv_file(filepath: str, chunksize: int = CSV_CHUNK_SIZE) -> TabularInspection:
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    if path.stat().st_size == 0:
        raise FileLoadError(f'The CSV file "{path.name}" is empty.')
    raw = _read_sample_bytes(filepath)
    if not raw.strip():
        raise FileLoadError(f'The CSV file "{path.name}" is empty.')
    sample_text, encoding = _decode_sample(raw)
    delimiter = _detect_delimiter(sample_text)
    reader = pd.read_csv(
        filepath,
        dtype=str,
        encoding=encoding,
        sep=delimiter,
        keep_default_na=False,
        na_filter=False,
        quoting=csv.QUOTE_MINIMAL,
        chunksize=chunksize,
    )
    headers: List[str] = []
    preview_rows: List[pd.DataFrame] = []
    row_count = 0
    key_col: Optional[str] = None
    try:
        for chunk in reader:
            chunk = _normalise_csv_frame(chunk)
            if not headers:
                headers = list(chunk.columns)
                key_col = _business_key_column(headers)
            row_count += len(chunk)
            remaining = TABLE_PREVIEW_ROWS - sum(len(p) for p in preview_rows)
            if remaining > 0:
                preview_rows.append(chunk.head(remaining))
            del chunk
    finally:
        if hasattr(reader, "close"):
            reader.close()
    if not headers:
        raise FileLoadError(f'The CSV file "{path.name}" is missing a header row.')
    preview = pd.concat(preview_rows, ignore_index=True) if preview_rows else pd.DataFrame(columns=headers)
    return TabularInspection(
        headers=headers,
        row_count=row_count,
        preview=preview,
        encoding=encoding,
        delimiter=delimiter,
        is_csv=True,
        duplicate_key_count=0,
        key_column=key_col,
    )


def inspect_excel_sheet(filepath: str, sheet_name: str) -> TabularInspection:
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise FileLoadError(f"Worksheet '{sheet_name}' was not found in '{Path(filepath).name}'.")
        ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows)
        except StopIteration:
            raise FileLoadError(f"Worksheet '{sheet_name}' is empty.")
        headers = [str(v).strip() if v is not None else "" for v in header_row]
        while headers and headers[-1] == "":
            headers.pop()
        if not any(headers):
            raise FileLoadError(f"Worksheet '{sheet_name}' is missing a header row.")
        preview_data: List[List[str]] = []
        row_count = 0
        key_col = _business_key_column(headers)
        for raw in rows:
            values = ["" if v is None else str(v).strip() for v in raw[:len(headers)]]
            if not any(values):
                continue
            row_count += 1
            if len(preview_data) < TABLE_PREVIEW_ROWS:
                preview_data.append(values)
        preview = pd.DataFrame(preview_data, columns=headers)
        return TabularInspection(
            headers=headers,
            row_count=row_count,
            preview=preview,
            is_csv=False,
            sheet_name=sheet_name,
            duplicate_key_count=0,
            key_column=key_col,
        )
    finally:
        wb.close()


def read_hierarchy_file(filepath: str) -> pd.DataFrame:
    path = Path(filepath)
    validate_extension(str(path), HIERARCHY_EXTENSIONS, "hierarchy report")
    ext = path.suffix.lower()
    if ext in CSV_EXTENSIONS:
        df = read_csv_data(filepath)
    elif ext in EXCEL_EXTENSIONS:
        wb = load_excel_workbook(filepath)
        sheets = get_sheet_names(wb)
        wb.close()
        if not sheets:
            raise FileLoadError("The hierarchy Excel file has no worksheets.")
        df = read_excel_data(filepath, sheets[0])
    else:
        raise FileLoadError(f"Unsupported hierarchy file format: {ext}")
    for col in df.columns:
        df[col] = df[col].map(lambda value: value.strip() if isinstance(value, str) else value)
    return df


def read_source_data(
    filepath: str,
    sheet_name: Optional[str] = None,
    skip_type_hint_row: bool = False,
) -> pd.DataFrame:
    path = Path(filepath)
    validate_extension(str(path), SOURCE_EXTENSIONS, "consumer source data")
    if is_csv_file(str(path)):
        return read_csv_data(filepath)
    if not sheet_name:
        raise FileLoadError("A worksheet name is required for Excel consumer source files.")
    return read_excel_data(filepath, sheet_name, skip_type_hint_row=skip_type_hint_row)


def load_tabular_file(
    filepath: str,
    file_role: str,
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    """Dispatch table loading by file role. The target template is never read as CSV."""
    role = (file_role or "").strip().lower()
    if role in {"template", "target"}:
        validate_extension(filepath, TEMPLATE_EXTENSIONS, "target template")
        raise FileLoadError(
            "The target template must be an Excel workbook (.xlsx or .xls) because it "
            "contains the output structure and the ref_project_masters_lov worksheet. "
            "CSV cannot replace the template."
        )
    if role == "hierarchy":
        return read_hierarchy_file(filepath)
    if role in {"source", "consumer"}:
        return read_source_data(filepath, sheet_name)
    raise FileLoadError(f"Unknown file role: {file_role}")


def detect_target_sheet(wb: openpyxl.Workbook) -> Optional[str]:
    """Return the consumer-master sheet, or None if the choice is ambiguous."""
    candidates: List[str] = []
    for name in wb.sheetnames:
        if name.lower() == REFERENCE_SHEET_NAME.lower():
            continue
        name_lower = name.lower()
        if "consumer" in name_lower or "master" in name_lower:
            candidates.append(name)
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) == 0:
        non_ref = [n for n in wb.sheetnames if n.lower() != REFERENCE_SHEET_NAME.lower()]
        if len(non_ref) == 1:
            return non_ref[0]
        return None
    return None


def detect_source_sheet(wb: openpyxl.Workbook) -> Optional[str]:
    if len(wb.sheetnames) == 1:
        return wb.sheetnames[0]
    return None


def validate_file_readable(filepath: str) -> bool:
    path = Path(filepath)
    return path.exists() and path.is_file()
