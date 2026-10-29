"""Validation of files, references, hierarchy, consumers, UUIDs, and output."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Set

import pandas as pd

from .config import (
    UUID_FIELDS,
    MANDATORY_TARGET_FIELDS,
    METER_MAKE_REFERENCE,
    REFERENCE_CATEGORY_MAP,
    Category,
)
from .reference_lookup import ReferenceLookup
from .uuid_utils import is_valid_uuid, uuid_output_is_canonical

logger = logging.getLogger(__name__)

_SCIENTIFIC_IDENTIFIER = re.compile(r"^[+-]?\d+(?:\.\d+)?e[+-]?\d+$", re.IGNORECASE)


@dataclass
class ValidationIssue:
    severity: str
    category: str
    field: str = ""
    row: int = -1
    message: str = ""
    value: str = ""


@dataclass
class ValidationReport:
    issues: List[ValidationIssue] = field(default_factory=list)
    passed: bool = True

    @property
    def errors(self) -> List[ValidationIssue]:
        return [i for i in self.issues if i.severity == "ERROR"]

    @property
    def warnings(self) -> List[ValidationIssue]:
        return [i for i in self.issues if i.severity == "WARNING"]

    @property
    def error_count(self) -> int:
        return len(self.errors)

    @property
    def warning_count(self) -> int:
        return len(self.warnings)

    def add_error(self, category: str, message: str, field: str = "", row: int = -1, value: str = "") -> None:
        self.issues.append(ValidationIssue(
            severity="ERROR", category=category, field=field, row=row, message=message, value=value
        ))
        self.passed = False

    def add_warning(self, category: str, message: str, field: str = "", row: int = -1, value: str = "") -> None:
        self.issues.append(ValidationIssue(
            severity="WARNING", category=category, field=field, row=row, message=message, value=value
        ))

    def add_info(self, category: str, message: str, field: str = "", row: int = -1) -> None:
        self.issues.append(ValidationIssue(
            severity="INFO", category=category, field=field, row=row, message=message
        ))

    def merge(self, other: "ValidationReport") -> None:
        self.issues.extend(other.issues)
        if not other.passed:
            self.passed = False


def validate_reference_categories(
    ref_lookup: ReferenceLookup,
    dest_category: Category,
) -> ValidationReport:
    report = ValidationReport()
    for target_field, ref_cat in REFERENCE_CATEGORY_MAP.items():
        if not ref_lookup.category_exists(ref_cat):
            severity = report.add_warning if target_field in {
                "payment_mode", "metering_mode", "connection_category"
            } else report.add_warning
            severity(
                "reference",
                f"Reference category '{ref_cat}' for field '{target_field}' not found",
                field=target_field,
            )
    meter_make_ref = METER_MAKE_REFERENCE.get(dest_category)
    if meter_make_ref and not ref_lookup.category_exists(meter_make_ref):
        report.add_error(
            "reference",
            f"Meter make reference '{meter_make_ref}' for {dest_category.value} not found",
            field="meter_make",
        )
    if not ref_lookup.category_exists("CONSUMER IDENTIFICATION"):
        report.add_error(
            "reference",
            "Required reference category 'CONSUMER IDENTIFICATION' not found",
            field="consumer_identy",
        )
    for dup in ref_lookup.get_duplicate_report():
        report.add_warning("reference", f"Duplicate reference key: {dup}")
    return report


def validate_source_data(
    source_df: pd.DataFrame,
    required_source_columns: Set[str],
    row_count: int | None = None,
    allow_empty_frame: bool = False,
) -> ValidationReport:
    report = ValidationReport()
    record_count = len(source_df) if row_count is None else int(row_count)
    if source_df.empty and record_count == 0 and not allow_empty_frame:
        report.add_error("consumer", "Source data is empty")
        return report
    source_cols_lower = {c.lower().strip() for c in source_df.columns}
    for req_col in required_source_columns:
        if req_col.lower().strip() not in source_cols_lower:
            report.add_warning(
                "consumer",
                f"Required source column '{req_col}' not found",
                field=req_col,
            )
    if record_count:
        report.add_info("consumer", f"Source data has {record_count} records")
    return report


def validate_consumer_identifiers(
    source_df: pd.DataFrame,
    key_column: str,
) -> ValidationReport:
    """Flag identifiers already represented as scientific notation in the source."""
    report = ValidationReport()
    if key_column not in source_df.columns:
        return report
    matches = []
    for row_index, value in zip(source_df.index, source_df[key_column].tolist()):
        text = str(value or "").strip()
        if _SCIENTIFIC_IDENTIFIER.fullmatch(text):
            matches.append((int(row_index) + 2, text))
    if matches:
        report.add_warning(
            "consumer",
            f"{len(matches)} consumer identifier(s) are already written in scientific notation. "
            "The app will preserve the supplied text, but cannot recover digits lost before upload.",
            field=key_column,
        )
        for row_number, value in matches[:100]:
            report.add_warning(
                "consumer",
                "Consumer identifier is in scientific notation; check the original source value.",
                field=key_column,
                row=row_number,
                value=value,
            )
    return report


def validate_output_record(
    record: Dict[str, str],
    row_num: int,
    target_headers: List[str],
) -> ValidationReport:
    report = ValidationReport()
    for mf in MANDATORY_TARGET_FIELDS:
        value = record.get(mf, "")
        if not value or (isinstance(value, str) and not value.strip()):
            report.add_warning(
                "output",
                f"Mandatory field '{mf}' is empty",
                field=mf,
                row=row_num,
            )
    for uf in UUID_FIELDS:
        value = record.get(uf, "")
        if value and isinstance(value, str) and value.strip():
            val = value.strip()
            if not uuid_output_is_canonical(val):
                if is_valid_uuid(val):
                    report.add_error(
                        "uuid",
                        f"UUID field '{uf}' is not in {{uuid}} format",
                        field=uf,
                        row=row_num,
                        value=val,
                    )
                else:
                    report.add_error(
                        "uuid",
                        f"UUID field '{uf}' contains invalid UUID",
                        field=uf,
                        row=row_num,
                        value=val,
                    )
    return report


def validate_output_dataset(
    output_records: List[Dict[str, str]],
    target_headers: List[str],
    business_key: str = "consumer_number",
) -> ValidationReport:
    report = ValidationReport()
    if not output_records:
        report.add_error("output", "No output records generated")
        return report
    report.add_info("output", f"Total output records: {len(output_records)}")
    for record in output_records:
        record_keys = list(record.keys())
        if record_keys != target_headers:
            missing_in_record = set(target_headers) - set(record_keys)
            extra_in_record = set(record_keys) - set(target_headers)
            if missing_in_record:
                report.add_warning("output", f"Missing target columns in output: {missing_in_record}")
            if extra_in_record:
                report.add_warning("output", f"Extra columns in output: {extra_in_record}")
            break
    seen_keys: Dict[str, int] = {}
    for idx, record in enumerate(output_records):
        key = str(record.get(business_key, "")).strip()
        if key:
            if key in seen_keys:
                report.add_error(
                    "output",
                    f"Duplicate business key '{key}' at rows {seen_keys[key]} and {idx}",
                    field=business_key,
                    row=idx,
                )
            else:
                seen_keys[key] = idx
        record_report = validate_output_record(record, idx, target_headers)
        report.merge(record_report)
    return report
