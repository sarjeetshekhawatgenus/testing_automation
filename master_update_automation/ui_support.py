"""UI helpers for workflow status, filenames, and user-facing messages.

This module contains presentation logic only. It does not perform lookups,
hierarchy matching, or other transformation rules.
"""

from __future__ import annotations

import html
import ctypes
import os
import re
from datetime import datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from .config import MAX_UPLOAD_SIZE_LABEL, OPERATIONS, Operation
from .excel_writer import generate_output_filename


WORKFLOW_STEPS: List[Tuple[int, str]] = [
    (1, "Template"),
    (2, "Operation"),
    (3, "Hierarchy"),
    (4, "Consumer Data"),
    (5, "Mappings"),
    (6, "Validate"),
    (7, "Export"),
]

UI_OPERATION_GROUPS: Dict[str, List[int]] = {
    "Category Conversion": [1, 2, 3, 4, 8, 9],
    "Same-Category Update": [5, 6, 7],
}

ENGINE_MANDATORY_SOURCE_TARGETS = {
    "consumer_number",
    "account_number",
    "consumer_name",
    "consumer_identy",
    "cons_connect_type",
    "is_old_mtr_available",
    "cons_available_mst",
}

# meter_category is populated from the template reference sheet (Main Meter).
ENGINE_POPULATED_TARGETS = {"meter_category"}

RESULT_KEYS = (
    "pipeline_result",
    "output_bytes",
    "output_name",
    "review_bytes",
    "review_name",
    "audit_bytes",
    "error_bytes",
    "audit_name",
    "error_name",
)

MAX_TABLE_PREVIEW_ROWS = 50

PROCESSING_STAGES = [
    "Reading input files",
    "Validating workbook structure",
    "Resolving reference UUIDs, matching hierarchy, and applying business rules",
    "Validating transformed data",
    "Preparing the audit report",
]


def file_fingerprint(name: str, size: Optional[int]) -> Tuple[str, int]:
    return (str(name or ""), int(size or 0))


def format_file_size(size: Optional[int]) -> str:
    if size is None or size < 0:
        return "Unknown size"
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    if size < 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} MB"
    return f"{size / (1024 * 1024 * 1024):.2f} GB"


def upload_limit_caption() -> str:
    return f"Maximum file size: {MAX_UPLOAD_SIZE_LABEL} per file"


def available_memory_bytes() -> Optional[int]:
    """Return currently available physical memory where the OS exposes it."""
    if os.name == "nt":
        class _MemoryStatus(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        try:
            status = _MemoryStatus()
            status.dwLength = ctypes.sizeof(status)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return int(status.ullAvailPhys)
        except (AttributeError, OSError, TypeError):
            pass
    try:
        page_size = os.sysconf("SC_PAGE_SIZE")
        available_pages = os.sysconf("SC_AVPHYS_PAGES")
        return int(page_size * available_pages)
    except (AttributeError, OSError, ValueError):
        return None


def estimate_peak_memory_bytes(
    source_size: Optional[int],
    source_is_csv: bool,
    hierarchy_size: Optional[int],
    template_size: Optional[int],
) -> int:
    """Conservative peak estimate for pandas records, indexes, audits, and Excel output."""
    source_multiplier = 30 if source_is_csv else 60
    return (
        max(0, int(source_size or 0)) * source_multiplier
        + max(0, int(hierarchy_size or 0)) * 12
        + max(0, int(template_size or 0)) * 5
        + 512 * 1024 * 1024
    )


def operation_explanation(op: Operation) -> str:
    if op.is_same_category:
        return (
            f"Same-category update for {op.source.value}. "
            "Lookups and transformations still run; identification is not converted."
        )
    return (
        f"Convert records from {op.source.value} to {op.destination.value}. "
        f"Meter-make lookups use the {op.destination.value} reference category."
    )


def download_basename(op: Operation, when: Optional[datetime] = None) -> str:
    stamp = (when or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return f"Consumer_Master_{op.source.value}_to_{op.destination.value}_{stamp}"


def output_workbook_name(op: Operation, when: Optional[datetime] = None) -> str:
    stamp = (when or datetime.now()).strftime("%Y%m%d_%H%M%S")
    path = generate_output_filename(
        dest_category=op.destination,
        source_category=op.source,
        timestamp=stamp,
        output_dir=".",
    )
    from pathlib import Path
    return Path(path).name


def report_filename(op: Operation, kind: str, when: Optional[datetime] = None) -> str:
    return f"{download_basename(op, when)}_{kind}.xlsx"


def results_status_label(result: Any) -> str:
    if result is None:
        return "Not run"
    failed = getattr(result.audit, "failed_count", 0)
    review = getattr(result.audit, "review_count", 0)
    warnings = getattr(result.audit, "warning_count", 0)
    total = getattr(result.audit, "total_records", 0)
    val_warnings = len(getattr(result.validation, "warnings", []) or [])
    if not getattr(result, "ready_for_output", False):
        if total and failed == total:
            return "Failed"
        return "Action required"
    if warnings or val_warnings or review:
        return "Complete with warnings"
    return "Complete"


def action_counts(entries: Iterable[Any]) -> Dict[str, int]:
    inserted = 0
    updated = 0
    for entry in entries:
        action = str(getattr(entry, "action", "") or "").upper()
        status = str(getattr(entry, "status", "") or "").upper()
        if status in {"FAILED", "REQUIRES_REVIEW"}:
            continue
        if action == "INSERT":
            inserted += 1
        elif action == "UPDATE":
            updated += 1
    return {"inserted": inserted, "updated": updated}


def compute_step_statuses(state: Mapping[str, Any], current_step: int) -> List[Dict[str, str]]:
    statuses: List[Dict[str, str]] = []
    checks = {
        1: _template_status(state),
        2: "complete" if state.get("operation_id") else "pending",
        3: _hierarchy_status(state),
        4: _source_status(state),
        5: _mapping_status(state),
        6: _validate_status(state),
        7: _export_status(state),
    }
    for step, label in WORKFLOW_STEPS:
        status = checks[step]
        if step == current_step and status not in {"complete", "attention"}:
            status = "current"
        elif step == current_step and status == "pending":
            status = "current"
        statuses.append({"step": str(step), "label": label, "status": status})
    return statuses


def _template_status(state: Mapping[str, Any]) -> str:
    if not state.get("template_path"):
        return "pending"
    if state.get("template_error") or not state.get("ref_sheet_ok") or not state.get("target_sheet"):
        return "attention"
    return "complete"


def _hierarchy_status(state: Mapping[str, Any]) -> str:
    if not state.get("hierarchy_path"):
        return "pending"
    if state.get("hierarchy_error"):
        return "attention"
    summary = state.get("hierarchy_summary") or {}
    if summary.get("duplicates") or summary.get("missing_levels"):
        return "attention"
    return "complete"


def _source_status(state: Mapping[str, Any]) -> str:
    if not state.get("source_path"):
        return "pending"
    if not state.get("source_is_csv") and not state.get("source_sheet"):
        return "pending"
    if state.get("source_error") or not state.get("source_key_column"):
        return "attention"
    return "complete"


def _mapping_status(state: Mapping[str, Any]) -> str:
    if not state.get("source_path"):
        return "pending"
    if state.get("mapping_blockers"):
        return "attention"
    if state.get("mapping_confirmed"):
        return "complete"
    return "pending"


def _validate_status(state: Mapping[str, Any]) -> str:
    result = state.get("pipeline_result")
    if result is None:
        return "pending"
    if not getattr(result, "ready_for_output", False):
        return "attention"
    return "complete"


def _export_status(state: Mapping[str, Any]) -> str:
    if state.get("output_bytes"):
        return "complete"
    if state.get("pipeline_result") is None:
        return "pending"
    return "pending"


def invalidate_processing_results(state: Any) -> None:
    for key in RESULT_KEYS:
        state[key] = None


def stepper_html(statuses: Sequence[Mapping[str, str]]) -> str:
    parts = ['<div class="mda-stepper-container" aria-label="Workflow progress"><ol class="mda-stepper">']
    for idx, item in enumerate(statuses):
        label = html.escape(item["label"])
        status = html.escape(item["status"])
        step = html.escape(item["step"])
        step_num = step.zfill(2)
        icon = "✓" if status == "complete" else ("!" if status == "attention" else step_num)
        
        tag_html = ""
        if status == "current":
            tag_html = '<span class="mda-step-tag mda-tag-active">Active</span>'
        elif status == "complete":
            tag_html = '<span class="mda-step-tag mda-tag-done">Done</span>'
        elif status == "attention":
            tag_html = '<span class="mda-step-tag mda-tag-attn">Review</span>'

        parts.append(
            f'<li class="mda-step mda-step-{status}">'
            f'<span class="mda-step-badge">{icon}</span>'
            f'<div class="mda-step-info">'
            f'<span class="mda-step-label">{label}</span>'
            f'{tag_html}'
            f'</div>'
            f'<span class="mda-step-state" style="display:none">{status.replace("attention", "needs attention")}</span>'
            f'</li>'
        )
        if idx < len(statuses) - 1:
            conn_class = "mda-conn-active" if status == "complete" else "mda-conn-pending"
            parts.append(f'<li class="mda-connector {conn_class}" aria-hidden="true"></li>')
    parts.append("</ol></div>")
    return "".join(parts)




def humanize_exception(exc: BaseException, context: str) -> Tuple[str, str]:
    detail = str(exc).strip() or exc.__class__.__name__
    lowered = detail.lower()
    file_label = {
        "template": "target template",
        "hierarchy": "hierarchy report",
        "source": "consumer source file",
        "output": "output workbook",
    }.get(context, "uploaded file")

    if "ref_project_masters_lov" in lowered:
        return (
            "The uploaded target template is missing the required reference worksheet "
            "`ref_project_masters_lov`. This sheet supplies UUIDs such as Main Meter. "
            "Please upload a template that includes that worksheet.",
            detail,
        )
    if "failed to load excel" in lowered or "not a zip file" in lowered:
        return (
            f"The {file_label} could not be opened as a valid Excel workbook. "
            "Confirm the file is not password-protected or corrupted, then upload it again.",
            detail,
        )
    if "unsupported hierarchy" in lowered:
        return (
            "The hierarchy report must be a CSV or Excel file. "
            "Please export the report in one of those formats and upload it again.",
            detail,
        )
    if "exceeds the maximum allowed size" in lowered or "maximum allowed size" in lowered:
        return (detail, detail)
    if "insufficient disk space" in lowered or "no space" in lowered:
        return (
            f"There is not enough disk space to store or process the {file_label}. "
            "Free disk space and try again.",
            detail,
        )
    if "not enough memory" in lowered or "memory error" in lowered:
        return (
            f"The {file_label} is too large for the available memory. "
            "Reduce the file size or run the application on a machine with more RAM.",
            detail,
        )
    if "delimiter" in lowered:
        return (
            "The CSV delimiter could not be detected. Save the file as a comma-separated "
            "UTF-8 CSV (semicolon, tab, and pipe are also supported) and upload it again.",
            detail,
        )
    if "encoding" in lowered:
        return (
            "The CSV encoding could not be determined. "
            "Save the file as UTF-8 CSV and upload it again.",
            detail,
        )
    if "empty" in lowered:
        return (
            f"The {file_label} is empty. Upload a file that contains a header row and data.",
            detail,
        )
    if "unsupported extension" in lowered:
        return (detail, detail)
    if "file not found" in lowered:
        return (
            f"The {file_label} is no longer available on disk. Please upload it again.",
            detail,
        )
    if "target sheet" in lowered:
        return (
            "The target consumer-master worksheet could not be identified. "
            "Select the correct sheet after the template loads.",
            detail,
        )
    return (
        f"The {file_label} could not be processed. {detail} "
        "Review the file and try again. Technical details are available below.",
        detail,
    )


def humanize_validation_issue(category: str, message: str, field: str = "", row: int = -1) -> str:
    location_bits = []
    if field:
        location_bits.append(f"column `{field}`")
    if row and row > 0:
        location_bits.append(f"row {row}")
    location = f" ({', '.join(location_bits)})" if location_bits else ""
    action = {
        "reference": "Check `ref_project_masters_lov` for a matching name and a valid UUID.",
        "hierarchy": "Verify the hierarchy report contains unique codes for each network level.",
        "consumer": "Confirm the consumer identifier exists and is unique in the source workbook.",
        "uuid": "Replace the value with a valid UUID from the template reference or hierarchy report.",
        "source": "Add the missing source column or map it on the Mappings step.",
        "output": "Resolve the listed issues before treating the workbook as a successful master update.",
        "mapping": "Adjust the source-to-target mapping, then run validation again.",
    }.get(category, "Correct the listed issue and repeat the affected step.")
    return f"{message}{location} {action}"


_GENERIC_ERROR = re.compile(r"^(error|failed|exception)[:\s]*processing", re.I)


def is_generic_error(message: str) -> bool:
    return bool(_GENERIC_ERROR.match((message or "").strip()))
