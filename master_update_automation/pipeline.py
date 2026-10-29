"""Reusable transformation pipeline from source data to the target template."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd

from .audit_logger import AuditEntry, AuditLogger, FieldResolution
from .column_mapper import (
    MappingResult,
    create_column_mapping,
    get_hierarchy_mapping_dict,
    get_mapping_dict,
)
from .config import (
    BUSINESS_KEY_SOURCE_CANDIDATES,
    BUSINESS_KEY_TARGET,
    CATEGORY_SOURCE_COLUMNS,
    CONSUMER_UAN_SUFFIX_VALUE,
    FIXED_METER_CATEGORY_LOOKUP_NAME,
    HIERARCHY_LEVELS,
    MANDATORY_HIERARCHY_FIELDS,
    MANDATORY_HIERARCHY_LEVELS,
    MANDATORY_TARGET_FIELDS,
    OPTIONAL_LOOKUP_FIELDS,
    REFERENCE_CATEGORY_MAP,
    REQUIRED_SOURCE_COLUMNS,
    TARGET_FIELD_TYPES,
    TARGET_HIERARCHY_MAP,
    UUID_FIELDS,
    VALUE_ALIASES,
    AuditStatus,
    Category,
    FieldType,
    Operation,
)
from .consumer_rules import apply_consumer_identification_rules, apply_meter_make_rules
from .file_loader import (
    detect_target_sheet,
    is_csv_file,
    iter_csv_data_chunks,
    load_excel_workbook,
    read_csv_headers,
    read_hierarchy_file,
    read_source_data,
)
from .hierarchy_mapper import HierarchyMapper, HierarchyRecord
from .reference_lookup import LookupResult, ReferenceLookup
from .schema_detector import TemplateSchema, detect_template_schema
from .uuid_utils import is_valid_uuid, normalise_uuid
from .validator import (
    ValidationReport,
    validate_output_dataset,
    validate_reference_categories,
    validate_source_data,
)

logger = logging.getLogger(__name__)
_SCIENTIFIC_IDENTIFIER = re.compile(r"^[+-]?\d+(?:\.\d+)?e[+-]?\d+$", re.IGNORECASE)


@dataclass
class PipelineResult:
    records: List[Dict[str, str]]
    audit: AuditLogger
    validation: ValidationReport
    mapping: MappingResult
    hierarchy_summary: Dict
    template_schema: TemplateSchema
    ready_for_output: bool
    source_row_count: int
    duplicate_identifiers: Dict[str, List[int]] = field(default_factory=dict)


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "<na>"}:
        return ""
    return text


def _alias_value(field: str, value: str) -> str:
    aliases = VALUE_ALIASES.get(field, {})
    return aliases.get(value.upper(), value) if value else value


def find_source_key_column(source_headers: List[str]) -> Optional[str]:
    lower_map = {h.strip().lower(): h for h in source_headers if h}
    for candidate in BUSINESS_KEY_SOURCE_CANDIDATES:
        if candidate in lower_map:
            return lower_map[candidate]
    return None


def _row_columns(row) -> List[str]:
    return list(row.keys()) if hasattr(row, "keys") else list(row.index)


def extract_hierarchy_keys(
    row,
    hierarchy_map: Dict[str, str],
) -> Dict[str, str]:
    keys: Dict[str, str] = {}
    for src_col, internal in hierarchy_map.items():
        level_field = internal.replace("_hierarchy_", "")
        keys[level_field] = _clean(row.get(src_col, ""))
    for level in HIERARCHY_LEVELS:
        for kind in ("code", "name"):
            field = f"{level}_{kind}"
            if field in keys and keys[field]:
                continue
            candidates = [
                col for col in _row_columns(row)
                if str(col).strip().lower().replace("_", " ").replace("-", " ")
                == f"{level} {kind}"
            ]
            if len(candidates) == 1:
                keys[field] = _clean(row.get(candidates[0], ""))
    return keys


def _source_category_hint(row) -> str:
    for col in _row_columns(row):
        if str(col).strip().lower() in CATEGORY_SOURCE_COLUMNS:
            return _clean(row.get(col, "")).upper()
    return ""


def _ordered_record(headers: List[str], values: Dict[str, str]) -> Dict[str, str]:
    lookup = {k.lower().strip(): v for k, v in values.items()}
    ordered: Dict[str, str] = {}
    for header in headers:
        ordered[header] = lookup.get(header.lower().strip(), "")
    return ordered


def transform_record(
    row,
    source_row_number: int,
    operation: Operation,
    mapping_result: MappingResult,
    ref_lookup: ReferenceLookup,
    hierarchy_mapper: HierarchyMapper,
    target_headers: List[str],
    match_action: str,
    business_key: str,
    template_schema: Optional[TemplateSchema] = None,
    approve_hierarchy_correction: bool = False,
) -> Tuple[Dict[str, str], AuditEntry]:
    entry = AuditEntry(
        source_row=source_row_number,
        business_key=business_key,
        operation=operation.label,
        status=AuditStatus.INSERTED.value if match_action == "INSERT" else AuditStatus.UPDATED.value,
        action=match_action,
    )
    values: Dict[str, str] = {h.lower().strip(): "" for h in target_headers}
    mapping_dict = get_mapping_dict(mapping_result)
    hierarchy_map = get_hierarchy_mapping_dict(mapping_result)
    schema_fields = template_schema.fields if template_schema else {}
    inverse_mappings: Dict[str, List[str]] = {}
    for src_col, tgt_col in mapping_dict.items():
        inverse_mappings.setdefault(tgt_col.lower().strip(), []).append(src_col)

    ambiguous_by_target: Dict[str, List[str]] = {}
    for target, sources in mapping_result.ambiguous_mappings:
        ambiguous_by_target[target.lower().strip()] = sources

    def _field_schema(field_name: str):
        return schema_fields.get(field_name.lower().strip())

    def _is_mandatory(field_name: str) -> bool:
        field_schema = _field_schema(field_name)
        if field_schema:
            return field_schema.is_mandatory
        key = field_name.lower().strip()
        return key in MANDATORY_TARGET_FIELDS or key in MANDATORY_HIERARCHY_FIELDS

    def _is_uuid_field(field_name: str) -> bool:
        field_schema = _field_schema(field_name)
        return field_name.lower().strip() in UUID_FIELDS or bool(
            field_schema and field_schema.is_uuid
        )

    def _default_method(field_name: str) -> str:
        key = field_name.lower().strip()
        field_schema = _field_schema(key)
        field_type = (
            field_schema.field_type if field_schema else TARGET_FIELD_TYPES.get(key, FieldType.TEXT)
        )
        if key in TARGET_HIERARCHY_MAP or field_type == FieldType.HIERARCHY_LOOKUP:
            return "Hierarchy lookup"
        if key == "meter_category":
            return "Approved Main Meter reference rule"
        if key == "consumer_uan_suffix":
            return "Fixed business rule: consumer UAN suffix -1"
        if key == "consumer_identy":
            return "Operation-specific consumer identification rule and reference lookup"
        if key == "meter_make":
            return f"Destination-category meter make reference lookup ({operation.destination.value})"
        if key in REFERENCE_CATEGORY_MAP or field_type == FieldType.REFERENCE_LOOKUP:
            category = REFERENCE_CATEGORY_MAP.get(key, "configured reference category")
            return f"Reference lookup: {category}"
        if field_schema and field_schema.is_ref_lookup and key not in REFERENCE_CATEGORY_MAP:
            return "Reference lookup (category not configured)"
        return "Direct source mapping"

    field_resolutions: Dict[str, FieldResolution] = {}
    reported_field_issues: set = set()
    for header in target_headers:
        key = header.lower().strip()
        mandatory = _is_mandatory(header)
        sources = inverse_mappings.get(key, [])
        candidates = ambiguous_by_target.get(key, [])
        resolution = FieldResolution(
            target_field=header,
            source_field="; ".join(candidates or sources),
            resolution_method=_default_method(header),
            status="UNRESOLVED",
            reason=(
                f"Conflicting source columns map to '{header}'"
                if candidates else (
                    "No source column mapping exists for this target field."
                    if not sources and _default_method(header) == "Direct source mapping"
                    else "No verified value has been resolved."
                )
            ),
            mandatory=mandatory,
            corrective_action=(
                f"Resolve '{header}' before treating this record as complete."
                if mandatory else "No action required; this field is optional."
            ),
        )
        if candidates:
            resolution.status = "AMBIGUOUS"
        field_resolutions[key] = resolution

    def _set_field(
        field_name: str,
        value: str,
        status: str,
        method: str,
        source_field: str = "",
        reason: str = "",
        resolved_value: str = "",
    ) -> None:
        key = field_name.lower().strip()
        values[key] = value
        resolution = field_resolutions.get(key)
        if resolution:
            resolution.source_field = source_field
            resolution.resolution_method = method
            resolution.status = status
            resolution.reason = reason
            resolution.corrective_action = ""
            source_values = [
                _clean(row.get(source_name, ""))
                for source_name in source_field.split("; ")
                if source_name
            ]
            resolution.source_value = "; ".join(source_values)
            resolution.resolved_value = resolved_value or value
            resolution.output_value = value

    def _unresolved_field(
        field_name: str,
        reason: str,
        method: Optional[str] = None,
        source_field: str = "",
        status: str = "UNRESOLVED",
        severity: str = "",
    ) -> None:
        key = field_name.lower().strip()
        resolution = field_resolutions.get(key)
        mandatory = _is_mandatory(field_name)
        if resolution:
            resolution.source_field = source_field
            resolution.resolution_method = method or _default_method(field_name)
            resolution.status = status
            resolution.reason = reason
            resolution.mandatory = mandatory
            source_values = [
                _clean(row.get(source_name, ""))
                for source_name in source_field.split("; ")
                if source_name
            ]
            resolution.source_value = "; ".join(source_values)
            resolution.resolved_value = ""
            resolution.output_value = ""
            resolution.corrective_action = (
                f"Resolve '{field_name}' before treating this record as complete."
                if mandatory else "No action required; this field is optional."
            )
        message = f"Target field '{field_name}' is {status.lower()}: {reason}"
        if severity == "error" or (mandatory and severity != "warning"):
            entry.errors.append(message)
            reported_field_issues.add(key)
        elif severity == "warning":
            entry.warnings.append(message)
            reported_field_issues.add(key)

    entry.field_resolutions = list(field_resolutions.values())

    # This target field is defined as a fixed value for every generated row.
    # Do not infer it from an optional or similarly named source column.
    if "consumer_uan_suffix" in values:
        _set_field(
            "consumer_uan_suffix",
            CONSUMER_UAN_SUFFIX_VALUE,
            "APPROVED_DEFAULT",
            _default_method("consumer_uan_suffix"),
            resolved_value=CONSUMER_UAN_SUFFIX_VALUE,
        )

    hint = _source_category_hint(row)
    if hint and hint in {c.value for c in Category} and hint != operation.source.value:
        entry.warnings.append(
            f"Source meter category '{hint}' does not match selected source "
            f"'{operation.source.value}'"
        )

    for src_col, tgt_col in mapping_dict.items():
        tgt_lower = tgt_col.lower().strip()
        if tgt_lower == "consumer_uan_suffix":
            continue
        field_schema = _field_schema(tgt_lower)
        ft = field_schema.field_type if field_schema else TARGET_FIELD_TYPES.get(tgt_lower, FieldType.TEXT)
        if ft in (FieldType.REFERENCE_LOOKUP, FieldType.BUSINESS_RULE, FieldType.HIERARCHY_LOOKUP):
            continue
        source_value = _clean(row.get(src_col, ""))
        raw = _alias_value(tgt_lower, source_value)
        direct_method = (
            "Direct source mapping with configured alias transformation"
            if source_value and raw != source_value else "Direct source mapping"
        )
        if not raw:
            _unresolved_field(
                tgt_col,
                f"Source field '{src_col}' is blank.",
                source_field=src_col,
                severity="error" if _is_mandatory(tgt_col) else "",
            )
        elif _is_uuid_field(tgt_lower):
            if is_valid_uuid(raw):
                _set_field(
                    tgt_lower,
                    normalise_uuid(raw),
                    "MAPPED_FROM_SOURCE",
                    "Direct UUID source mapping with canonical brace formatting",
                    src_col,
                )
            else:
                _unresolved_field(
                    tgt_col,
                    f"Source field '{src_col}' contains an invalid UUID: {raw}",
                    method="Direct UUID source mapping",
                    source_field=src_col,
                    status="INVALID",
                    severity="error" if _is_mandatory(tgt_col) else "warning",
                )
        else:
            _set_field(tgt_lower, raw, "MAPPED_FROM_SOURCE", direct_method, src_col)

    special_targets = set(REFERENCE_CATEGORY_MAP) | {
        "meter_make", "consumer_identy", "consumer_uan_suffix"
    }
    for target, sources in mapping_result.ambiguous_mappings:
        key = target.lower().strip()
        if key.startswith("_hierarchy_") or key in special_targets:
            continue
        _unresolved_field(
            target,
            f"Conflicting source columns map to '{target}': {', '.join(sources)}",
            method="Source column mapping",
            source_field="; ".join(sources),
            status="AMBIGUOUS",
            severity="error" if _is_mandatory(target) else "warning",
        )
        entry.status = AuditStatus.REQUIRES_REVIEW.value

    hierarchy_keys = extract_hierarchy_keys(row, hierarchy_map)
    hierarchy: HierarchyRecord = hierarchy_mapper.match_consumer(row, hierarchy_keys)
    entry.hierarchy_corrections = [dict(item) for item in hierarchy.suggested_corrections]
    if approve_hierarchy_correction and hierarchy.suggested_corrections:
        hierarchy.requires_review = False
        hierarchy.warnings = [
            warning for warning in hierarchy.warnings
            if not warning.startswith("Hierarchy conflict:")
        ]
        entry.hierarchy_corrections = [
            {**item, "approved": "true"} for item in entry.hierarchy_corrections
        ]
        entry.transformations.extend(
            f"Approved hierarchy correction: {item['level']} {item['kind']} "
            f"'{item['source_value']}' → '{item['suggested_value']}'"
            for item in hierarchy.suggested_corrections
        )
    entry.warnings.extend(hierarchy.warnings)
    if hierarchy.errors:
        entry.errors.extend(hierarchy.errors)
    if hierarchy.requires_review:
        entry.hierarchy_status = "REQUIRES_REVIEW"
        entry.status = AuditStatus.REQUIRES_REVIEW.value
    elif any(m.is_ambiguous for m in hierarchy.matches.values()):
        entry.hierarchy_status = "AMBIGUOUS"
        entry.status = AuditStatus.REQUIRES_REVIEW.value
    elif not any(m.success for m in hierarchy.matches.values()):
        entry.hierarchy_status = "MISSING"
    else:
        failed_required = [
            lvl for lvl in MANDATORY_HIERARCHY_LEVELS
            if not hierarchy.matches.get(lvl, None) or not hierarchy.matches[lvl].success
        ]
        entry.hierarchy_status = "OK" if not failed_required else "PARTIAL"

    matched_level = next(
        (level for level in reversed(HIERARCHY_LEVELS)
         if hierarchy.matches.get(level) and hierarchy.matches[level].success),
        "",
    )
    matched_level_sources = [
        src_col for src_col, internal in hierarchy_map.items()
        if matched_level and (
            internal.lower().endswith(f"_{matched_level}_code")
            or internal.lower().endswith(f"_{matched_level}_name")
        )
        and hierarchy_keys.get(internal.replace("_hierarchy_", ""), "")
    ]
    for target_field, (level, kind) in TARGET_HIERARCHY_MAP.items():
        match = hierarchy.matches.get(level)
        level_sources = [
            src_col for src_col, internal in hierarchy_map.items()
            if internal.lower().endswith(f"_{level}_code")
            or internal.lower().endswith(f"_{level}_name")
            if hierarchy_keys.get(internal.replace("_hierarchy_", ""), "")
        ]
        source_fields = level_sources or matched_level_sources
        source_field = "; ".join(source_fields)
        hierarchy_ambiguity = next(
            (
                sources for target, sources in mapping_result.ambiguous_mappings
                if target.lower().strip().startswith(f"_hierarchy_{level}_")
            ),
            None,
        )
        if hierarchy_ambiguity:
            _unresolved_field(
                target_field,
                f"Conflicting source hierarchy columns: {', '.join(hierarchy_ambiguity)}",
                method="Hierarchy source mapping",
                source_field="; ".join(hierarchy_ambiguity),
                status="AMBIGUOUS",
                severity="error" if _is_mandatory(target_field) else "warning",
            )
            entry.status = AuditStatus.REQUIRES_REVIEW.value
        elif match and match.success:
            if hierarchy.requires_review:
                method = (
                    f"Hierarchy path resolved by {matched_level} identifier; "
                    "source hierarchy conflict requires review"
                )
                source_field = "; ".join(dict.fromkeys(matched_level_sources + level_sources))
            elif matched_level and matched_level != level:
                verify = f"; verified by {', '.join(level_sources)}" if level_sources else ""
                method = f"Hierarchy path resolved by {matched_level} identifier{verify}"
                source_field = "; ".join(dict.fromkeys(matched_level_sources + level_sources))
            else:
                method = "Exact hierarchy identifier match"
            if approve_hierarchy_correction and hierarchy.suggested_corrections:
                method = f"{method}; user approved hierarchy correction"
            suggested_only = hierarchy.requires_review and bool(hierarchy.suggested_corrections)
            display_value = (
                match.name if target_field.lower().endswith("_name")
                else match.code if target_field.lower().endswith("_code")
                else match.uuid
            )
            _set_field(
                target_field,
                match.uuid,
                "USER_APPROVED_HIERARCHY" if approve_hierarchy_correction and hierarchy.suggested_corrections else "SUGGESTED_HIERARCHY" if suggested_only else "RESOLVED_FROM_HIERARCHY",
                method,
                source_field,
                resolved_value=display_value,
            )
            if suggested_only and field_resolutions.get(target_field.lower().strip()):
                field_resolutions[target_field.lower().strip()].corrective_action = (
                    "Review and accept the suggested hierarchy correction before using this value."
                )
        else:
            supplied = bool(
                hierarchy_keys.get(f"{level}_code", "")
                or hierarchy_keys.get(f"{level}_name", "")
            )
            ambiguous = bool(match and match.is_ambiguous)
            invalid = bool(match and "invalid" in match.error.lower())
            reason = (match.error if match and match.error else f"No source identifier was provided for {level}.")
            _unresolved_field(
                target_field,
                reason,
                method="Exact hierarchy identifier match",
                source_field=source_field,
                status="AMBIGUOUS" if ambiguous else "INVALID" if invalid else "UNRESOLVED",
                severity=(
                    "error" if _is_mandatory(target_field)
                    else "warning" if supplied else ""
                ),
            )

    source_ident_col = None
    source_make_col = None
    for m in mapping_result.mappings:
        if m.target_column.lower().strip() == "consumer_identy":
            source_ident_col = m.source_column
        if m.target_column.lower().strip() == "meter_make":
            source_make_col = m.source_column
    if source_ident_col is None and "consumer_identy" not in ambiguous_by_target:
        for col in _row_columns(row):
            if str(col).strip().lower() == "consumer identification":
                source_ident_col = col
    if source_make_col is None and "meter_make" not in ambiguous_by_target:
        for col in _row_columns(row):
            if str(col).strip().lower() == "meter make":
                source_make_col = col

    ident_value = _clean(row.get(source_ident_col, "")) if source_ident_col is not None else ""
    if "consumer_identy" in ambiguous_by_target:
        ident_result = LookupResult(
            success=False,
            error=(
                "Conflicting source columns map to consumer_identy: "
                + ", ".join(ambiguous_by_target["consumer_identy"])
            ),
            is_ambiguous=True,
        )
        ident_rule = "Ambiguous source field mapping"
    else:
        ident_result, ident_rule = apply_consumer_identification_rules(operation, ident_value, ref_lookup)
    entry.transformations.append(ident_rule)
    if ident_result.success and is_valid_uuid(ident_result.uuid):
        rule_applied = ident_rule.startswith(("Rule A", "Rule B"))
        _set_field(
            "consumer_identy",
            normalise_uuid(ident_result.uuid),
            "DERIVED_BY_RULE" if rule_applied else "RESOLVED_FROM_REFERENCE",
            ident_rule,
            source_ident_col or "",
            resolved_value=ident_result.name or ident_value,
        )
    else:
        entry.failed_lookups.append(f"consumer_identy={ident_value}")
        _unresolved_field(
            "consumer_identy",
            ident_result.error or "Consumer identification lookup failed or returned an invalid UUID.",
            method=ident_rule,
            source_field=source_ident_col or "",
            status="AMBIGUOUS" if ident_result.is_ambiguous else "INVALID" if ident_result.is_invalid or ident_result.success else "UNRESOLVED",
            severity="error" if _is_mandatory("consumer_identy") or ident_result.is_ambiguous else "warning",
        )
        if ident_result.is_ambiguous:
            entry.status = AuditStatus.REQUIRES_REVIEW.value

    make_value = _clean(row.get(source_make_col, "")) if source_make_col is not None else ""
    if "meter_make" in ambiguous_by_target:
        _unresolved_field(
            "meter_make",
            "Conflicting source columns map to meter_make: " + ", ".join(ambiguous_by_target["meter_make"]),
            method="Meter make source mapping",
            source_field="; ".join(ambiguous_by_target["meter_make"]),
            status="AMBIGUOUS",
            severity="error" if _is_mandatory("meter_make") else "warning",
        )
        entry.status = AuditStatus.REQUIRES_REVIEW.value
    elif make_value:
        make_result, make_rule = apply_meter_make_rules(operation, make_value, ref_lookup)
        entry.transformations.append(make_rule)
        if make_result.success and is_valid_uuid(make_result.uuid):
            _set_field(
                "meter_make",
                normalise_uuid(make_result.uuid),
                "RESOLVED_FROM_REFERENCE",
                make_rule,
                source_make_col or "",
                resolved_value=make_result.name or make_value,
            )
        else:
            entry.failed_lookups.append(f"meter_make={make_value}")
            _unresolved_field(
                "meter_make",
                make_result.error or "Meter make lookup failed or returned an invalid UUID.",
                method=make_rule,
                source_field=source_make_col or "",
                status="AMBIGUOUS" if make_result.is_ambiguous else "INVALID" if make_result.is_invalid or make_result.success else "UNRESOLVED",
                severity="error" if _is_mandatory("meter_make") else "warning",
            )
            if make_result.is_ambiguous:
                entry.status = AuditStatus.REQUIRES_REVIEW.value
    else:
        entry.transformations.append("Meter make left blank (source empty)")
        _unresolved_field(
            "meter_make",
            (f"Source field '{source_make_col}' is blank." if source_make_col else "Meter make source field is missing."),
            method=_default_method("meter_make"),
            source_field=source_make_col or "",
            severity="error" if _is_mandatory("meter_make") else "",
        )

    lookup_source_by_target: Dict[str, str] = {}
    for m in mapping_result.mappings:
        if not m.target_column.startswith("_"):
            lookup_source_by_target[m.target_column.lower().strip()] = m.source_column

    for target_field, _category in REFERENCE_CATEGORY_MAP.items():
        if target_field in {"consumer_identy", "meter_make"}:
            continue
        if target_field == "meter_category":
            raw = FIXED_METER_CATEGORY_LOOKUP_NAME
            source_col = ""
            method = "Approved default rule: Main Meter reference lookup"
        else:
            src_col = lookup_source_by_target.get(target_field)
            raw = _clean(row.get(src_col, "")) if src_col else ""
            raw = _alias_value(target_field, raw)
            source_col = src_col or ""
            method = f"Reference lookup in {_category}"
            if not raw:
                ambiguous_sources = ambiguous_by_target.get(target_field)
                _unresolved_field(
                    target_field,
                    (
                        f"Conflicting source columns map to '{target_field}': "
                        + ", ".join(ambiguous_sources)
                        if ambiguous_sources else
                        f"Source field '{src_col}' is blank." if src_col else "No source field is mapped."
                    ),
                    method=method,
                    source_field="; ".join(ambiguous_sources) if ambiguous_sources else source_col,
                    status="AMBIGUOUS" if ambiguous_sources else "UNRESOLVED",
                    severity="error" if _is_mandatory(target_field) else "",
                )
                if ambiguous_sources:
                    entry.status = AuditStatus.REQUIRES_REVIEW.value
                continue
        result: LookupResult = ref_lookup.lookup_for_target_field(
            target_field, raw, dest_category=operation.destination
        )
        if result.success and is_valid_uuid(result.uuid):
            _set_field(
                target_field,
                normalise_uuid(result.uuid),
                "APPROVED_DEFAULT" if target_field == "meter_category" else "RESOLVED_FROM_REFERENCE",
                method,
                source_col,
                f"Resolved lookup value '{raw}'." if target_field == "meter_category" else "",
                resolved_value=raw,
            )
            entry.transformations.append(f"{target_field}: '{raw}' → {result.uuid}")
        else:
            entry.failed_lookups.append(f"{target_field}={raw}")
            if target_field == "meter_category" and not result.error:
                result.error = (
                    f"Approved value '{FIXED_METER_CATEGORY_LOOKUP_NAME}' in "
                    f"'{_category}' has no valid UUID."
                )
            status = "AMBIGUOUS" if result.is_ambiguous else "INVALID" if result.is_invalid or (result.success and result.uuid) else "UNRESOLVED"
            _unresolved_field(
                target_field,
                result.error or f"Value '{raw}' could not be resolved in '{_category}'.",
                method=method,
                source_field=source_col,
                status=status,
                severity=(
                    "error" if _is_mandatory(target_field) or result.is_ambiguous
                    else "warning"
                ),
            )
            if result.is_ambiguous:
                entry.status = AuditStatus.REQUIRES_REVIEW.value

    for field_name in UUID_FIELDS:
        current = values.get(field_name, "")
        if current:
            if is_valid_uuid(current):
                values[field_name] = normalise_uuid(current)
            else:
                values[field_name] = ""
                _unresolved_field(
                    field_name,
                    f"Resolved value is not a valid UUID: {current}",
                    method=_default_method(field_name),
                    status="INVALID",
                    severity="error" if _is_mandatory(field_name) else "warning",
                )

    for header in target_headers:
        key = header.lower().strip()
        resolution = field_resolutions.get(key)
        if (
            resolution and resolution.mandatory and not values.get(key, "")
            and key not in reported_field_issues
        ):
            _unresolved_field(
                header,
                resolution.reason or "Mandatory field has no resolved output value.",
                method=resolution.resolution_method,
                source_field=resolution.source_field,
                status=resolution.status,
                severity="error",
            )

    # Include every target column in field audit, including fields with no
    # configured source mapping and fields newly introduced by the template.
    entry.field_resolutions = [
        field_resolutions[h.lower().strip()]
        for h in target_headers
        if h.lower().strip() in field_resolutions
    ]

    if match_action in {"FAILED", "DUPLICATE", "AMBIGUOUS"}:
        entry.status = (
            AuditStatus.REQUIRES_REVIEW.value
            if match_action == "AMBIGUOUS"
            else AuditStatus.FAILED.value
        )
        if match_action != "AMBIGUOUS":
            entry.errors.append(f"Record matching action: {match_action}")
    elif entry.errors:
        if entry.status != AuditStatus.REQUIRES_REVIEW.value:
            entry.status = AuditStatus.FAILED.value
    elif entry.warnings:
        if entry.status not in {AuditStatus.REQUIRES_REVIEW.value}:
            entry.status = AuditStatus.WARNING.value
    elif match_action == "UPDATE":
        entry.status = AuditStatus.UPDATED.value
    elif match_action == "INSERT":
        entry.status = AuditStatus.INSERTED.value

    return _ordered_record(target_headers, values), entry


def run_pipeline(
    template_path: str,
    target_sheet: str,
    operation: Operation,
    hierarchy_path: str,
    source_path: str,
    source_sheet: Optional[str] = None,
    custom_map: Optional[Dict[str, str]] = None,
    progress: Optional[Callable[[str], None]] = None,
    approved_hierarchy_rows: Optional[set] = None,
    source_row_count_hint: Optional[int] = None,
) -> PipelineResult:
    def _progress(message: str) -> None:
        if progress:
            progress(message)

    try:
        return _run_pipeline(
            template_path=template_path,
            target_sheet=target_sheet,
            operation=operation,
            hierarchy_path=hierarchy_path,
            source_path=source_path,
            source_sheet=source_sheet,
            custom_map=custom_map,
            notify=_progress,
            approved_hierarchy_rows=approved_hierarchy_rows or set(),
            source_row_count_hint=source_row_count_hint,
        )
    except MemoryError as exc:
        raise MemoryError(
            "The current environment does not have enough memory to process this file. "
            "Reduce the file size, close other applications, or run on a machine with more RAM."
        ) from exc


def _run_pipeline(
    template_path: str,
    target_sheet: str,
    operation: Operation,
    hierarchy_path: str,
    source_path: str,
    source_sheet: Optional[str],
    custom_map: Optional[Dict[str, str]],
    notify: Callable[[str], None],
    approved_hierarchy_rows: set,
    source_row_count_hint: Optional[int],
) -> PipelineResult:
    notify("Reading input files")
    wb = load_excel_workbook(template_path)
    if not target_sheet:
        detected = detect_target_sheet(wb)
        if not detected:
            wb.close()
            raise ValueError("Target sheet could not be identified. Please select it explicitly.")
        target_sheet = detected
    notify("Validating workbook structure")
    schema = detect_template_schema(wb, target_sheet)
    ref_lookup = ReferenceLookup()
    ref_lookup.load_from_workbook(wb)
    wb.close()

    source_is_csv = is_csv_file(source_path)
    if source_is_csv:
        source_headers = read_csv_headers(source_path)
        empty_source_frame = pd.DataFrame(columns=source_headers)
        source_count_hint = source_row_count_hint
        source_chunks = iter_csv_data_chunks(source_path)
    else:
        source_df = read_source_data(source_path, source_sheet, skip_type_hint_row=False)
        source_headers = list(source_df.columns)
        empty_source_frame = source_df.iloc[0:0]
        source_count_hint = len(source_df)
        source_chunks = iter((source_df,))
    hierarchy_df = read_hierarchy_file(hierarchy_path)
    mapper = HierarchyMapper()
    hierarchy_summary = mapper.load_hierarchy(hierarchy_df)
    notify("Preparing reference lookups and hierarchy indexes")

    mapping = create_column_mapping(
        source_headers,
        schema.headers,
        custom_map=custom_map,
    )
    validation = ValidationReport()
    validation.merge(validate_reference_categories(ref_lookup, operation.destination))
    validation.merge(validate_source_data(
        empty_source_frame,
        REQUIRED_SOURCE_COLUMNS,
        row_count=source_count_hint,
        allow_empty_frame=source_is_csv and source_count_hint is None,
    ))
    for target, sources in mapping.ambiguous_mappings:
        field = schema.fields.get(target.lower().strip())
        mandatory = bool(field and field.is_mandatory) or target.lower().strip() in (
            MANDATORY_TARGET_FIELDS | MANDATORY_HIERARCHY_FIELDS
        )
        message = (
            f"Conflicting source columns map to '{target}': {', '.join(sources)}"
        )
        if mandatory:
            validation.add_error("mapping", message, field=target)
        else:
            validation.add_warning("mapping", message, field=target)
    for msg in mapper.validate_hierarchy_columns():
        validation.add_warning("hierarchy", msg)

    source_key = find_source_key_column(source_headers)
    if not source_key:
        validation.add_error("consumer", "Could not identify a consumer business key column")
        audit = AuditLogger(operation.label)
        return PipelineResult(
            records=[],
            audit=audit,
            validation=validation,
            mapping=mapping,
            hierarchy_summary=hierarchy_summary,
            template_schema=schema,
            ready_for_output=False,
            source_row_count=int(source_count_hint or 0),
        )

    audit = AuditLogger(operation.label)
    records: List[Dict[str, str]] = []
    notify("Resolving reference UUIDs, matching hierarchy, and applying business rules")
    seen_keys: Dict[str, int] = {}
    duplicates: Dict[str, List[int]] = {}
    processed = 0
    scientific_identifiers: List[Tuple[int, str]] = []
    scientific_identifier_count = 0
    for chunk in source_chunks:
        chunk_columns = list(chunk.columns)
        for source_index, row_values in zip(chunk.index, chunk.itertuples(index=False, name=None)):
            row = dict(zip(chunk_columns, row_values))
            row_index = processed if source_is_csv else int(source_index)
            excel_row = row_index + 2
            key = _clean(row.get(source_key, ""))
            match_error = ""
            if not key:
                action = "FAILED"
                match_error = f"Missing business key in column '{source_key}'"
            elif key in seen_keys:
                action = "DUPLICATE"
                match_error = (
                    f"Duplicate source key '{key}' "
                    f"(first seen at row {seen_keys[key]})"
                )
                duplicates.setdefault(key, [seen_keys[key]]).append(row_index)
            else:
                action = "INSERT"
                seen_keys[key] = row_index

            if _SCIENTIFIC_IDENTIFIER.fullmatch(key):
                scientific_identifier_count += 1
                if len(scientific_identifiers) < 100:
                    scientific_identifiers.append((excel_row, key))

            record, entry = transform_record(
                row=row,
                source_row_number=excel_row,
                operation=operation,
                mapping_result=mapping,
                ref_lookup=ref_lookup,
                hierarchy_mapper=mapper,
                target_headers=schema.headers,
                match_action=action,
                business_key=key,
                template_schema=schema,
                approve_hierarchy_correction=excel_row in approved_hierarchy_rows,
            )
            if match_error:
                entry.errors.append(match_error)
                entry.status = AuditStatus.FAILED.value
            audit.add_entry(entry)
            records.append(record)
            processed += 1
            if processed % 500 == 0:
                if source_count_hint:
                    notify(f"Processing records: {processed:,} / {source_count_hint:,}")
                else:
                    notify(f"Processing records: {processed:,}")

    if processed == 0:
        if not any(
            issue.category == "consumer" and "empty" in issue.message.lower()
            for issue in validation.errors
        ):
            validation.add_error("consumer", "Source data is empty")
    elif not source_count_hint:
        validation.add_info("consumer", f"Source data has {processed:,} records")
    if processed and processed % 500:
        if source_count_hint:
            notify(f"Processing records: {processed:,} / {source_count_hint:,}")
        else:
            notify(f"Processing records: {processed:,}")
    if scientific_identifier_count:
        validation.add_warning(
            "consumer",
            f"{scientific_identifier_count} consumer identifier(s) are already written in scientific notation. "
            "The app will preserve the supplied text, but cannot recover digits lost before upload.",
            field=source_key,
        )
        for row_number, value in scientific_identifiers:
            validation.add_warning(
                "consumer",
                "Consumer identifier is in scientific notation; check the original source value.",
                field=source_key,
                row=row_number,
                value=value,
            )

    notify("Validating transformed data")
    validation.merge(validate_output_dataset(records, schema.headers, BUSINESS_KEY_TARGET))
    failed_or_review = any(
        e.status in {AuditStatus.FAILED.value, AuditStatus.REQUIRES_REVIEW.value}
        for e in audit.entries
    )
    ready = (not failed_or_review) and validation.error_count == 0
    if not ready:
        validation.add_warning(
            "output",
            "Final master workbook should not be treated as successful while "
            "required lookups, hierarchy matches, or validations remain unresolved.",
        )
    return PipelineResult(
        records=records,
        audit=audit,
        validation=validation,
        mapping=mapping,
        hierarchy_summary=hierarchy_summary,
        template_schema=schema,
        ready_for_output=ready,
        source_row_count=processed,
        duplicate_identifiers=duplicates,
    )
