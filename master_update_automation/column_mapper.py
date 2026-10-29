"""Source-to-target column mapping with automatic detection and overrides."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from .config import SOURCE_TO_TARGET_MAP, TARGET_FIELD_TYPES, FieldType

logger = logging.getLogger(__name__)


@dataclass
class ColumnMapping:
    source_column: str
    target_column: str
    transformation: str = "direct"
    lookup_category: str = ""
    is_confirmed: bool = False
    confidence: float = 1.0


@dataclass
class MappingResult:
    mappings: List[ColumnMapping]
    unmapped_source: List[str]
    unmapped_target: List[str]
    ambiguous_mappings: List[Tuple[str, List[str]]]

    @property
    def all_confirmed(self) -> bool:
        return all(m.is_confirmed for m in self.mappings)

    @property
    def mapped_target_columns(self) -> set:
        return {m.target_column for m in self.mappings}


def normalise_header(header: str) -> str:
    return (
        header.strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
        .replace("?", "")
        .replace("  ", " ")
    )


def _transformation_for(target_col: str) -> str:
    if target_col.startswith("_hierarchy_"):
        return "hierarchy"
    ft = TARGET_FIELD_TYPES.get(target_col, FieldType.TEXT)
    if ft == FieldType.REFERENCE_LOOKUP:
        return "reference_lookup"
    if ft == FieldType.BUSINESS_RULE:
        return "business_rule"
    if ft == FieldType.HIERARCHY_LOOKUP:
        return "hierarchy"
    return "direct"


def _actual_target(target_headers: List[str], logical_name: str) -> Optional[str]:
    for th in target_headers:
        if th and th.lower().strip() == logical_name.lower().strip():
            return th
    return None


def create_column_mapping(
    source_headers: List[str],
    target_headers: List[str],
    custom_map: Optional[Dict[str, str]] = None,
) -> MappingResult:
    mappings: List[ColumnMapping] = []
    mapped_sources: set = set()
    mapped_targets: set = set()
    ambiguous: List[Tuple[str, List[str]]] = []
    ambiguous_targets: set = set()

    target_norm_map: Dict[str, str] = {}
    for th in target_headers:
        if th:
            target_norm_map[normalise_header(th)] = th

    if custom_map:
        custom_by_target: Dict[str, List[str]] = {}
        target_spelling: Dict[str, str] = {}
        for src, tgt in custom_map.items():
            if src in source_headers and tgt in target_headers:
                key = tgt.lower().strip()
                custom_by_target.setdefault(key, []).append(src)
                target_spelling[key] = tgt
        for key, sources in custom_by_target.items():
            if len(set(sources)) > 1:
                ambiguous.append((target_spelling[key], list(dict.fromkeys(sources))))
                ambiguous_targets.add(key)
                mapped_sources.update(sources)
                continue
            src = sources[0]
            tgt = target_spelling[key]
            mappings.append(ColumnMapping(
                source_column=src,
                target_column=tgt,
                transformation=_transformation_for(tgt.lower().strip()),
                is_confirmed=True,
                confidence=1.0,
            ))
            mapped_sources.add(src)
            mapped_targets.add(tgt)

    # Gather every automatic candidate before selecting any mapping. Source
    # header order must not decide which of two fields supplies one output.
    candidates_by_target: Dict[str, List[Tuple[str, str, float]]] = {}
    for src_header in source_headers:
        if src_header in mapped_sources:
            continue
        src_norm = normalise_header(src_header)
        if src_norm in SOURCE_TO_TARGET_MAP:
            logical_target = SOURCE_TO_TARGET_MAP[src_norm]
            if logical_target.startswith("_hierarchy_"):
                target_col = logical_target
            else:
                target_col = _actual_target(target_headers, logical_target)
            confidence = 0.9
        elif src_norm in target_norm_map:
            target_col = target_norm_map[src_norm]
            confidence = 0.8
        else:
            continue
        if not target_col:
            continue
        key = target_col.lower().strip()
        if key in ambiguous_targets or target_col in mapped_targets:
            continue
        candidates_by_target.setdefault(key, []).append((src_header, target_col, confidence))

    for key, candidates in candidates_by_target.items():
        unique_sources = list(dict.fromkeys(src for src, _, _ in candidates))
        if len(unique_sources) > 1:
            ambiguous.append((candidates[0][1], unique_sources))
            ambiguous_targets.add(key)
            mapped_sources.update(unique_sources)
            continue
        src_header, target_col, confidence = candidates[0]
        transformation = _transformation_for(target_col.lower().strip())
        mappings.append(ColumnMapping(
            source_column=src_header,
            target_column=target_col,
            transformation=transformation,
            is_confirmed=True,
            confidence=confidence,
        ))
        mapped_sources.add(src_header)
        if not target_col.startswith("_hierarchy_"):
            mapped_targets.add(target_col)

    unmapped_source = [s for s in source_headers if s and s not in mapped_sources]
    unmapped_target = [
        t for t in target_headers
        if t and t not in mapped_targets and t.lower().strip() not in {"id"}
    ]
    result = MappingResult(
        mappings=mappings,
        unmapped_source=unmapped_source,
        unmapped_target=unmapped_target,
        ambiguous_mappings=ambiguous,
    )
    logger.info(
        "Column mapping: %s mapped, %s unmapped source, %s unmapped target",
        len(mappings),
        len(unmapped_source),
        len(unmapped_target),
    )
    return result


def get_mapping_dict(mapping_result: MappingResult) -> Dict[str, str]:
    return {
        m.source_column: m.target_column
        for m in mapping_result.mappings
        if not m.target_column.startswith("_")
    }


def get_hierarchy_mapping_dict(mapping_result: MappingResult) -> Dict[str, str]:
    return {
        m.source_column: m.target_column
        for m in mapping_result.mappings
        if m.target_column.startswith("_hierarchy_")
    }
