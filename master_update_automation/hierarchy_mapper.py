"""Match consumer records to the network hierarchy report."""

from __future__ import annotations

import logging
from collections import OrderedDict
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .config import HIERARCHY_CSV_COLUMN_MAP, HIERARCHY_LEVELS
from .uuid_utils import is_valid_uuid, normalise_uuid

logger = logging.getLogger(__name__)


@dataclass
class HierarchyMatch:
    success: bool
    level: str = ""
    uuid: str = ""
    name: str = ""
    code: str = ""
    error: str = ""
    is_ambiguous: bool = False
    match_count: int = 0


@dataclass
class HierarchyRecord:
    matches: Dict[str, HierarchyMatch]
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    matched_row_index: Optional[int] = None
    requires_review: bool = False
    suggested_corrections: List[Dict[str, str]] = field(default_factory=list)

    @property
    def all_matched(self) -> bool:
        return all(m.success for m in self.matches.values())

    @property
    def failed_levels(self) -> List[str]:
        return [level for level, m in self.matches.items() if not m.success]


class HierarchyMapper:
    def __init__(self) -> None:
        self._df: Optional[pd.DataFrame] = None
        self._column_map: Dict[str, Dict[str, str]] = {}
        self._available_levels: List[str] = []
        self._code_indices: Dict[str, Dict[str, List[int]]] = {}
        self._name_indices: Dict[str, Dict[str, List[int]]] = {}
        self._duplicates: Dict[str, List[str]] = {}
        self._column_values: Dict[str, List[Any]] = {}
        self._match_cache: OrderedDict = OrderedDict()
        self._match_cache_limit = 50_000

    def load_hierarchy(self, df: pd.DataFrame) -> Dict[str, Any]:
        self._df = df.reset_index(drop=True)
        self._detect_columns()
        self._column_values = {
            column: self._df[column].tolist()
            for columns in self._column_map.values()
            for column in columns.values()
            if column in self._df.columns
        }
        self._build_indices()
        self._match_cache.clear()
        return {
            "total_rows": len(df),
            "available_levels": self._available_levels,
            "detected_columns": self._column_map,
            "missing_levels": [lvl for lvl in HIERARCHY_LEVELS if lvl not in self._available_levels],
            "duplicates": self._duplicates,
        }

    def _detect_columns(self) -> None:
        if self._df is None:
            return
        df_cols_lower = {c.strip().lower(): c for c in self._df.columns}
        self._column_map.clear()
        self._available_levels = []
        for level in HIERARCHY_LEVELS:
            expected = HIERARCHY_CSV_COLUMN_MAP.get(level, {})
            detected: Dict[str, str] = {}
            for field_type in ("name", "code", "id"):
                expected_col = expected.get(field_type, "")
                if not expected_col:
                    continue
                if expected_col in self._df.columns:
                    detected[field_type] = expected_col
                else:
                    exp_lower = expected_col.strip().lower()
                    if exp_lower in df_cols_lower:
                        detected[field_type] = df_cols_lower[exp_lower]
            if detected:
                self._column_map[level] = detected
                self._available_levels.append(level)

    def _build_indices(self) -> None:
        if self._df is None:
            return
        self._code_indices.clear()
        self._name_indices.clear()
        self._duplicates.clear()
        for level in self._available_levels:
            cols = self._column_map.get(level, {})
            code_col = cols.get("code")
            name_col = cols.get("name")
            id_col = cols.get("id")
            if name_col and name_col in self._df.columns:
                name_index: Dict[str, List[int]] = {}
                for idx, name_value in enumerate(self._column_values[name_col]):
                    if pd.isna(name_value):
                        continue
                    name_str = str(name_value).strip()
                    if name_str:
                        name_index.setdefault(name_str.casefold(), []).append(idx)
                self._name_indices[level] = name_index
            if not code_col or code_col not in self._df.columns:
                continue
            index: Dict[str, List[int]] = {}
            for idx, code_val in enumerate(self._column_values[code_col]):
                if pd.isna(code_val):
                    continue
                code_str = str(code_val).strip()
                if not code_str:
                    continue
                index.setdefault(code_str, []).append(idx)
            if id_col and id_col in self._df.columns:
                for code_str, idxs in index.items():
                    ids = {
                        self._identity_key(str(self._column_values[id_col][i]).strip())
                        for i in idxs
                        if pd.notna(self._column_values[id_col][i])
                    }
                    ids.discard("nan")
                    ids.discard("")
                    if len(ids) > 1:
                        self._duplicates.setdefault(level, []).append(code_str)
            self._code_indices[level] = index

    def match_consumer(
        self,
        source_record: Dict[str, str],
        hierarchy_keys: Dict[str, str],
    ) -> HierarchyRecord:
        key = tuple(
            (
                str(hierarchy_keys.get(f"{level}_{kind}", "") or "").strip().casefold()
                if kind == "name"
                else str(hierarchy_keys.get(f"{level}_{kind}", "") or "").strip()
            )
            for level in HIERARCHY_LEVELS
            for kind in ("code", "name")
        )
        cached = self._match_cache.get(key)
        if cached is not None:
            self._match_cache.move_to_end(key)
            return deepcopy(cached)
        result = self._match_consumer_uncached(source_record, hierarchy_keys)
        self._match_cache[key] = result
        if len(self._match_cache) > self._match_cache_limit:
            self._match_cache.popitem(last=False)
        return deepcopy(result)

    def _match_consumer_uncached(
        self,
        source_record: Dict[str, str],
        hierarchy_keys: Dict[str, str],
    ) -> HierarchyRecord:
        errors: List[str] = []
        warnings: List[str] = []
        matched_idx: Optional[int] = None
        used_level = ""
        reconciled_conflicts: List[Tuple[str, str, str, str]] = []
        requires_review = False

        search_order = ["dt", "feeder", "substation", "subdivision", "division", "circle", "zone", "discom"]
        for level in search_order:
            result = self._match_level(level, hierarchy_keys)
            if result is None:
                continue
            status, idxs = result
            if status == "missing":
                continue
            if status == "ambiguous":
                warnings.append(
                    f"Ambiguous match for {level} using available identifiers "
                    f"({len(idxs)} candidates). Record flagged for review."
                )
                return self._empty_record(
                    errors + [f"Ambiguous hierarchy match at {level}"],
                    warnings,
                    ambiguous_level=level,
                    match_count=len(idxs),
                )
            matched_idx = idxs[0]
            used_level = level
            break

        deepest_source_level = next(
            (
                level for level in reversed(HIERARCHY_LEVELS)
                if hierarchy_keys.get(f"{level}_code", "").strip()
                or hierarchy_keys.get(f"{level}_name", "").strip()
            ),
            "",
        )
        if deepest_source_level:
            deepest_index = HIERARCHY_LEVELS.index(deepest_source_level)
            matched_index = HIERARCHY_LEVELS.index(used_level) if used_level else -1
            if deepest_index > matched_index:
                reconciliation = self._reconcile_single_conflict(
                    deepest_source_level, hierarchy_keys
                )
                if reconciliation:
                    status, idxs, conflicts = reconciliation
                    if status == "unique":
                        matched_idx = idxs[0]
                        used_level = deepest_source_level
                        reconciled_conflicts = conflicts
                        requires_review = True
                        conflict_level, conflict_kind, source_value, matched_value = conflicts[0]
                        warnings.append(
                            f"Hierarchy conflict: source {conflict_level} {conflict_kind} "
                            f"'{source_value}' differs from the uniquely corroborated "
                            f"hierarchy path value '{matched_value}'. The path was resolved "
                            "from the other exact hierarchy identifiers; review the source value."
                        )
                    elif status == "ambiguous":
                        requires_review = True
                        warnings.append(
                            f"The source hierarchy identifiers conflict and do not identify "
                            f"one unique path through {deepest_source_level}; deeper hierarchy "
                            "values were not inferred."
                        )

        if matched_idx is None:
            errors.append("No matching hierarchy row found using codes or parent-qualified names")
            return self._empty_record(errors, warnings, requires_review=requires_review)

        assert self._df is not None
        matched_row = matched_idx
        matches: Dict[str, HierarchyMatch] = {}
        used_level_index = HIERARCHY_LEVELS.index(used_level)
        for level_index, level in enumerate(HIERARCHY_LEVELS):
            # A row matched by a parent entity (for example, a feeder) can
            # contain several child entities (DTs). Never infer those child
            # identifiers from whichever flattened row happened to come first.
            if level_index > used_level_index:
                matches[level] = HierarchyMatch(
                    success=False,
                    level=level,
                    error=(
                        f"Level '{level}' was not identified by the source; "
                        f"the hierarchy row was matched at '{used_level}'"
                    ),
                )
                continue
            if level not in self._available_levels:
                matches[level] = HierarchyMatch(
                    success=False,
                    level=level,
                    error=f"Level '{level}' not available in hierarchy report",
                )
                continue
            cols = self._column_map[level]
            uuid_val = self._cell(matched_row, cols.get("id"))
            name_val = self._cell(matched_row, cols.get("name"))
            code_val = self._cell(matched_row, cols.get("code"))
            requested_code = hierarchy_keys.get(f"{level}_code", "").strip()
            requested_name = hierarchy_keys.get(f"{level}_name", "").strip()
            mismatches = []
            if requested_code and requested_code != code_val:
                mismatches.append((level, "code", requested_code, code_val))
            if requested_name and requested_name.casefold() != name_val.casefold():
                mismatches.append((level, "name", requested_name, name_val))
            if mismatches and any(
                mismatch not in reconciled_conflicts for mismatch in mismatches
            ):
                matches[level] = HierarchyMatch(
                    success=False,
                    level=level,
                    error=(
                        f"Source {level} identifier does not agree with the "
                        "matched hierarchy row"
                    ),
                )
            elif uuid_val and is_valid_uuid(uuid_val):
                matches[level] = HierarchyMatch(
                    success=True,
                    level=level,
                    uuid=normalise_uuid(uuid_val),
                    name=name_val,
                    code=code_val,
                    match_count=1,
                )
            else:
                matches[level] = HierarchyMatch(
                    success=False,
                    level=level,
                    error=(
                        f"UUID for level '{level}' is missing or invalid in "
                        "the matched hierarchy row"
                    ),
                )

        parent_issues = self._validate_parents(
            hierarchy_keys,
            matches,
            ignored_identifiers={
                (level, kind) for level, kind, _, _ in reconciled_conflicts
            },
        )
        warnings.extend(parent_issues)
        return HierarchyRecord(
            matches=matches,
            errors=errors,
            warnings=warnings,
            matched_row_index=matched_idx,
            requires_review=requires_review,
            suggested_corrections=[
                {
                    "level": level,
                    "kind": kind,
                    "source_value": source_value,
                    "suggested_value": matched_value,
                }
                for level, kind, source_value, matched_value in reconciled_conflicts
            ],
        )

    def _reconcile_single_conflict(
        self,
        level: str,
        hierarchy_keys: Dict[str, str],
    ) -> Optional[Tuple[str, List[int], List[Tuple[str, str, str, str]]]]:
        """Find a unique path when one supplied name/code conflicts but its pair and path corroborate the other value."""
        if self._df is None or level not in self._available_levels:
            return None
        cols = self._column_map[level]
        source_code = hierarchy_keys.get(f"{level}_code", "").strip()
        source_name = hierarchy_keys.get(f"{level}_name", "").strip()
        candidate_rows: set = set()
        if source_code and cols.get("code"):
            candidate_rows.update(self._code_indices.get(level, {}).get(source_code, []))
        if source_name and cols.get("name"):
            candidate_rows.update(
                self._name_indices.get(level, {}).get(source_name.casefold(), [])
            )
        if not candidate_rows:
            return None

        level_index = HIERARCHY_LEVELS.index(level)
        matched_rows: List[int] = []
        conflicts_by_row: Dict[int, List[Tuple[str, str, str, str]]] = {}
        for row_index in candidate_rows:
            row = row_index
            mismatches: List[Tuple[str, str, str, str]] = []
            comparable = True
            for check_level in HIERARCHY_LEVELS[: level_index + 1]:
                check_cols = self._column_map.get(check_level, {})
                for kind in ("code", "name"):
                    source_value = hierarchy_keys.get(f"{check_level}_{kind}", "").strip()
                    if not source_value:
                        continue
                    column = check_cols.get(kind)
                    if not column:
                        comparable = False
                        break
                    hierarchy_value = self._cell(row, column)
                    equal = (
                        source_value == hierarchy_value
                        if kind == "code"
                        else source_value.casefold() == hierarchy_value.casefold()
                    )
                    if not equal:
                        mismatches.append((check_level, kind, source_value, hierarchy_value))
                if not comparable:
                    break
            if not comparable or len(mismatches) != 1:
                continue

            conflict_level, conflict_kind, _, _ = mismatches[0]
            companion_kind = "name" if conflict_kind == "code" else "code"
            companion_source = hierarchy_keys.get(
                f"{conflict_level}_{companion_kind}", ""
            ).strip()
            companion_column = self._column_map.get(conflict_level, {}).get(companion_kind)
            if not companion_source or not companion_column:
                continue
            companion_value = self._cell(row, companion_column)
            companion_matches = (
                companion_source.casefold() == companion_value.casefold()
                if companion_kind == "name"
                else companion_source == companion_value
            )
            if not companion_matches:
                continue
            matched_rows.append(row_index)
            conflicts_by_row[row_index] = mismatches

        if not matched_rows:
            return None
        unique_rows = self._unique_indices(level, matched_rows)
        if len(unique_rows) != 1:
            return "ambiguous", unique_rows or matched_rows, []
        entity_id = self._cell(unique_rows[0], cols.get("id"))
        if not entity_id or not self._same_ancestor_path(level, matched_rows, entity_id):
            return "ambiguous", unique_rows, []
        return "unique", unique_rows, conflicts_by_row[matched_rows[0]]

    def _match_level(
        self,
        level: str,
        hierarchy_keys: Dict[str, str],
    ) -> Optional[Tuple[str, List[int]]]:
        source_code = hierarchy_keys.get(f"{level}_code", "").strip()
        source_name = hierarchy_keys.get(f"{level}_name", "").strip()
        if not source_code and not source_name:
            return None
        if self._df is None or level not in self._available_levels:
            return ("missing", [])

        cols = self._column_map[level]
        candidates: Optional[List[int]] = None
        for key, column, kind in (
            (source_code, cols.get("code"), "code"),
            (source_name, cols.get("name"), "name"),
        ):
            if not key:
                continue
            # A supplied identifier is evidence only if the hierarchy report
            # contains the corresponding field and its exact value matches.
            if not column:
                return ("missing", [])
            if kind == "code":
                found = self._code_indices.get(level, {}).get(key, [])
            else:
                found = self._name_indices.get(level, {}).get(key.casefold(), [])
            if candidates is None:
                candidates = list(found)
            else:
                found_set = set(found)
                candidates = [idx for idx in candidates if idx in found_set]
        if not candidates:
            return ("missing", [])

        narrowed = self._narrow_indices(candidates, hierarchy_keys, level)
        unique_n = self._unique_indices(level, narrowed)
        if len(unique_n) == 1:
            # Repeated flattened rows for the same entity are safe only when
            # their own and ancestor identifiers agree. Otherwise the report
            # does not establish one unambiguous hierarchy path.
            entity_id_col = cols.get("id")
            entity_id = self._cell(unique_n[0], entity_id_col)
            if entity_id and not self._same_ancestor_path(level, narrowed, entity_id):
                return ("ambiguous", narrowed)
            return ("unique", unique_n)
        if len(unique_n) > 1 or (len(narrowed) > 1 and not cols.get("id")):
            return ("ambiguous", unique_n or narrowed)
        return ("missing", [])

    def _unique_indices(self, level: str, idxs: List[int]) -> List[int]:
        if self._df is None or not idxs:
            return idxs
        id_col = self._column_map.get(level, {}).get("id")
        if not id_col or id_col not in self._df.columns:
            return idxs
        seen: Dict[str, int] = {}
        for i in idxs:
            uid = self._cell(i, id_col).lower()
            identity = self._identity_key(uid)
            if identity and identity not in seen:
                seen[identity] = i
        return list(seen.values())

    def _narrow_indices(
        self,
        idxs: List[int],
        hierarchy_keys: Dict[str, str],
        current_level: str,
    ) -> List[int]:
        if self._df is None or current_level not in HIERARCHY_LEVELS:
            return idxs
        current_idx = HIERARCHY_LEVELS.index(current_level)
        remaining = list(idxs)
        for parent_level in HIERARCHY_LEVELS[:current_idx]:
            parent_code = hierarchy_keys.get(f"{parent_level}_code", "").strip()
            parent_name = hierarchy_keys.get(f"{parent_level}_name", "").strip()
            if parent_level not in self._column_map:
                if parent_code or parent_name:
                    return []
                continue
            parent_code_col = self._column_map[parent_level].get("code")
            parent_name_col = self._column_map[parent_level].get("name")
            # Every supplied parent identifier constrains the candidate set.
            # An unavailable or conflicting parent column must not be ignored.
            if parent_code:
                if not parent_code_col:
                    return []
                remaining = [
                    i for i in remaining
                    if self._cell(i, parent_code_col) == parent_code
                ]
            if parent_name:
                if not parent_name_col:
                    return []
                remaining = [
                    i for i in remaining
                    if self._cell(i, parent_name_col).upper()
                    == parent_name.upper()
                ]
            if not remaining:
                return []
        return remaining

    def _same_ancestor_path(self, level: str, idxs: List[int], entity_id: str) -> bool:
        """Return whether repeated rows for one entity agree through that level."""
        if self._df is None:
            return False
        id_col = self._column_map.get(level, {}).get("id")
        if not id_col:
            return len(idxs) == 1
        entity_rows = [
            i for i in idxs
            if self._identity_key(self._cell(i, id_col))
            == self._identity_key(entity_id)
        ]
        level_index = HIERARCHY_LEVELS.index(level)
        for path_level in HIERARCHY_LEVELS[: level_index + 1]:
            path_cols = self._column_map.get(path_level, {})
            for kind in ("id", "code", "name"):
                column = path_cols.get(kind)
                if not column:
                    continue
                raw_values = [self._cell(i, column) for i in entity_rows]
                values = {
                    self._identity_key(value) if kind == "id" else value.casefold()
                    for value in raw_values
                }
                if len(values) > 1:
                    return False
        return bool(entity_rows)

    def _identity_key(self, value: str) -> str:
        return normalise_uuid(value).casefold() if is_valid_uuid(value) else value.casefold()

    def _validate_parents(
        self,
        hierarchy_keys: Dict[str, str],
        matches: Dict[str, HierarchyMatch],
        ignored_identifiers: Optional[set] = None,
    ) -> List[str]:
        issues: List[str] = []
        ignored_identifiers = ignored_identifiers or set()
        for level, match in matches.items():
            if not match.success:
                continue
            src_code = hierarchy_keys.get(f"{level}_code", "").strip()
            if (
                src_code
                and match.code
                and src_code != match.code
                and (level, "code") not in ignored_identifiers
            ):
                issues.append(
                    f"Parent/child check: source {level} code '{src_code}' "
                    f"differs from matched '{match.code}'"
                )
        return issues

    def _cell(self, row: Any, col: Optional[str]) -> str:
        if not col:
            return ""
        if isinstance(row, int):
            values = self._column_values.get(col)
            if values is None or row < 0 or row >= len(values):
                return ""
            val = values[row]
        elif col in row.index:
            val = row[col]
        else:
            return ""
        if pd.isna(val):
            return ""
        return str(val).strip()

    def _empty_record(
        self,
        errors: List[str],
        warnings: List[str],
        ambiguous_level: str = "",
        match_count: int = 0,
        requires_review: bool = False,
    ) -> HierarchyRecord:
        matches = {}
        for level in HIERARCHY_LEVELS:
            matches[level] = HierarchyMatch(
                success=False,
                level=level,
                error=errors[0] if errors else "Hierarchy not resolved",
                is_ambiguous=level == ambiguous_level,
                match_count=match_count,
            )
        return HierarchyRecord(
            matches=matches,
            errors=errors,
            warnings=warnings,
            requires_review=requires_review,
        )

    def get_available_levels(self) -> List[str]:
        return self._available_levels.copy()

    def get_duplicate_report(self) -> Dict[str, List[str]]:
        return self._duplicates.copy()

    def validate_hierarchy_columns(self) -> List[str]:
        warnings = []
        required_with_id = ["zone", "circle", "division", "subdivision"]
        for level in required_with_id:
            if level not in self._available_levels:
                warnings.append(f"Hierarchy level '{level}' not detected in report")
            elif "id" not in self._column_map.get(level, {}):
                warnings.append(
                    f"Hierarchy level '{level}' found but UUID/ID column not detected"
                )
        return warnings
