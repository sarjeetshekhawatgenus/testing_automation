"""Resolve business values to UUIDs from ref_project_masters_lov."""

from __future__ import annotations

import logging
from collections import OrderedDict
from dataclasses import dataclass, replace
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .config import (
    Category,
    METER_MAKE_REFERENCE,
    REFERENCE_CATEGORY_ALIASES,
    REFERENCE_CATEGORY_MAP,
    REFERENCE_SHEET_NAME,
)
from .uuid_utils import is_valid_uuid, normalise_uuid

logger = logging.getLogger(__name__)


@dataclass
class ReferenceEntry:
    id: str
    name: str
    code: str
    master_name: str


@dataclass
class LookupResult:
    success: bool
    uuid: str = ""
    name: str = ""
    error: str = ""
    category: str = ""
    search_value: str = ""
    is_ambiguous: bool = False
    is_invalid: bool = False


class ReferenceLookup:
    def __init__(self) -> None:
        self._entries: List[ReferenceEntry] = []
        self._by_category: Dict[str, List[ReferenceEntry]] = {}
        self._by_name: Dict[Tuple[str, str], List[ReferenceEntry]] = {}
        self._by_code: Dict[Tuple[str, str], List[ReferenceEntry]] = {}
        self._duplicate_keys: List[str] = []
        self._lookup_cache: OrderedDict = OrderedDict()
        self._lookup_cache_limit = 100_000

    def load_from_workbook(self, wb) -> None:
        if REFERENCE_SHEET_NAME not in wb.sheetnames:
            raise ValueError(
                f"Reference sheet '{REFERENCE_SHEET_NAME}' not found in workbook. "
                f"Available sheets: {wb.sheetnames}"
            )
        ws = wb[REFERENCE_SHEET_NAME]
        rows = []
        headers = [c.value for c in next(ws.iter_rows(max_row=1))]
        col_index = {
            str(h).strip().lower(): i
            for i, h in enumerate(headers) if h
        }
        id_i = col_index.get("id", 0)
        name_i = col_index.get("name", 1)
        code_i = col_index.get("code", 2)
        master_i = col_index.get("master_name", 3)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None:
                continue
            values = list(row)
            def _cell(i: int) -> str:
                if i >= len(values) or values[i] is None:
                    return ""
                return str(values[i]).strip()
            rows.append({
                "id": _cell(id_i),
                "name": _cell(name_i),
                "code": _cell(code_i),
                "master_name": _cell(master_i),
            })
        self._load_rows(rows)

    def load_from_dataframe(self, df: pd.DataFrame) -> None:
        def cell_text(value) -> str:
            if value is None or pd.isna(value):
                return ""
            return str(value).strip()

        rows = []
        for row in df.to_dict(orient="records"):
            rows.append({
                "id": cell_text(row.get("id", "")),
                "name": cell_text(row.get("name", "")),
                "code": cell_text(row.get("code", "")),
                "master_name": cell_text(row.get("master_name", "")),
            })
        self._load_rows(rows)

    def _load_rows(self, rows: List[dict]) -> None:
        self._entries.clear()
        self._by_category.clear()
        self._by_name.clear()
        self._by_code.clear()
        self._duplicate_keys.clear()
        self._lookup_cache.clear()

        for row in rows:
            entry_id = row.get("id", "")
            master_name = row.get("master_name", "")
            if not master_name:
                continue
            entry = ReferenceEntry(
                id=entry_id,
                name=row.get("name", ""),
                code=row.get("code", ""),
                master_name=master_name,
            )
            self._entries.append(entry)
            cat_key = entry.master_name.upper()
            self._by_category.setdefault(cat_key, []).append(entry)
            if entry.name:
                self._by_name.setdefault((cat_key, entry.name.upper()), []).append(entry)
            if entry.code:
                for cp in [c.strip().upper() for c in entry.code.split(",")]:
                    if cp:
                        self._by_code.setdefault((cat_key, cp), []).append(entry)

        for key, entries in self._by_name.items():
            ids = {e.id.lower() for e in entries}
            if len(ids) > 1:
                self._duplicate_keys.append(f"{key[0]} / name '{key[1]}'")
        for key, entries in self._by_code.items():
            ids = {e.id.lower() for e in entries}
            if len(ids) > 1:
                self._duplicate_keys.append(f"{key[0]} / code '{key[1]}'")

        logger.info(
            "Loaded %s reference entries across %s categories",
            len(self._entries),
            len(self._by_category),
        )

    def get_categories(self) -> List[str]:
        return sorted(self._by_category.keys())

    def get_duplicate_report(self) -> List[str]:
        return list(self._duplicate_keys)

    def get_entries_for_category(self, category: str) -> List[ReferenceEntry]:
        return self._by_category.get(category.upper(), [])

    def _unique_entry(self, entries: List[ReferenceEntry], category: str, value: str) -> LookupResult:
        invalid_entries = [entry for entry in entries if not is_valid_uuid(entry.id)]
        if invalid_entries:
            return LookupResult(
                success=False,
                is_invalid=True,
                error=(
                    f"Reference value '{value}' in category '{category}' has "
                    "a missing or invalid UUID"
                ),
                category=category,
                search_value=value,
            )
        ids = {normalise_uuid(e.id).lower() for e in entries}
        if len(ids) > 1:
            return LookupResult(
                success=False,
                is_ambiguous=True,
                error=(
                    f"Ambiguous reference for '{value}' in category '{category}': "
                    f"{len(ids)} distinct UUIDs"
                ),
                category=category,
                search_value=value,
            )
        # All remaining candidates carry the same valid UUID, so choosing a
        # representative cannot change the output identifier.
        entry = entries[0]
        return LookupResult(
            success=True,
            uuid=normalise_uuid(entry.id),
            name=entry.name,
            category=category,
            search_value=value,
        )

    def lookup_by_name(self, category: str, name: str) -> LookupResult:
        cache_key = ("name", str(category or "").upper(), str(name or "").strip().upper())
        cached = self._lookup_cache.get(cache_key)
        if cached is not None:
            self._lookup_cache.move_to_end(cache_key)
            return replace(cached)
        result = self._lookup_by_name_uncached(category, name)
        self._cache_lookup_result(cache_key, result)
        return replace(result)

    def _lookup_by_name_uncached(self, category: str, name: str) -> LookupResult:
        if not name or not category:
            return LookupResult(
                success=False,
                error=f"Empty lookup: category='{category}', name='{name}'",
                category=category,
                search_value=name or "",
            )
        cat_upper = category.upper()
        name_upper = str(name).strip().upper()
        # Prefer entries under the exact requested category. A template may
        # contain both a canonical category and a misspelled legacy category;
        # combining them can make a valid value (for example, "Yes") appear
        # ambiguous when the two rows have different UUIDs.
        name_entries = self._by_name.get((cat_upper, name_upper), [])
        if name_entries:
            return self._unique_entry(name_entries, category, name)
        code_entries = self._by_code.get((cat_upper, name_upper), [])
        if code_entries:
            return self._unique_entry(code_entries, category, name)

        # Fall back to alternate category spellings only when the canonical
        # category has no matching name or code.
        alias_categories = [
            alias.upper()
            for alias in REFERENCE_CATEGORY_ALIASES.get(cat_upper, [])
        ]
        name_entries = [
            entry
            for alias in alias_categories
            for entry in self._by_name.get((alias, name_upper), [])
        ]
        if name_entries:
            return self._unique_entry(name_entries, category, name)
        code_entries = [
            entry
            for alias in alias_categories
            for entry in self._by_code.get((alias, name_upper), [])
        ]
        if code_entries:
            return self._unique_entry(code_entries, category, name)

        return LookupResult(
            success=False,
            error=f"Value '{name}' not found in category '{category}'",
            category=category,
            search_value=name,
        )

    def lookup_by_code(self, category: str, code: str) -> LookupResult:
        cache_key = ("code", str(category or "").upper(), str(code or "").strip().upper())
        cached = self._lookup_cache.get(cache_key)
        if cached is not None:
            self._lookup_cache.move_to_end(cache_key)
            return replace(cached)
        result = self._lookup_by_code_uncached(category, code)
        self._cache_lookup_result(cache_key, result)
        return replace(result)

    def _lookup_by_code_uncached(self, category: str, code: str) -> LookupResult:
        if not code or not category:
            return LookupResult(
                success=False,
                error=f"Empty lookup: category='{category}', code='{code}'",
                category=category,
                search_value=code or "",
            )
        cat_upper = category.upper()
        code_upper = str(code).strip().upper()
        entries = self._by_code.get((cat_upper, code_upper), [])
        if entries:
            return self._unique_entry(entries, category, code)
        return LookupResult(
            success=False,
            error=f"Code '{code}' not found in category '{category}'",
            category=category,
            search_value=code,
        )

    def _cache_lookup_result(
        self,
        key: Tuple[str, str, str],
        result: LookupResult,
    ) -> None:
        self._lookup_cache[key] = replace(result)
        if len(self._lookup_cache) > self._lookup_cache_limit:
            self._lookup_cache.popitem(last=False)

    def lookup_meter_make(self, dest_category: Category, meter_make_name: str) -> LookupResult:
        ref_category = METER_MAKE_REFERENCE.get(dest_category)
        if not ref_category:
            return LookupResult(
                success=False,
                error=f"No meter make reference defined for category '{dest_category.value}'",
                category="METER MAKE",
                search_value=meter_make_name,
            )
        return self.lookup_by_name(ref_category, meter_make_name)

    def lookup_for_target_field(
        self,
        target_field: str,
        value: str,
        dest_category: Optional[Category] = None,
    ) -> LookupResult:
        target_lower = target_field.lower().strip()
        if target_lower == "meter_make":
            if dest_category is None:
                return LookupResult(
                    success=False,
                    error="Destination category required for meter_make lookup",
                    category="METER MAKE",
                    search_value=value,
                )
            return self.lookup_meter_make(dest_category, value)
        ref_category = REFERENCE_CATEGORY_MAP.get(target_lower)
        if not ref_category:
            return LookupResult(
                success=False,
                error=f"No reference category configured for field '{target_field}'",
                category="",
                search_value=value,
            )
        return self.lookup_by_name(ref_category, value)

    def category_exists(self, category: str) -> bool:
        if category.upper() in self._by_category:
            return True
        for alias in REFERENCE_CATEGORY_ALIASES.get(category.upper(), []):
            if alias.upper() in self._by_category:
                return True
        return False

    def validate_categories(self, required_categories: List[str]) -> List[str]:
        return [cat for cat in required_categories if not self.category_exists(cat)]
