"""Match source records to existing target records for insert/update."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class MatchResult:
    source_index: int
    business_key: str
    action: str
    matched_target_index: Optional[int] = None
    error: str = ""


def detect_duplicate_keys(df: pd.DataFrame, key_column: str) -> Dict[str, List[int]]:
    duplicates: Dict[str, List[int]] = {}
    if key_column not in df.columns:
        return duplicates
    for idx, raw_key in zip(df.index, df[key_column].tolist()):
        key = str(raw_key or "").strip()
        if key:
            duplicates.setdefault(key, []).append(idx)
    return {k: v for k, v in duplicates.items() if len(v) > 1}


def match_records(
    source_df: pd.DataFrame,
    target_df: Optional[pd.DataFrame],
    source_key_column: str,
    target_key_column: Optional[str] = None,
) -> List[MatchResult]:
    if target_key_column is None:
        target_key_column = source_key_column
    results: List[MatchResult] = []
    target_index: Dict[str, List[int]] = {}
    if target_df is not None and target_key_column in target_df.columns:
        for idx, raw_key in zip(target_df.index, target_df[target_key_column].tolist()):
            key = str(raw_key or "").strip()
            if key:
                target_index.setdefault(key, []).append(idx)

    source_keys_seen: Dict[str, int] = {}
    if source_key_column in source_df.columns:
        source_keys = source_df[source_key_column].tolist()
    else:
        source_keys = [""] * len(source_df)
    for idx, raw_key in zip(source_df.index, source_keys):
        key = str(raw_key or "").strip()
        if not key:
            results.append(MatchResult(
                source_index=idx,
                business_key="",
                action="FAILED",
                error=f"Missing business key in column '{source_key_column}'",
            ))
            continue
        if key in source_keys_seen:
            results.append(MatchResult(
                source_index=idx,
                business_key=key,
                action="DUPLICATE",
                error=f"Duplicate source key '{key}' (first seen at row {source_keys_seen[key]})",
            ))
            continue
        source_keys_seen[key] = idx
        if target_df is None or not target_index:
            results.append(MatchResult(
                source_index=idx,
                business_key=key,
                action="INSERT",
            ))
            continue
        target_matches = target_index.get(key, [])
        if len(target_matches) == 0:
            results.append(MatchResult(source_index=idx, business_key=key, action="INSERT"))
        elif len(target_matches) == 1:
            results.append(MatchResult(
                source_index=idx,
                business_key=key,
                action="UPDATE",
                matched_target_index=target_matches[0],
            ))
        else:
            results.append(MatchResult(
                source_index=idx,
                business_key=key,
                action="AMBIGUOUS",
                error=f"Multiple target matches for key '{key}': {len(target_matches)} found",
            ))
    return results
