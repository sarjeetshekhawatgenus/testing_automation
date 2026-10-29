import pandas as pd

from master_update_automation.config import Category
from master_update_automation.reference_lookup import ReferenceLookup
from tests.conftest import GENUS_HTCT, MAIN_METER_ID, TOD_ID, reference_frame


def test_missing_reference_value():
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(reference_frame())
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "UNKNOWN")
    assert not result.success
    assert result.uuid == ""
    assert "UNKNOWN" in result.error


def test_duplicate_reference_value_is_ambiguous():
    df = reference_frame()
    extra = pd.DataFrame([{
        "id": "00000000-0000-0000-0000-000000000099",
        "name": "TOD",
        "code": "99",
        "master_name": "CONSUMER IDENTIFICATION",
    }])
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(pd.concat([df, extra], ignore_index=True))
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "TOD")
    assert not result.success
    assert result.is_ambiguous


def test_invalid_reference_uuid_is_not_returned():
    df = reference_frame()
    df.loc[df["name"] == "TOD", "id"] = "NOT-A-UUID"
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(df)
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "TOD")
    assert not result.success
    assert result.uuid == ""
    assert result.is_invalid


def test_missing_reference_uuid_is_not_hidden_when_another_row_matches():
    df = reference_frame()
    extra = pd.DataFrame([{
        "id": "",
        "name": "TOD",
        "code": "99",
        "master_name": "CONSUMER IDENTIFICATION",
    }])
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(pd.concat([df, extra], ignore_index=True))
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "TOD")
    assert not result.success
    assert result.uuid == ""
    assert result.is_invalid


def test_valid_and_invalid_conflicting_reference_rows_are_not_arbitrated():
    df = reference_frame()
    extra = pd.DataFrame([{
        "id": "NOT-A-UUID",
        "name": "TOD",
        "code": "99",
        "master_name": "CONSUMER IDENTIFICATION",
    }])
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(pd.concat([df, extra], ignore_index=True))
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "TOD")
    assert not result.success
    assert result.uuid == ""
    assert result.is_invalid


def test_meter_make_does_not_cross_categories():
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(reference_frame())
    htct = lookup.lookup_meter_make(Category.HTCT, "Genus")
    wc = lookup.lookup_meter_make(Category.WC, "Genus")
    assert htct.uuid == f"{{{GENUS_HTCT}}}"
    assert htct.uuid != wc.uuid
    missing = lookup.lookup_meter_make(Category.HTCT, "Secure")
    assert not missing.success


def test_tod_uuid_comes_from_reference_not_hardcode():
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(reference_frame())
    result = lookup.lookup_by_name("CONSUMER IDENTIFICATION", "TOD")
    assert result.uuid == f"{{{TOD_ID}}}"
    assert result.uuid.startswith("{") and result.uuid.endswith("}")


def test_main_meter_uuid_comes_from_reference_not_hardcode():
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(reference_frame())
    result = lookup.lookup_for_target_field("meter_category", "Main Meter")
    assert result.success
    assert result.uuid == f"{{{MAIN_METER_ID}}}"
    assert result.uuid.startswith("{") and result.uuid.endswith("}")
    assert "{{" not in result.uuid


def test_main_meter_missing_from_reference():
    df = reference_frame()
    df = df[~((df["name"] == "Main Meter") & (df["master_name"] == "METER CATEGORY"))]
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(df)
    result = lookup.lookup_for_target_field("meter_category", "Main Meter")
    assert not result.success
    assert result.uuid == ""
    assert "Main Meter" in result.error
