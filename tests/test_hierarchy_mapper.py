from master_update_automation.hierarchy_mapper import HierarchyMapper
from tests.conftest import DT_ID, FEEDER_ID


def test_unique_dt_code_match(hierarchy_df):
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer({}, {"dt_code": "1085140006"})
    assert rec.matches["dt"].success
    assert rec.matches["dt"].uuid == f"{{{DT_ID}}}"
    assert rec.matches["subdivision"].success


def test_missing_hierarchy_match(hierarchy_df):
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer({}, {"dt_code": "DOES-NOT-EXIST"})
    assert not rec.matches["dt"].success
    assert rec.errors or not rec.matches["dt"].success


def test_ambiguous_hierarchy_match_not_guessed(hierarchy_df):
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer({}, {"dt_code": "DUPCODE"})
    assert rec.matches["dt"].is_ambiguous or any("Ambiguous" in e for e in rec.errors)
    assert not rec.matches["dt"].success
    assert rec.matches["dt"].uuid == ""


def test_feeder_match_does_not_choose_a_child_dt_from_flattened_rows(hierarchy_df):
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer({}, {"feeder_code": "1085140"})
    assert rec.matches["feeder"].success
    assert rec.matches["feeder"].uuid == f"{{{FEEDER_ID}}}"
    assert not rec.matches["dt"].success
    assert rec.matches["dt"].uuid == ""


def test_conflicting_parent_identifier_prevents_hierarchy_match(hierarchy_df):
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer(
        {},
        {"dt_code": "1085140006", "discom_code": "WRONG-DISCOM"},
    )
    assert not rec.matches["dt"].success
    assert rec.matches["dt"].uuid == ""


def test_invalid_hierarchy_uuid_is_not_populated(hierarchy_df):
    hierarchy_df.loc[0, "DT ID"] = "NOT-A-UUID"
    mapper = HierarchyMapper()
    mapper.load_hierarchy(hierarchy_df)
    rec = mapper.match_consumer({}, {"dt_code": "1085140006"})
    assert not rec.matches["dt"].success
    assert "invalid" in rec.matches["dt"].error.lower()
