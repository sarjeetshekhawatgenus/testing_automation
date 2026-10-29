from master_update_automation.column_mapper import create_column_mapping


def test_known_source_headers_map_to_template():
    source = [
        "Consumer Number", "Account Number", "Consumer Name",
        "Consumer Identification", "Meter Make", "DT Code", "Extra Column",
    ]
    target = [
        "consumer_number", "account_number", "consumer_name",
        "consumer_identy", "meter_make", "dt_code",
    ]
    result = create_column_mapping(source, target)
    mapped = {m.source_column: m.target_column for m in result.mappings}
    assert mapped["Consumer Number"] == "consumer_number"
    assert mapped["Consumer Identification"] == "consumer_identy"
    assert mapped["DT Code"].startswith("_hierarchy_")
    assert "Extra Column" in result.unmapped_source


def test_missing_mandatory_source_columns_reported():
    result = create_column_mapping(["Foo"], ["consumer_number", "consumer_name"])
    assert "consumer_number" in result.unmapped_target
    assert "consumer_name" in result.unmapped_target


def test_alias_columns_for_one_target_are_not_selected_by_header_order():
    result = create_column_mapping(
        ["Consumer UAN", "Consumer UAN Suffix"],
        ["consumer_uan_suffix"],
    )
    assert not result.mappings
    assert result.ambiguous_mappings == [
        ("consumer_uan_suffix", ["Consumer UAN", "Consumer UAN Suffix"])
    ]
