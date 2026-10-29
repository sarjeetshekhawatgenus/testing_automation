"""Shared fixtures for unit tests. Sample UUIDs here are fixture-only, not application defaults."""

from __future__ import annotations

import pandas as pd
import pytest
from openpyxl import Workbook

from master_update_automation.config import Category, Operation
from master_update_automation.reference_lookup import ReferenceLookup


TOD_ID = "f738fdf0-3a5b-4512-bb9b-0ec238884edb"
ESC_ID = "6058cd3a-2070-4494-8997-516e35a647d3"
NSC_ID = "53b06e5f-dec4-4a9e-8b4f-fc04660b8afc"
LEC_ID = "bc24b70e-e7d3-40f4-b3ac-77a38099e219"
MAIN_METER_ID = "ed1f2293-af64-4346-978b-324d5998f539"
GENUS_HTCT = "db764fe0-dce5-4cb3-a3a3-7cc34ba2d05d"
GENUS_LTCT = "7841b49b-f68b-4f0d-842d-b8be16522b73"
GENUS_WC = "9d679ccf-f26e-4bad-8634-6c226c3aeb79"
SECURE_LTCT = "d4087218-46a6-4388-9e1e-be1c88439556"
DT_ID = "11111111-1111-1111-1111-111111111111"
FEEDER_ID = "22222222-2222-2222-2222-222222222222"
SUBDIV_ID = "33333333-3333-3333-3333-333333333333"
DIV_ID = "44444444-4444-4444-4444-444444444444"
CIRCLE_ID = "55555555-5555-5555-5555-555555555555"
ZONE_ID = "66666666-6666-6666-6666-666666666666"
DISCOM_ID = "77777777-7777-7777-7777-777777777777"
IS_OLD_MTR_YES_ID = "9a2b42b2-44d3-4f25-8d1d-432b316f0bb3"
IS_OLD_MTR_NO_ID = "4e29901b-fdd8-4130-8dfd-01928c436e7a"
AMBIG_DT_A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
AMBIG_DT_B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"


def reference_frame() -> pd.DataFrame:
    rows = [
        (ESC_ID, "ESC", "1", "CONSUMER IDENTIFICATION"),
        (TOD_ID, "TOD", "3", "CONSUMER IDENTIFICATION"),
        (NSC_ID, "NSC", "2", "CONSUMER IDENTIFICATION"),
        (LEC_ID, "LEC", "4", "CONSUMER IDENTIFICATION"),
        (GENUS_WC, "Genus", "1", "WC METER MAKE"),
        (GENUS_LTCT, "Genus", "1", "LTCT METER MAKE"),
        (SECURE_LTCT, "Secure", "3", "LTCT METER MAKE"),
        (GENUS_HTCT, "Genus", "1", "HTCT METER MAKE"),
        (MAIN_METER_ID, "Main Meter", "1", "METER CATEGORY"),
        ("d4700e36-e076-4c66-92b1-d8ecb1df6648", "Regular", "R", "CONSUMER STATUS"),
        ("11737692-ff66-47ab-a4b2-1c682c28c35b", "3-Phase", "3,3 Phase,3 PH", "PHASE TYPE"),
        ("d0dc8abf-ae82-4470-be5b-d29535b4c937", "Rural", "2", "GEOGRAPHICAL AREA"),
        ("404eca77-d1b0-4799-81fd-7a18506b1c27", "ESC", "1", "CONSUMER CONNECTION TYPE"),
        ("9a2b42b2-44d3-4f25-8d1d-432b316f0bb3", "Yes", "1", "IS OLD METER AVAILABLE"),
        ("4e29901b-fdd8-4130-8dfd-01928c436e7a", "No", "2", "IS OLD METER AVAILABLE"),
        ("5339df35-f8d6-44ad-a362-644260ad888e", "OK", "R", "METER STATUS"),
        ("1f373a1b-acab-4c86-a060-956d7f89a097", "HP", "2", "SANCTIONED LOAD UNIT"),
        ("a33c1350-f011-4620-972c-ae96be5fd1f5", "Forward Mode", "Fwd", "METERING MODE"),
    ]
    return pd.DataFrame(rows, columns=["id", "name", "code", "master_name"])


@pytest.fixture
def ref_lookup() -> ReferenceLookup:
    lookup = ReferenceLookup()
    lookup.load_from_dataframe(reference_frame())
    return lookup


@pytest.fixture
def hierarchy_df() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Discom ID": DISCOM_ID, "Discom Name": "JVVNL", "Discom Code": "GJV96",
            "Zone ID": ZONE_ID, "Zone Name": "Jaipur", "Zone Code": "JPR",
            "Circle ID": CIRCLE_ID, "Circle Name": "KOTPUTLI", "Circle Code": "2117000",
            "Division ID": DIV_ID, "Division Name": "XEN_OM_KOTPUTALI", "Division Code": "2117200",
            "Subdivision ID": SUBDIV_ID, "Subdivision Name": "AEN_OM_VIRAT NAGAR", "Subdivision Code": "2117240",
            "Substation ID": "88888888-8888-8888-8888-888888888888",
            "Substation Name": "AANTELA", "Substation Code": "1099",
            "Feeder ID": FEEDER_ID, "Feeder Name": "ODA CRUSHER", "Feeder Code": "1085140",
            "DT ID": DT_ID, "DT Name": "GOVINDAM MINERALS", "DT Code": "1085140006",
        },
        {
            "Discom ID": DISCOM_ID, "Discom Name": "JVVNL", "Discom Code": "GJV96",
            "Zone ID": ZONE_ID, "Zone Name": "Jaipur", "Zone Code": "JPR",
            "Circle ID": CIRCLE_ID, "Circle Name": "KOTPUTLI", "Circle Code": "2117000",
            "Division ID": DIV_ID, "Division Name": "XEN_OM_KOTPUTALI", "Division Code": "2117200",
            "Subdivision ID": SUBDIV_ID, "Subdivision Name": "AEN_OM_VIRAT NAGAR", "Subdivision Code": "2117240",
            "Substation ID": "88888888-8888-8888-8888-888888888888",
            "Substation Name": "AANTELA", "Substation Code": "1099",
            "Feeder ID": FEEDER_ID, "Feeder Name": "ODA CRUSHER", "Feeder Code": "1085140",
            "DT ID": AMBIG_DT_A, "DT Name": "DT A", "DT Code": "DUPCODE",
        },
        {
            "Discom ID": DISCOM_ID, "Discom Name": "JVVNL", "Discom Code": "GJV96",
            "Zone ID": ZONE_ID, "Zone Name": "Jaipur", "Zone Code": "JPR",
            "Circle ID": CIRCLE_ID, "Circle Name": "KOTPUTLI", "Circle Code": "2117000",
            "Division ID": DIV_ID, "Division Name": "XEN_OM_KOTPUTALI", "Division Code": "2117200",
            "Subdivision ID": "99999999-9999-9999-9999-999999999999",
            "Subdivision Name": "OTHER SUB", "Subdivision Code": "0000001",
            "Substation ID": "88888888-8888-8888-8888-888888888888",
            "Substation Name": "OTHER SS", "Substation Code": "2000",
            "Feeder ID": "12121212-1212-1212-1212-121212121212",
            "Feeder Name": "OTHER FDR", "Feeder Code": "2000001",
            "DT ID": AMBIG_DT_B, "DT Name": "DT B", "DT Code": "DUPCODE",
        },
    ])


def make_template_workbook(path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "HTCT Consumer Master"
    headers = [
        "id", "discom_name", "zone_name", "zone_code", "circle_name", "circle_code",
        "division_name", "division_code", "subdivision_name", "subdivision_code",
        "substation_name", "substation_code", "feeder_name", "feeder_code",
        "dt_name", "dt_code", "old_consumer_number", "consumer_number",
        "account_number", "consumer_uan_suffix", "consumer_name", "father_name",
        "consumer_address", "mobile_number", "email", "tariff_code",
        "meter_category", "connection_category", "consumer_status", "connected_load",
        "sanctioned_load_unit", "sanctioned_load", "meter_phase", "meter_serial_no",
        "meter_make", "meter_status", "metering_voltage", "supply_voltage",
        "contract_demand", "mf", "geographical_are", "cons_available_mst",
        "cons_connect_type", "consumer_identy", "is_old_mtr_available",
        "xi_kva_code", "transaction_id", "issmartmeter", "application_id",
        "application_date", "geo_location", "payment_mode", "metering_mode",
        "last_twelve_month_avg_consm", "last_fy_avg_consm", "last_yr_same_month_consm",
        "transaction_id_for_mdm", "initial_master_sync_id",
    ]
    hints = ["uuid"] + ["uuid[] (ref_level_entries) mandatory"] * 9 + ["uuid[] (ref_level_entries)"] * 6
    hints += ["text", "text mandatory", "text mandatory", "text mandatory", "text mandatory"]
    hints += ["text"] * 7
    hints += ["uuid[] (ref_project_masters_lov) mandatory", "uuid[] (ref_project_masters_lov)",
              "uuid[] (ref_project_masters_lov)", "double precision",
              "uuid[] (ref_project_masters_lov)", "double precision",
              "uuid[] (ref_project_masters_lov)", "text",
              "uuid[] (ref_project_masters_lov)", "uuid[] (ref_project_masters_lov)"]
    hints += ["double precision"] * 4
    hints += ["uuid[] (ref_project_masters_lov)", "text mandatory",
              "uuid[] (ref_project_masters_lov) mandatory",
              "uuid[] (ref_project_masters_lov) mandatory",
              "uuid[] (ref_project_masters_lov) mandatory"]
    hints += ["double precision", "text", "text", "text", "text", "text",
              "uuid[] (ref_project_masters_lov)", "uuid[] (ref_project_masters_lov)",
              "double precision", "double precision", "double precision", "text", "text"]
    ws.append(headers)
    ws.append(hints[:len(headers)])
    ref = wb.create_sheet("ref_project_masters_lov")
    ref.append(["id", "name", "code", "master_name"])
    for row in reference_frame().itertuples(index=False):
        ref.append(list(row))
    wb.save(path)


def operation(src: str, dst: str) -> Operation:
    return Operation(Category[src], Category[dst])
