"""
Configuration module for Master Excel Update Automation.

Contains all configurable mappings, field definitions, reference categories,
and processing settings. Modify this module to adjust behavior without
rewriting the processing engine.
"""

from dataclasses import dataclass
from enum import Enum
from typing import Dict, List, Set, Tuple


class Category(Enum):
    """Consumer meter categories."""
    WC = "WC"
    LTCT = "LTCT"
    HTCT = "HTCT"


@dataclass(frozen=True)
class Operation:
    """Represents a source → destination category conversion."""
    source: Category
    destination: Category

    @property
    def label(self) -> str:
        return f"{self.source.value} → {self.destination.value}"

    @property
    def is_conversion(self) -> bool:
        return self.source != self.destination

    @property
    def is_same_category(self) -> bool:
        return self.source == self.destination


OPERATIONS: Dict[int, Operation] = {
    1: Operation(Category.WC, Category.HTCT),
    2: Operation(Category.WC, Category.LTCT),
    3: Operation(Category.HTCT, Category.WC),
    4: Operation(Category.LTCT, Category.WC),
    5: Operation(Category.WC, Category.WC),
    6: Operation(Category.LTCT, Category.LTCT),
    7: Operation(Category.HTCT, Category.HTCT),
    8: Operation(Category.LTCT, Category.HTCT),
    9: Operation(Category.HTCT, Category.LTCT),
}

OPERATION_GROUPS = {
    "Category Conversion Operations": [1, 2, 3, 4],
    "Same-Category and Cross-Category Update Operations": [5, 6, 7, 8, 9],
}

REFERENCE_SHEET_NAME = "ref_project_masters_lov"
REFERENCE_COLUMNS = {
    "id": "id",
    "name": "name",
    "code": "code",
    "master_name": "master_name",
}

# Always resolve meter_category from this name in METER CATEGORY (never from source).
FIXED_METER_CATEGORY_LOOKUP_NAME = "Main Meter"

# Consumer UAN suffix is a fixed target value, independent of source data.
CONSUMER_UAN_SUFFIX_VALUE = "-1"

# Target column name → reference master_name value
REFERENCE_CATEGORY_MAP: Dict[str, str] = {
    "meter_category": "METER CATEGORY",
    "connection_category": "CONNECTION CATEGORY",
    "consumer_status": "CONSUMER STATUS",
    "sanctioned_load_unit": "SANCTIONED LOAD UNIT",
    "meter_phase": "PHASE TYPE",
    "meter_status": "METER STATUS",
    "geographical_are": "GEOGRAPHICAL AREA",
    "cons_connect_type": "CONSUMER CONNECTION TYPE",
    "consumer_identy": "CONSUMER IDENTIFICATION",
    "is_old_mtr_available": "IS OLD METER AVAILABLE",
    "metering_mode": "METERING MODE",
    "payment_mode": "PAYMENT MODE",
}

# Alternate reference category spellings observed in templates
REFERENCE_CATEGORY_ALIASES: Dict[str, List[str]] = {
    "IS OLD METER AVAILABLE": ["IS OLD METER AVAILABE"],
}

METER_MAKE_REFERENCE: Dict[Category, str] = {
    Category.WC: "WC METER MAKE",
    Category.LTCT: "LTCT METER MAKE",
    Category.HTCT: "HTCT METER MAKE",
}

UUID_FIELDS: Set[str] = {
    "id",
    "discom_name",
    "zone_name", "zone_code",
    "circle_name", "circle_code",
    "division_name", "division_code",
    "subdivision_name", "subdivision_code",
    "substation_name", "substation_code",
    "feeder_name", "feeder_code",
    "dt_name", "dt_code",
    "meter_category",
    "connection_category",
    "consumer_status",
    "sanctioned_load_unit",
    "meter_phase",
    "meter_make",
    "meter_status",
    "geographical_are",
    "cons_connect_type",
    "consumer_identy",
    "is_old_mtr_available",
    "payment_mode",
    "metering_mode",
}

HIERARCHY_LEVELS: List[str] = [
    "discom", "zone", "circle", "division",
    "subdivision", "substation", "feeder", "dt",
]

HIERARCHY_CSV_COLUMN_MAP: Dict[str, Dict[str, str]] = {
    "discom": {"name": "Discom Name", "code": "Discom Code", "id": "Discom ID"},
    "zone": {"name": "Zone Name", "code": "Zone Code", "id": "Zone ID"},
    "circle": {"name": "Circle Name", "code": "Circle Code", "id": "Circle ID"},
    "division": {"name": "Division Name", "code": "Division Code", "id": "Division ID"},
    "subdivision": {"name": "Subdivision Name", "code": "Subdivision Code", "id": "Subdivision ID"},
    "substation": {"name": "Substation Name", "code": "Substation Code", "id": "Substation ID"},
    "feeder": {"name": "Feeder Name", "code": "Feeder Code", "id": "Feeder ID"},
    "dt": {"name": "DT Name", "code": "DT Code", "id": "DT ID"},
}

TARGET_HIERARCHY_MAP: Dict[str, Tuple[str, str]] = {
    "discom_name": ("discom", "id"),
    "zone_name": ("zone", "id"),
    "zone_code": ("zone", "id"),
    "circle_name": ("circle", "id"),
    "circle_code": ("circle", "id"),
    "division_name": ("division", "id"),
    "division_code": ("division", "id"),
    "subdivision_name": ("subdivision", "id"),
    "subdivision_code": ("subdivision", "id"),
    "substation_name": ("substation", "id"),
    "substation_code": ("substation", "id"),
    "feeder_name": ("feeder", "id"),
    "feeder_code": ("feeder", "id"),
    "dt_name": ("dt", "id"),
    "dt_code": ("dt", "id"),
}

SOURCE_TO_TARGET_MAP: Dict[str, str] = {
    "response id": "id",
    "consumer number": "consumer_number",
    "old consumer number": "old_consumer_number",
    "account number": "account_number",
    "consumer uan": "consumer_uan_suffix",
    "consumer uan suffix": "consumer_uan_suffix",
    "consumer name": "consumer_name",
    "father name": "father_name",
    "consumer address": "consumer_address",
    "mobile number": "mobile_number",
    "email": "email",
    "discom name": "_hierarchy_discom_name",
    "discom code": "_hierarchy_discom_code",
    "zone name": "_hierarchy_zone_name",
    "zone code": "_hierarchy_zone_code",
    "circle name": "_hierarchy_circle_name",
    "circle code": "_hierarchy_circle_code",
    "division name": "_hierarchy_division_name",
    "division code": "_hierarchy_division_code",
    "subdivision name": "_hierarchy_subdivision_name",
    "subdivision code": "_hierarchy_subdivision_code",
    "substation name": "_hierarchy_substation_name",
    "substation code": "_hierarchy_substation_code",
    "feeder name": "_hierarchy_feeder_name",
    "feeder code": "_hierarchy_feeder_code",
    "dt name": "_hierarchy_dt_name",
    "dt code": "_hierarchy_dt_code",
    "tariff code": "tariff_code",
    "meter category": "meter_category",
    "connection category": "connection_category",
    "consumer status": "consumer_status",
    "connected load": "connected_load",
    "sanctioned load unit": "sanctioned_load_unit",
    "sanctioned load": "sanctioned_load",
    "meter phase": "meter_phase",
    "meter serial number": "meter_serial_no",
    "meter serial no": "meter_serial_no",
    "meter make": "meter_make",
    "meter status": "meter_status",
    "metering voltage": "metering_voltage",
    "supply voltage": "supply_voltage",
    "contract demand": "contract_demand",
    "mf": "mf",
    "geographical area type": "geographical_are",
    "geographical area": "geographical_are",
    "is consumer available in master?": "cons_available_mst",
    "is consumer available in master": "cons_available_mst",
    "consumer connection type": "cons_connect_type",
    "consumer identification": "consumer_identy",
    "is old meter available?": "is_old_mtr_available",
    "is old meter available": "is_old_mtr_available",
    "xi kva code": "xi_kva_code",
    "transaction id": "transaction_id",
    "issmartmeter": "issmartmeter",
    "is smart meter": "issmartmeter",
    "application id": "application_id",
    "application date": "application_date",
    "geo location": "geo_location",
    "payment mode": "payment_mode",
    "metering mode": "metering_mode",
    "last 12 month average consumption": "last_twelve_month_avg_consm",
    "last financial year average consumption": "last_fy_avg_consm",
    "last year same month consumption": "last_yr_same_month_consm",
    "transaction id for mdm": "transaction_id_for_mdm",
    "initial master sync id": "initial_master_sync_id",
}


class FieldType(Enum):
    """How a target field gets its value."""
    DIRECT_COPY = "direct_copy"
    REFERENCE_LOOKUP = "reference_lookup"
    HIERARCHY_LOOKUP = "hierarchy_lookup"
    BUSINESS_RULE = "business_rule"
    NUMERIC = "numeric"
    TEXT = "text"
    BLANK = "blank"
    DEFAULT = "default"


TARGET_FIELD_TYPES: Dict[str, FieldType] = {
    "id": FieldType.DIRECT_COPY,
    "discom_name": FieldType.HIERARCHY_LOOKUP,
    "zone_name": FieldType.HIERARCHY_LOOKUP,
    "zone_code": FieldType.HIERARCHY_LOOKUP,
    "circle_name": FieldType.HIERARCHY_LOOKUP,
    "circle_code": FieldType.HIERARCHY_LOOKUP,
    "division_name": FieldType.HIERARCHY_LOOKUP,
    "division_code": FieldType.HIERARCHY_LOOKUP,
    "subdivision_name": FieldType.HIERARCHY_LOOKUP,
    "subdivision_code": FieldType.HIERARCHY_LOOKUP,
    "substation_name": FieldType.HIERARCHY_LOOKUP,
    "substation_code": FieldType.HIERARCHY_LOOKUP,
    "feeder_name": FieldType.HIERARCHY_LOOKUP,
    "feeder_code": FieldType.HIERARCHY_LOOKUP,
    "dt_name": FieldType.HIERARCHY_LOOKUP,
    "dt_code": FieldType.HIERARCHY_LOOKUP,
    "old_consumer_number": FieldType.TEXT,
    "consumer_number": FieldType.TEXT,
    "account_number": FieldType.TEXT,
    "consumer_uan_suffix": FieldType.BUSINESS_RULE,
    "consumer_name": FieldType.TEXT,
    "father_name": FieldType.TEXT,
    "consumer_address": FieldType.TEXT,
    "mobile_number": FieldType.TEXT,
    "email": FieldType.TEXT,
    "tariff_code": FieldType.TEXT,
    "meter_category": FieldType.REFERENCE_LOOKUP,
    "connection_category": FieldType.REFERENCE_LOOKUP,
    "consumer_status": FieldType.REFERENCE_LOOKUP,
    "connected_load": FieldType.NUMERIC,
    "sanctioned_load_unit": FieldType.REFERENCE_LOOKUP,
    "sanctioned_load": FieldType.NUMERIC,
    "meter_phase": FieldType.REFERENCE_LOOKUP,
    "meter_serial_no": FieldType.TEXT,
    "meter_make": FieldType.BUSINESS_RULE,
    "meter_status": FieldType.REFERENCE_LOOKUP,
    "metering_voltage": FieldType.NUMERIC,
    "supply_voltage": FieldType.NUMERIC,
    "contract_demand": FieldType.NUMERIC,
    "mf": FieldType.NUMERIC,
    "geographical_are": FieldType.REFERENCE_LOOKUP,
    "cons_available_mst": FieldType.TEXT,
    "cons_connect_type": FieldType.REFERENCE_LOOKUP,
    "consumer_identy": FieldType.BUSINESS_RULE,
    "is_old_mtr_available": FieldType.REFERENCE_LOOKUP,
    "xi_kva_code": FieldType.NUMERIC,
    "transaction_id": FieldType.TEXT,
    "issmartmeter": FieldType.TEXT,
    "application_id": FieldType.TEXT,
    "application_date": FieldType.TEXT,
    "geo_location": FieldType.TEXT,
    "payment_mode": FieldType.REFERENCE_LOOKUP,
    "metering_mode": FieldType.REFERENCE_LOOKUP,
    "last_twelve_month_avg_consm": FieldType.NUMERIC,
    "last_fy_avg_consm": FieldType.NUMERIC,
    "last_yr_same_month_consm": FieldType.NUMERIC,
    "transaction_id_for_mdm": FieldType.TEXT,
    "initial_master_sync_id": FieldType.TEXT,
}

MANDATORY_TARGET_FIELDS: Set[str] = {
    "consumer_number",
    "account_number",
    "consumer_uan_suffix",
    "consumer_name",
    "meter_category",
    "cons_available_mst",
    "cons_connect_type",
    "consumer_identy",
    "is_old_mtr_available",
}

MANDATORY_HIERARCHY_FIELDS: Set[str] = {
    "discom_name",
    "zone_name", "zone_code",
    "circle_name", "circle_code",
    "division_name", "division_code",
    "subdivision_name", "subdivision_code",
}

MANDATORY_HIERARCHY_LEVELS: Set[str] = {
    "discom", "zone", "circle", "division", "subdivision",
}

OPTIONAL_LOOKUP_FIELDS: Set[str] = {
    "connection_category",
    "consumer_status",
    "sanctioned_load_unit",
    "meter_phase",
    "meter_make",
    "meter_status",
    "geographical_are",
    "payment_mode",
    "metering_mode",
}

REQUIRED_SOURCE_COLUMNS: Set[str] = {
    "consumer number",
    "account number",
    "consumer name",
    "consumer identification",
}

BUSINESS_KEY_TARGET = "consumer_number"
BUSINESS_KEY_SOURCE_CANDIDATES = [
    "consumer number",
    "consumer_number",
    "account number",
    "account_number",
]

TEXT_IDENTIFIER_FIELDS: Set[str] = {
    "consumer_number", "account_number", "mobile_number",
    "meter_serial_no", "old_consumer_number", "consumer_uan_suffix",
    "tariff_code", "transaction_id", "application_id",
}

VALUE_ALIASES: Dict[str, Dict[str, str]] = {
    "is_old_mtr_available": {
        "Y": "Yes", "N": "No", "TRUE": "Yes", "FALSE": "No",
        "1": "Yes", "0": "No",
    },
    "cons_available_mst": {
        "Y": "Yes", "N": "No", "TRUE": "Yes", "FALSE": "No",
        "1": "Yes", "0": "No",
    },
}

TYPE_HINT_ROW = 2
DATA_START_ROW = 3

CATEGORY_SOURCE_COLUMNS = ["meter category", "meter_category"]

# Per-file upload limit. Streamlit maxUploadSize is in MiB (1024 * 1024 bytes),
# so 3072 MiB == 3 GiB. The UI label uses "3 GB" to match this configured cap.
MAX_UPLOAD_SIZE_MIB = 3072
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MIB * 1024 * 1024
MAX_UPLOAD_SIZE_LABEL = "3 GB"

TEMPLATE_EXTENSIONS = (".xlsx", ".xls")
HIERARCHY_EXTENSIONS = (".csv", ".xlsx", ".xls")
SOURCE_EXTENSIONS = (".csv", ".xlsx", ".xls")
CSV_EXTENSIONS = (".csv",)
EXCEL_EXTENSIONS = (".xlsx", ".xls")

CSV_CHUNK_SIZE = 100_000
CSV_SAMPLE_BYTES = 65_536
UPLOAD_COPY_CHUNK_BYTES = 1024 * 1024
TABLE_PREVIEW_ROWS = 50


class AuditStatus(Enum):
    SUCCESS = "SUCCESS"
    UPDATED = "UPDATED"
    INSERTED = "INSERTED"
    WARNING = "WARNING"
    FAILED = "FAILED"
    REQUIRES_REVIEW = "REQUIRES_REVIEW"
