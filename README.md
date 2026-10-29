# MasterAutomation

**Python-based Consumer Master Data Automation for Excel and CSV Workflows**

A Streamlit application that automates the transformation of electrical distribution consumer records into a prescribed Excel master template using reference UUID lookups, network hierarchy resolution, and configurable business rules.

---

## Overview

MasterAutomation provides a guided, seven-step web workflow for populating consumer-master Excel templates. It reads consumer data from Excel or CSV files, resolves field values to UUIDs from a built-in reference sheet (`ref_project_masters_lov`), matches records against a network hierarchy report, applies operation-specific business rules, validates the output, and writes a template-compliant Excel workbook.

The application supports nine source-to-destination category operations across **WC** (Whole Current), **LTCT** (Low Tension Current Transformer), and **HTCT** (High Tension Current Transformer) meter categories.

## Key Features

- **Guided 7-step workflow** — Upload template → Select operation → Upload hierarchy → Upload source data → Review mappings → Validate → Export.
- **Nine supported operations** — All category conversion, same-category update, and cross-category combinations.
- **Excel and CSV input** — Consumer source and hierarchy files accept `.csv`, `.xlsx`, and `.xls`. The target template requires Excel.
- **Reference UUID resolution** — Business values (meter make, consumer status, phase type, etc.) are resolved to UUIDs from the template's `ref_project_masters_lov` worksheet.
- **Category-specific meter-make mapping** — Meter-make lookups use the destination category's reference table (WC METER MAKE, LTCT METER MAKE, or HTCT METER MAKE).
- **Network hierarchy matching** — Resolves DISCOM → Zone → Circle → Division → Subdivision → Substation → Feeder → DT identifiers with parent-qualified disambiguation.
- **Human-approved hierarchy corrections** — Shows a unique suggested correction for mismatched source identifiers and requires explicit acceptance before the record is considered resolved.
- **Consumer identification business rules** — WC → HTCT/LTCT with ESC auto-resolves to TOD (Rule A); HTCT/LTCT → WC uses source identification (Rule B).
- **UUID formatting** — Every UUID written to the output is normalised to `{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}`. Double braces are collapsed, invalid UUIDs are flagged.
- **Dry-run validation** — Full transformation and validation before generating the output workbook.
- **Audit and error reports** — Downloadable per-record audit log and error report as Excel files.
- **Field-level provenance** — Audit fields include the input value, resolved business value, output value, and resolution method.
- **Large file preflight** — Estimates peak memory use and warns before processing large files.
- **Processing progress and cancellation** — Shows processed-record counts and can stop a long run.
- **Repeated-value acceleration** — Reuses hierarchy and reference lookup results for repeated values.
- **CSV streaming** — Large CSV files are inspected in chunks without storing every identifier during upload checks.

## Supported Operations

| #   | Operation     | Type                  | Description                              |
| --- | ------------- | --------------------- | ---------------------------------------- |
| 1   | WC → HTCT     | Category Conversion   | Convert WC consumers to HTCT             |
| 2   | WC → LTCT     | Category Conversion   | Convert WC consumers to LTCT             |
| 3   | HTCT → WC     | Category Conversion   | Convert HTCT consumers to WC             |
| 4   | LTCT → WC     | Category Conversion   | Convert LTCT consumers to WC             |
| 5   | WC → WC       | Same-Category Update  | Update WC consumers                      |
| 6   | LTCT → LTCT   | Same-Category Update  | Update LTCT consumers                    |
| 7   | HTCT → HTCT   | Same-Category Update  | Update HTCT consumers                    |
| 8   | LTCT → HTCT   | Cross-Category Update | Convert LTCT consumers to HTCT           |
| 9   | HTCT → LTCT   | Cross-Category Update | Convert HTCT consumers to LTCT           |

## Workflow

1. **Upload Template** — Upload the target Excel workbook containing the output column structure and the `ref_project_masters_lov` reference sheet.
2. **Select Operation** — Choose one of the nine source → destination category operations.
3. **Upload Hierarchy** — Upload the GAA / Network Hierarchy Report (CSV or Excel) used to resolve DISCOM through DT identifiers.
4. **Upload Source Data** — Upload the consumer source data file (CSV or Excel) containing the records to be transformed.
5. **Review Mappings** — Confirm automatic source-to-target column mappings, with the option to override individual mappings.
6. **Validate** — Execute the full transformation pipeline and review per-record validation results, audit logs, and error reports.
7. **Export** — Generate the master workbook from eligible rows; failed and review-required rows can be exported separately.

## Technology Stack

| Component        | Technology             |
| ---------------- | ---------------------- |
| Language         | Python 3.11+           |
| Web Framework    | Streamlit              |
| Data Processing  | pandas                 |
| Excel I/O        | openpyxl               |
| Testing          | pytest                 |

## Prerequisites

- **Python 3.11** or later
- **pip** (Python package installer)

## Installation

```bash
# Clone the repository
git clone https://github.com/sarjeetshekhawatgenus/Master-Automation.git
cd Master-Automation

# Create a virtual environment (recommended)
python -m venv .venv

# Activate the virtual environment
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Configuration

### Streamlit settings

Server configuration is in [`.streamlit/config.toml`](.streamlit/config.toml):

```toml
[server]
maxUploadSize = 3072     # Maximum upload size in MiB (3 GiB)
headless = true

[browser]
gatherUsageStats = false
```

Restart Streamlit after changing `config.toml`. The official Windows build is kept at `dist-uan-suffix\\MasterAutomation\\MasterAutomation.exe`.

### Application configuration

All field mappings, reference categories, operations, UUID fields, and processing constants are defined in [`master_update_automation/config.py`](master_update_automation/config.py). Key configurable elements:

| Setting                      | Purpose                                           |
| ---------------------------- | ------------------------------------------------- |
| `SOURCE_TO_TARGET_MAP`       | Source column → target column name mapping         |
| `REFERENCE_CATEGORY_MAP`     | Target field → reference `master_name` category    |
| `METER_MAKE_REFERENCE`       | Destination category → meter-make reference name   |
| `UUID_FIELDS`                | Fields that must contain valid UUIDs               |
| `MANDATORY_TARGET_FIELDS`    | Required output fields                             |
| `VALUE_ALIASES`              | Input value normalisation (e.g., Y/N → Yes/No)    |
| `HIERARCHY_CSV_COLUMN_MAP`   | Expected hierarchy report column names             |

### Environment variables

The application does not require environment variables. See [`.env.example`](.env.example) for optional overrides.

## Running the Application

```bash
streamlit run app.py
```

The application opens at `http://localhost:8501` by default.

## Input Files

| File               | Required Format       | Purpose                                                                                          |
| ------------------ | --------------------- | ------------------------------------------------------------------------------------------------ |
| Target Template    | `.xlsx` or `.xls`     | Defines the output column structure. Must contain a `ref_project_masters_lov` worksheet.          |
| Hierarchy Report   | `.csv`, `.xlsx`, `.xls` | Provides network hierarchy identifiers (Discom through DT) with Name, Code, and ID columns.    |
| Consumer Source    | `.csv`, `.xlsx`, `.xls` | Contains the consumer records to be transformed. CSV files skip worksheet selection.             |

CSV files are read as text (`dtype=str`) to preserve leading zeros, consumer numbers, and UUIDs. Supported encodings: UTF-8 (with or without BOM), Windows-1252, and Latin-1. Supported delimiters: comma, semicolon, tab, and pipe.

## Business Rules

### Consumer Identification

| Scenario                                          | Rule   | Behaviour                                        |
| ------------------------------------------------- | ------ | ------------------------------------------------ |
| WC → HTCT/LTCT, source identification is **ESC**  | Rule A | Resolves to **TOD** UUID                         |
| WC → HTCT/LTCT, source identification is not ESC  | —      | Standard lookup                                  |
| HTCT/LTCT → WC                                    | Rule B | Uses source identification value directly        |
| Same category (WC→WC, LTCT→LTCT, HTCT→HTCT)      | Rule C | Standard lookup; no ESC→TOD override             |
| LTCT ↔ HTCT                                       | Rule D | Standard lookup                                  |

### Meter Category

The `meter_category` field is always resolved to **Main Meter** from the `METER CATEGORY` category in `ref_project_masters_lov`. Its UUID is never hardcoded and is never copied from the source file.

### Meter Make

Meter-make lookups use the **destination** category's reference:

- WC destination → `WC METER MAKE`
- LTCT destination → `LTCT METER MAKE`
- HTCT destination → `HTCT METER MAKE`

### UUID Formatting

All UUIDs in the output are formatted as `{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}`:

- Existing braces are stripped to prevent double-wrapping.
- UUID structure is validated.
- Invalid UUID values are reported, not silently dropped.

## Output

The application generates three standard downloads and an optional review-only workbook:

1. **Master Workbook** — A copy of the uploaded template with transformed consumer data populated.
2. **Audit Report** — Per-record status including transformations applied, hierarchy matches, and lookup results.
3. **Error Report** — Filtered view of failed, warning, and review-required records.
4. **Review-only Workbook (optional)** — Failed and review-required rows in a separate workbook, excluded from the eligible master workbook.

## Audit Statuses

| Status            | Meaning                                                  |
| ----------------- | -------------------------------------------------------- |
| `INSERTED`        | New record added                                         |
| `UPDATED`         | Existing record updated                                  |
| `WARNING`         | Non-critical issues (optional lookup failures)           |
| `FAILED`          | Critical errors (missing required references)            |
| `REQUIRES_REVIEW` | Ambiguous matches or suggested corrections not yet accepted |

## Running Tests

```bash
python -m pytest tests/ -v
```

The test suite (154 tests) validates UUID normalisation, all nine consumer-identification rules, meter-make routing, reference lookups, hierarchy matching, column mapping, the full transformation pipeline, file loading, edge cases, and UI support logic. All test data is synthetic.

## Project Structure

```
MasterAutomation/
├── app.py                              # Streamlit entrypoint
├── requirements.txt                    # Runtime and test dependencies
├── pytest.ini                          # Test configuration
├── .gitignore                          # Git exclusion rules
├── .env.example                        # Environment variable template
├── .streamlit/
│   └── config.toml                     # Streamlit server and theme settings
│
├── master_update_automation/           # Application package
│   ├── __init__.py                     # Package metadata (version)
│   ├── app.py                          # Streamlit UI (7-step guided wizard)
│   ├── config.py                       # Operations, mappings, field types, constants
│   ├── pipeline.py                     # Transformation pipeline orchestrator
│   ├── file_loader.py                  # Excel/CSV loading, validation, temp storage
│   ├── schema_detector.py             # Template schema and field-type detection
│   ├── column_mapper.py               # Automatic source→target column mapping
│   ├── reference_lookup.py            # UUID resolution from ref_project_masters_lov
│   ├── hierarchy_mapper.py            # Network hierarchy matching and disambiguation
│   ├── consumer_rules.py              # Business rules (ESC→TOD, meter make)
│   ├── record_matcher.py              # Insert/update/duplicate record detection
│   ├── validator.py                    # Data validation checks
│   ├── excel_writer.py                # Output workbook generation
│   ├── uuid_utils.py                  # UUID normalisation and formatting
│   ├── audit_logger.py                # Audit and error report generation
│   └── ui_support.py                  # UI display helpers and status logic
│
└── tests/                              # Test suite (154 tests)
    ├── conftest.py                     # Shared fixtures and synthetic data
    ├── test_uuid_utils.py             # UUID formatting and normalisation
    ├── test_consumer_rules.py         # All 9 operation rules + meter make
    ├── test_reference_lookup.py       # Reference UUID resolution
    ├── test_hierarchy_mapper.py       # Hierarchy matching and disambiguation
    ├── test_column_mapper.py          # Column mapping logic
    ├── test_file_loader.py            # File loading and CSV handling
    ├── test_pipeline_integration.py   # End-to-end pipeline tests
    ├── test_edge_cases.py             # Edge cases and validators
    ├── test_app_smoke.py              # Application entry-point smoke tests
    └── test_ui_support.py             # UI helper tests
```

## Troubleshooting

| Issue                                    | Resolution                                                                                   |
| ---------------------------------------- | -------------------------------------------------------------------------------------------- |
| Missing `ref_project_masters_lov` sheet  | Upload a template that includes the reference worksheet.                                     |
| Invalid reference value                  | Check that the value exists in the matching `master_name` category in the reference sheet.    |
| Unmatched hierarchy record               | Verify the hierarchy report contains the required DISCOM/Zone/Circle/Division/Subdivision codes. |
| Upload fails silently                    | Restart Streamlit after changing `maxUploadSize` in `.streamlit/config.toml`.                 |
| CSV encoding error                       | Save the file as UTF-8 CSV (with or without BOM).                                            |
| Missing column mapping                   | Use the override panel on the Mappings step to manually assign a source column.              |
| Out-of-memory error                      | Reduce file size, close other applications, or use a machine with more RAM.                  |
| Insufficient disk space                  | Free temporary storage space and retry.                                                      |

## Security and Data Handling

- **Do not commit** real consumer data, hierarchy reports, or production datasets to the repository.
- The `.gitignore` excludes the `input/` and `output/` directories by default.
- The application processes data locally in-memory and through temporary files. No data is sent to external services.
- Review all uploaded files before sharing the repository to ensure no sensitive data is included.

## Limitations

- The configured upload limit is **3072 MiB (3 GiB)** per file. The UI estimates memory use before processing, but actual use depends on file shape and available system resources. Large Excel workbooks can expand substantially when loaded.
- There is no Docker, reverse-proxy, or cloud deployment configuration in this repository. The application is designed to run locally with Streamlit.
- The target template must be an Excel workbook. CSV cannot replace the template because it requires multiple worksheets including `ref_project_masters_lov`.
- Hierarchy disambiguation depends on parent-level codes. If all parent codes are identical across duplicate entries, the match is flagged as ambiguous for manual review.

## License

Internal use only.
