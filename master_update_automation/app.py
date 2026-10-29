"""Guided Streamlit workflow for master Excel update automation."""

from __future__ import annotations

import html
import io
import re
import threading
import tempfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from time import monotonic

import pandas as pd
import streamlit as st

from .column_mapper import create_column_mapping
from . import __version__
from .config import (
    HIERARCHY_EXTENSIONS,
    OPERATIONS,
    REFERENCE_SHEET_NAME,
    SOURCE_EXTENSIONS,
    TEXT_IDENTIFIER_FIELDS,
    TEMPLATE_EXTENSIONS,
)
from .excel_writer import write_output_workbook
from .file_loader import (
    detect_source_sheet,
    detect_target_sheet,
    get_sheet_headers,
    get_sheet_names,
    inspect_csv_file,
    inspect_excel_sheet,
    load_excel_workbook,
    read_hierarchy_file,
    remove_temp_file,
    save_upload_stream,
    validate_extension,
)
from .hierarchy_mapper import HierarchyMapper
from .pipeline import find_source_key_column, run_pipeline
from .schema_detector import detect_template_schema
from .ui_support import (
    ENGINE_MANDATORY_SOURCE_TARGETS,
    ENGINE_POPULATED_TARGETS,
    MAX_TABLE_PREVIEW_ROWS,
    UI_OPERATION_GROUPS,
    WORKFLOW_STEPS,
    action_counts,
    available_memory_bytes,
    compute_step_statuses,
    estimate_peak_memory_bytes,
    file_fingerprint,
    format_file_size,
    humanize_exception,
    humanize_validation_issue,
    invalidate_processing_results,
    operation_explanation,
    output_workbook_name,
    report_filename,
    results_status_label,
    stepper_html,
    upload_limit_caption,
)


def _inject_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200');

        :root {
            --mda-primary: #0F5C56;
            --mda-primary-hover: #0B3B38;
            --mda-primary-dark: #00433E;
            --mda-accent: #10B981;
            --mda-bg: #F8F9FA;
            --mda-surface: #FFFFFF;
            --mda-surface-alt: #EFF4FF;
            --mda-surface-card: #FFFFFF;
            --mda-border: #E2E8F0;
            --mda-border-strong: #CBD5E1;
            --mda-text: #0D1C2E;
            --mda-text-secondary: #475569;
            --mda-text-muted: #64748B;
            --mda-success: #059669;
            --mda-success-bg: #ECFDF5;
            --mda-success-border: #A7F3D0;
            --mda-warning: #D97706;
            --mda-warning-bg: #FFFBEB;
            --mda-warning-border: #FDE68A;
            --mda-danger: #DC2626;
            --mda-danger-bg: #FEF2F2;
            --mda-danger-border: #FECACA;
            --mda-info: #0284C7;
            --mda-info-bg: #F0F9FF;
            --mda-info-border: #BAE6FD;
        }

        html, body, .stApp {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            color: var(--mda-text);
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        /* Material Symbols ligatures: ensure glyph renders and never spills text into button labels */
        [data-testid="stIconMaterial"], .material-symbols-rounded, .material-symbols-outlined {
            font-family: 'Material Symbols Rounded' !important;
            font-weight: normal !important;
            font-style: normal !important;
            font-size: 1.25rem !important;
            line-height: 1 !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            direction: ltr !important;
            -webkit-font-feature-settings: 'liga' 1 !important;
            font-feature-settings: 'liga' 1 !important;
            -webkit-font-smoothing: antialiased !important;
            overflow: hidden !important;
            max-width: 1.5rem !important;
            margin-right: 0.35rem !important;
        }

        .stApp {
            background:
                radial-gradient(ellipse at 50% 0%, rgba(15, 92, 86, 0.04) 0%, transparent 450px),
                linear-gradient(180deg, #FFFFFF 0%, var(--mda-bg) 260px),
                var(--mda-bg);
        }

        .block-container {
            padding-top: 0.75rem;
            padding-bottom: 3rem;
            max-width: 1240px;
            overflow-x: hidden;
        }

        h1, h2, h3, h4 {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            color: var(--mda-text);
            font-weight: 700;
            letter-spacing: -0.02em;
        }

        #MainMenu, footer, header[data-testid="stHeader"] {
            display: none !important;
        }

        /* Unified Top App Header */
        .mda-topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            padding: 0.75rem 1.15rem;
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 8px;
            margin-bottom: 0.75rem;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        }

        .mda-topbar-brand {
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }

        .mda-mark {
            width: 36px;
            height: 36px;
            border-radius: 8px;
            background: var(--mda-primary);
            color: #FFFFFF;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-weight: 700;
            font-size: 0.88rem;
            letter-spacing: 0.05em;
            box-shadow: 0 1px 2px rgba(15, 92, 86, 0.3);
            flex-shrink: 0;
        }

        .mda-title-wrap {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .mda-title-row {
            display: flex;
            align-items: center;
            gap: 0.55rem;
        }

        .mda-title {
            font-size: 1.12rem !important;
            font-weight: 700 !important;
            color: var(--mda-text) !important;
            letter-spacing: -0.015em;
            margin: 0 !important;
            line-height: 1.2 !important;
        }

        .mda-version-pill {
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.68rem;
            font-weight: 600;
            color: var(--mda-primary);
            background: rgba(15, 92, 86, 0.08);
            border: 1px solid rgba(15, 92, 86, 0.2);
            padding: 0.12rem 0.45rem;
            border-radius: 4px;
        }

        .mda-kicker {
            font-size: 0.78rem;
            font-weight: 500;
            color: var(--mda-text-secondary);
            margin: 0;
            line-height: 1.3;
        }

        .mda-topbar-meta {
            display: flex;
            align-items: center;
            gap: 0.6rem;
            flex-shrink: 0;
        }

        .mda-runtime-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            padding: 0.25rem 0.65rem;
            border-radius: 999px;
            background: var(--mda-surface-alt);
            border: 1px solid var(--mda-border-strong);
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--mda-text);
        }

        .mda-pulsing-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: var(--mda-accent);
            box-shadow: 0 0 0 2px rgba(16, 185, 129, 0.25);
        }

        .mda-docs-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.22rem 0.6rem;
            border-radius: 6px;
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            font-size: 0.75rem;
            font-weight: 500;
            color: var(--mda-text-secondary);
        }

        .mda-profile-pill {
            display: inline-flex;
            align-items: center;
            padding: 0.25rem 0.6rem;
            border-radius: 6px;
            background: #E7F3F1;
            border: 1px solid #B8DFDB;
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--mda-primary);
        }

        /* Stepper */
        .mda-stepper-container {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 8px;
            padding: 0.65rem 0.85rem;
            margin-bottom: 0.85rem;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
        }

        .mda-stepper {
            list-style: none;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.35rem;
            padding: 0;
            margin: 0;
            overflow-x: auto;
        }

        .mda-step {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.4rem 0.65rem;
            border-radius: 6px;
            border: 1px solid var(--mda-border);
            background: #FFFFFF;
            color: var(--mda-text-secondary);
            font-size: 0.8rem;
            flex-shrink: 0;
            transition: all 0.15s ease;
        }

        .mda-connector {
            flex: 1 1 auto;
            height: 2px;
            min-width: 8px;
            max-width: 32px;
            background: #E2E8F0;
            border-radius: 1px;
            list-style: none;
        }

        .mda-connector.mda-conn-active {
            background: var(--mda-success);
        }

        .mda-connector.mda-conn-pending {
            background: #E2E8F0;
        }

        .mda-step-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            height: 22px;
            border-radius: 50%;
            background: var(--mda-surface-alt);
            color: var(--mda-text-muted);
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.7rem;
            font-weight: 700;
        }

        .mda-step-info {
            display: flex;
            flex-direction: column;
            gap: 1px;
        }

        .mda-step-label {
            font-weight: 600;
            font-size: 0.8rem;
            color: var(--mda-text);
            line-height: 1.2;
        }

        .mda-step-tag {
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.6rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            padding: 1px 4px;
            border-radius: 3px;
            width: fit-content;
            line-height: 1.2;
        }

        .mda-tag-active {
            background: rgba(15, 92, 86, 0.1);
            color: var(--mda-primary);
        }

        .mda-tag-done {
            background: var(--mda-success-bg);
            color: var(--mda-success);
        }

        .mda-tag-attn {
            background: var(--mda-warning-bg);
            color: var(--mda-warning);
        }

        .mda-step-current {
            border-color: var(--mda-primary) !important;
            background: #F0FDF4 !important;
            box-shadow: 0 0 0 1px var(--mda-primary);
        }

        .mda-step-current .mda-step-badge {
            background: var(--mda-primary) !important;
            color: #FFFFFF !important;
        }

        .mda-step-current .mda-step-label {
            color: var(--mda-primary) !important;
            font-weight: 700 !important;
        }

        .mda-step-complete {
            border-color: var(--mda-success-border) !important;
            background: #F0FDF4 !important;
        }

        .mda-step-complete .mda-step-badge {
            background: var(--mda-success) !important;
            color: #FFFFFF !important;
        }

        .mda-step-complete .mda-step-label {
            color: #047857;
        }

        .mda-step-attention {
            border-color: var(--mda-warning-border) !important;
            background: var(--mda-warning-bg) !important;
        }

        .mda-step-attention .mda-step-badge {
            background: var(--mda-warning) !important;
            color: #FFFFFF !important;
        }

        .mda-step-attention .mda-step-label {
            color: #B45309;
        }

        /* Instruction Banner */
        .mda-banner {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 8px;
            padding: 1rem 1.35rem;
            margin-bottom: 0.85rem;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        }

        .mda-banner-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-bottom: 0.45rem;
        }

        .mda-stepchip {
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.7rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: var(--mda-primary);
            background: var(--mda-success-bg);
            border: 1px solid var(--mda-success-border);
            padding: 0.2rem 0.55rem;
            border-radius: 4px;
            display: inline-block;
        }

        .mda-telemetry {
            display: flex;
            align-items: center;
            gap: 0.45rem;
        }

        .mda-telemetry-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.18rem 0.5rem;
            background: var(--mda-surface-alt);
            border: 1px solid var(--mda-border);
            border-radius: 4px;
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.68rem;
            color: var(--mda-text-secondary);
        }

        .mda-dot {
            width: 6px;
            height: 6px;
            border-radius: 50%;
            background: var(--mda-accent);
        }

        .mda-banner-title {
            margin: 0.2rem 0 0.35rem 0 !important;
            font-size: 1.25rem !important;
            color: var(--mda-text) !important;
        }

        .mda-lede {
            color: var(--mda-text-secondary);
            font-size: 0.88rem;
            margin: 0;
            line-height: 1.45;
        }

        /* File Card */
        .mda-file {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-left: 4px solid var(--mda-primary);
            border-radius: 6px;
            padding: 0.85rem 1.1rem;
            margin: 0.5rem 0 0.85rem 0;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
        }

        .mda-file strong {
            display: block;
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.92rem;
            color: var(--mda-text);
            word-break: break-word;
        }

        .mda-meta {
            color: var(--mda-text-secondary);
            font-size: 0.8rem;
            margin-top: 0.25rem;
        }

        .mda-status {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            margin-top: 0.45rem;
            padding: 0.2rem 0.55rem;
            border-radius: 4px;
            font-family: 'JetBrains Mono', 'Segoe UI Mono', 'Cascadia Code', Consolas, monospace;
            font-size: 0.72rem;
            font-weight: 600;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            border: 1px solid transparent;
        }

        .mda-ok {
            background: var(--mda-success-bg);
            color: var(--mda-success);
            border-color: var(--mda-success-border);
        }

        .mda-warn {
            background: var(--mda-warning-bg);
            color: var(--mda-warning);
            border-color: var(--mda-warning-border);
        }

        .mda-fail {
            background: var(--mda-danger-bg);
            color: var(--mda-danger);
            border-color: var(--mda-danger-border);
        }

        .mda-idle {
            background: var(--mda-surface-alt);
            color: var(--mda-text-secondary);
            border-color: var(--mda-border-strong);
        }

        .mda-idle-hint {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            margin: 0.5rem 0 0.85rem 0;
            padding: 0.35rem 0.75rem;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 6px;
            font-size: 0.78rem;
            font-weight: 500;
            color: #64748B;
        }

        .mda-idle-dot {
            width: 6px;
            height: 6px;
            border-radius: 50%;
            background: #94A3B8;
        }

        /* Operations Selection Layout */
        .mda-op-section-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 0.4rem;
            margin-bottom: 0.6rem;
            border-bottom: 1px solid var(--mda-border);
        }

        .mda-op-section-title {
            font-size: 0.92rem;
            font-weight: 700;
            color: var(--mda-text);
        }

        .mda-op-badge-count {
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.7rem;
            padding: 0.15rem 0.45rem;
            border-radius: 4px;
            background: var(--mda-surface-alt);
            border: 1px solid var(--mda-border);
            color: var(--mda-text-secondary);
        }

        .mda-op-config-card {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 8px;
            padding: 1rem 1.25rem;
            margin: 0.85rem 0;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        }

        .mda-op-config-title {
            font-size: 0.88rem;
            font-weight: 700;
            color: var(--mda-primary);
            margin-bottom: 0.45rem;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }

        .mda-op-config-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 0.65rem;
            font-size: 0.82rem;
        }

        .mda-op-config-item {
            background: var(--mda-bg);
            border: 1px solid var(--mda-border);
            border-radius: 6px;
            padding: 0.55rem 0.75rem;
        }

        .mda-op-config-label {
            display: block;
            font-size: 0.68rem;
            font-family: 'JetBrains Mono', monospace;
            color: var(--mda-text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 0.2rem;
        }

        .mda-op-config-val {
            font-weight: 600;
            color: var(--mda-text);
        }

        /* KPI Metrics */
        .mda-kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
            gap: 0.55rem;
            margin: 0.85rem 0 1.15rem 0;
        }

        .mda-kpi-card {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 6px;
            padding: 0.75rem 0.85rem;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.02);
        }

        .mda-kpi-label {
            display: block;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.68rem;
            font-weight: 600;
            color: var(--mda-text-muted);
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 0.25rem;
        }

        .mda-kpi-val {
            font-family: 'JetBrains Mono', monospace;
            font-size: 1.35rem;
            font-weight: 700;
            color: var(--mda-text);
            line-height: 1.2;
        }

        /* Buttons */
        .stButton > button {
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.84rem;
            transition: all 0.15s ease;
            white-space: normal;
            border: 1px solid var(--mda-border-strong);
        }

        .stButton > button:hover {
            border-color: var(--mda-primary);
        }

        .stButton > button[kind="primary"] {
            background: var(--mda-primary) !important;
            border-color: var(--mda-primary-dark) !important;
            color: #FFFFFF !important;
            box-shadow: 0 1px 2px rgba(15, 92, 86, 0.25);
        }

        .stButton > button[kind="primary"]:hover {
            background: var(--mda-primary-hover) !important;
        }

        .stDownloadButton > button {
            border-radius: 6px;
            font-weight: 600;
            font-size: 0.84rem;
            white-space: normal;
            min-height: 2.6rem;
            border: 1px solid var(--mda-border-strong);
            background: #FFFFFF;
            transition: all 0.15s ease;
        }

        .stDownloadButton > button:hover {
            border-color: var(--mda-primary);
            background: var(--mda-surface-alt);
        }

        /* File Uploader */
        div[data-testid="stFileUploader"] {
            max-width: 100%;
        }

        div[data-testid="stFileUploader"] section {
            background: #FFFFFF;
            border: 1.5px dashed var(--mda-border-strong);
            border-radius: 8px;
            padding: 1.25rem;
            transition: border-color 0.15s ease;
        }

        div[data-testid="stFileUploader"] section:hover {
            border-color: var(--mda-primary);
        }

        /* Suppress raw ligature text in file uploader dropzone */
        div[data-testid="stFileUploaderDropzone"] [data-testid="stIconMaterial"] {
            display: none !important;
        }

        div[data-testid="stFileUploaderDropzone"] button {
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 0.45rem !important;
            font-weight: 600 !important;
            font-size: 0.85rem !important;
            white-space: nowrap !important;
            position: relative !important;
        }

        div[data-testid="stFileUploaderDropzone"] button::before {
            content: "" !important;
            display: inline-block !important;
            width: 15px !important;
            height: 15px !important;
            background-color: currentColor !important;
            -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cpath d='M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z'/%3E%3C/svg%3E") no-repeat center / contain !important;
            mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cpath d='M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z'/%3E%3C/svg%3E") no-repeat center / contain !important;
        }

        div[data-testid="stFileUploaderDropzoneInstructions"] {
            padding-left: 0.5rem !important;
        }

        /* Dataframe & Tables */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--mda-border);
            border-radius: 6px;
            overflow: hidden;
            background: #FFFFFF;
        }

        /* Export Cards */
        .mda-export-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 0.85rem;
            margin: 0.85rem 0;
        }

        .mda-export-card {
            background: #FFFFFF;
            border: 1px solid var(--mda-border);
            border-radius: 8px;
            padding: 1.15rem;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        }

        .mda-export-card-hero {
            border-color: var(--mda-primary);
            background: linear-gradient(180deg, #FFFFFF 0%, #F0FDF4 100%);
            box-shadow: 0 2px 6px rgba(15, 92, 86, 0.08);
        }

        .mda-export-title {
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--mda-text);
            margin-bottom: 0.25rem;
        }

        .mda-export-desc {
            font-size: 0.8rem;
            color: var(--mda-text-secondary);
            margin-bottom: 0.75rem;
        }

        /* Responsive */
        @media (max-width: 840px) {
            .mda-topbar { flex-direction: column; align-items: flex-start; gap: 0.5rem; }
            .mda-topbar-meta { width: 100%; justify-content: flex-start; flex-wrap: wrap; }
            .mda-hero { flex-direction: column; align-items: flex-start; }
            .mda-banner-top { flex-direction: column; align-items: flex-start; }
            .mda-telemetry { flex-wrap: wrap; }
            .mda-kpi-grid { grid-template-columns: repeat(2, 1fr); }
            .mda-stepper { overflow-x: auto; -webkit-overflow-scrolling: touch; }
        }

        @media (max-width: 500px) {
            .mda-kpi-grid { grid-template-columns: 1fr; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )



def _save_upload(uploaded, suffix: str, previous_path: str | None = None) -> str:
    path = save_upload_stream(uploaded, suffix, uploaded.name, getattr(uploaded, "size", None))
    if previous_path and previous_path != path:
        remove_temp_file(previous_path)
    return path


def _upload_caption(accepted: str) -> None:
    st.caption(f"{upload_limit_caption()}. Accepted formats: {accepted}.")


def _init_state() -> None:
    defaults = {
        "step": 1,
        "template_path": None,
        "template_name": None,
        "template_size": None,
        "template_fp": None,
        "template_error": None,
        "template_error_detail": None,
        "target_sheet": None,
        "template_headers": [],
        "sheet_names": [],
        "ref_sheet_ok": False,
        "ref_category_count": 0,
        "operation_id": None,
        "hierarchy_path": None,
        "hierarchy_name": None,
        "hierarchy_size": None,
        "hierarchy_fp": None,
        "hierarchy_summary": None,
        "hierarchy_error": None,
        "hierarchy_error_detail": None,
        "hierarchy_warnings": [],
        "source_path": None,
        "source_name": None,
        "source_size": None,
        "source_fp": None,
        "source_sheet": None,
        "source_headers": [],
        "source_row_count": None,
        "source_key_column": None,
        "source_error": None,
        "source_error_detail": None,
        "source_sheets": [],
        "source_is_csv": False,
        "source_preview": None,
        "source_parsed_key": None,
        "source_dup_warning": None,
        "mapped_column_count": 0,
        "unmapped_source_count": 0,
        "template_parsed_key": None,
        "custom_map": {},
        "mapping_blockers": [],
        "pipeline_result": None,
        "output_bytes": None,
        "output_name": None,
        "review_bytes": None,
        "review_name": None,
        "audit_bytes": None,
        "audit_name": None,
        "error_bytes": None,
        "error_name": None,
        "mapping_confirmed": False,
        "processing": False,
        "run_stamp": None,
        "pipeline_job": None,
        "pipeline_error": None,
        "pipeline_error_detail": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _clear_results() -> None:
    job = st.session_state.get("pipeline_job")
    if job:
        job["cancel_event"].set()
    invalidate_processing_results(st.session_state)
    st.session_state.run_stamp = None
    st.session_state.pipeline_error = None
    st.session_state.pipeline_error_detail = None


class ProcessingCancelled(Exception):
    """Raised when the user cancels an active validation run."""


def _store_pipeline_result(job_result) -> None:
    result, stamp, audit_bytes, error_bytes = job_result
    op = OPERATIONS[st.session_state.operation_id]
    st.session_state.run_stamp = stamp
    st.session_state.pipeline_result = result
    st.session_state.audit_bytes = audit_bytes
    st.session_state.audit_name = report_filename(op, "audit", stamp)
    st.session_state.error_bytes = error_bytes
    st.session_state.error_name = report_filename(op, "errors", stamp)
    st.session_state.output_bytes = None
    st.session_state.output_name = None
    st.session_state.review_bytes = None
    st.session_state.review_name = None
    st.session_state.pipeline_error = None
    st.session_state.pipeline_error_detail = None


def _start_pipeline_job(approved_rows=None) -> None:
    op = OPERATIONS[st.session_state.operation_id]
    template_path = st.session_state.template_path
    target_sheet = st.session_state.target_sheet
    hierarchy_path = st.session_state.hierarchy_path
    source_path = st.session_state.source_path
    source_sheet = None if st.session_state.source_is_csv else st.session_state.source_sheet
    source_row_count_hint = st.session_state.source_row_count
    custom_map = dict(st.session_state.custom_map or {})
    progress_state = {"message": "Starting validation…", "fraction": 0.0}
    started_at = monotonic()
    progress_lock = threading.Lock()
    cancel_event = threading.Event()
    executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="masterautomation")

    def report_progress(message: str) -> None:
        if cancel_event.is_set():
            raise ProcessingCancelled("Validation cancelled by user.")
        fraction = 0.0
        match = re.search(r"Processing records:\s*([\d,]+)\s*/\s*([\d,]+)", message)
        if match:
            total = int(match.group(2).replace(",", ""))
            done = int(match.group(1).replace(",", ""))
            fraction = done / total if total else 0.0
            elapsed = max(0.1, monotonic() - started_at)
            rate = done / elapsed
            eta = (total - done) / rate if rate else 0.0
            message = f"{message} · {rate:,.0f} rows/s · about {eta:,.0f}s remaining"
        with progress_lock:
            progress_state["message"] = message
            progress_state["fraction"] = fraction

    def run_and_prepare_reports():
        result = run_pipeline(
            template_path=template_path,
            target_sheet=target_sheet,
            operation=op,
            hierarchy_path=hierarchy_path,
            source_path=source_path,
            source_sheet=source_sheet,
            custom_map=custom_map,
            progress=report_progress,
            approved_hierarchy_rows=set(approved_rows or ()),
            source_row_count_hint=source_row_count_hint,
        )
        report_progress("Preparing audit and error reports")
        stamp = datetime.now()
        audit_buffer = io.BytesIO()
        with pd.ExcelWriter(audit_buffer, engine="openpyxl") as writer:
            result.audit.to_dataframe().to_excel(writer, sheet_name="Record Audit", index=False)
            result.audit.to_field_dataframe().to_excel(writer, sheet_name="Field Audit", index=False)
        error_buffer = io.BytesIO()
        with pd.ExcelWriter(error_buffer, engine="openpyxl") as writer:
            result.audit.get_error_report().to_excel(writer, sheet_name="Record Issues", index=False)
            result.audit.get_field_error_report().to_excel(writer, sheet_name="Field Issues", index=False)
        report_progress("Validation and reports complete")
        return result, stamp, audit_buffer.getvalue(), error_buffer.getvalue()

    future = executor.submit(run_and_prepare_reports)
    st.session_state.pipeline_job = {
        "future": future,
        "executor": executor,
        "cancel_event": cancel_event,
        "progress_state": progress_state,
        "progress_lock": progress_lock,
    }
    st.session_state.pipeline_error = None
    st.session_state.pipeline_error_detail = None
    st.session_state.processing = True


@st.fragment(run_every="1s")
def _render_pipeline_job() -> None:
    job = st.session_state.get("pipeline_job")
    if not job:
        return
    with job["progress_lock"]:
        message = job["progress_state"]["message"]
        fraction = job["progress_state"]["fraction"]
    st.progress(min(1.0, max(0.0, fraction)), text=message)
    if not job["future"].done():
        if st.button("Cancel validation", key="cancel_validation"):
            job["cancel_event"].set()
            st.info("Cancellation requested. The app will stop at the next processing checkpoint.")
        return

    try:
        result = job["future"].result()
        _store_pipeline_result(result)
    except ProcessingCancelled:
        st.session_state.pipeline_error = "Validation was cancelled. No new result was generated."
        st.session_state.pipeline_error_detail = ""
    except MemoryError as exc:
        st.session_state.pipeline_error = (
            "The current environment does not have enough memory to process this file. "
            "Reduce the file size or close other applications, then try again."
        )
        st.session_state.pipeline_error_detail = str(exc)
    except Exception as exc:
        user_message, detail = humanize_exception(exc, "source")
        st.session_state.pipeline_error = user_message
        st.session_state.pipeline_error_detail = detail
    finally:
        job["executor"].shutdown(wait=False, cancel_futures=True)
        st.session_state.pipeline_job = None
        st.session_state.processing = False
    st.rerun(scope="app")


def _render_hierarchy_correction_review(result) -> None:
    review_rows = []
    for entry in result.audit.entries:
        if entry.hierarchy_status != "REQUIRES_REVIEW" or not entry.hierarchy_corrections:
            continue
        parts = [
            f"{item.get('level', '').title()} {item.get('kind', '')}: "
            f"{item.get('source_value', '')} → {item.get('suggested_value', '')}"
            for item in entry.hierarchy_corrections
        ]
        review_rows.append({
            "Accept suggested correction": False,
            "Source row": entry.source_row,
            "Consumer number": entry.business_key,
            "Suggested hierarchy correction": "; ".join(parts),
        })
    if not review_rows:
        return
    st.markdown("**Review suggested hierarchy corrections**")
    st.caption(
        "The app found a unique hierarchy path, but it will not count that row as resolved "
        "until you accept the suggested correction. Blank source details stay blank."
    )
    edited = st.data_editor(
        pd.DataFrame(review_rows),
        hide_index=True,
        use_container_width=True,
        disabled=["Source row", "Consumer number", "Suggested hierarchy correction"],
        column_config={
            "Accept suggested correction": st.column_config.CheckboxColumn(
                "Accept suggested correction",
                help="Approve the hierarchy values shown for this consumer.",
            ),
        },
        key=f"hierarchy_review_{st.session_state.run_stamp}",
    )
    selected = edited.loc[edited["Accept suggested correction"] == True, "Source row"].tolist()
    if st.button(
        f"Accept selected corrections and revalidate ({len(selected)})",
        disabled=not selected or st.session_state.processing,
        key="accept_hierarchy_corrections",
    ):
        _start_pipeline_job({int(row) for row in selected})
        st.rerun()


def _issue_expander(title: str, detail: str) -> None:
    with st.expander(title):
        st.code(detail)


def _nav(back: int | None, nxt: int | None, next_disabled: bool = False) -> None:
    busy = bool(st.session_state.processing)
    if back and nxt:
        cols = st.columns([1.5, 1.5, 7])
        if cols[0].button("Back", use_container_width=True, disabled=busy):
            st.session_state.step = back
            st.rerun()
        if cols[1].button("Continue", use_container_width=True, disabled=next_disabled or busy, type="primary"):
            st.session_state.step = nxt
            st.rerun()
    elif nxt:
        cols = st.columns([2, 8])
        if cols[0].button("Continue", use_container_width=True, disabled=next_disabled or busy, type="primary"):
            st.session_state.step = nxt
            st.rerun()
    elif back:
        cols = st.columns([2, 8])
        if cols[0].button("Back", use_container_width=True, disabled=busy):
            st.session_state.step = back
            st.rerun()


def _step_heading(index: int, title: str, lede: str) -> None:
    step_name = WORKFLOW_STEPS[index - 1][1]
    st.markdown(
        f"""
        <div class="mda-banner">
          <div class="mda-banner-top">
            <span class="mda-stepchip">STEP {index:02d} OF 07 · {html.escape(step_name.upper())}</span>
            <div class="mda-telemetry">
              <span class="mda-telemetry-pill">
                <span class="mda-dot"></span>Session: <strong>SES-MASTER-DATA</strong>
              </span>
              <span class="mda-telemetry-pill">
                Runtime: <strong>Python 3.11+</strong>
              </span>
              <span class="mda-telemetry-pill">
                Engine: <strong>v{__version__}</strong>
              </span>
            </div>
          </div>
          <h2 class="mda-banner-title">{html.escape(title)}</h2>
          <p class="mda-lede">{html.escape(lede)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _file_card(name: str, size, kind: str, status: str, status_class: str, extra: str = "") -> None:
    extra_html = html.escape(extra) if extra else ""
    st.markdown(
        f"""
        <div class="mda-file">
          <div style="display: flex; align-items: flex-start; justify-content: space-between; gap: 0.75rem; flex-wrap: wrap;">
            <div>
              <strong>{html.escape(str(name or ""))}</strong>
              <div class="mda-meta">{html.escape(kind)} · {html.escape(format_file_size(size))}{f' · {extra_html}' if extra_html else ''}</div>
            </div>
            <span class="mda-status {html.escape(status_class)}">{html.escape(status)}</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def _display_identifier_value(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return format(value, ".0f")
    return str(value).strip()


def _preview_frame(df: pd.DataFrame, caption: str) -> None:
    if df is None or df.empty:
        st.caption("No rows to display.")
        return
    shown = df.head(MAX_TABLE_PREVIEW_ROWS).copy()
    for column in shown.columns:
        if pd.api.types.is_object_dtype(shown[column].dtype):
            # Streamlit serializes tables through Arrow. Mixed Python values
            # (for example, integer row numbers and text diagnostics) must be
            # normalized for display or Arrow can reject the whole table.
            shown[column] = shown[column].map(_display_identifier_value)
    identifier_columns = {}
    for column in shown.columns:
        normalized = str(column).strip().lower().replace(" ", "_")
        if normalized not in TEXT_IDENTIFIER_FIELDS and normalized not in {
            "consumer_no",
            "account_no",
        }:
            continue
        shown[column] = shown[column].map(_display_identifier_value)
        identifier_columns[column] = st.column_config.TextColumn(
            str(column),
            width="medium",
        )
    st.caption(f"{caption} Showing {len(shown)} of {len(df)} rows.")
    st.dataframe(
        shown,
        use_container_width=True,
        hide_index=True,
        column_config=identifier_columns,
    )


def step_template() -> None:
    _step_heading(
        1,
        "Upload the target template",
        "Upload the Excel workbook that defines the required output structure and reference values.",
    )
    uploaded = st.file_uploader("Target template (.xlsx or .xls)", type=["xlsx", "xls"], key="tpl")
    _upload_caption(".xlsx, .xls")
    if uploaded:
        fp = file_fingerprint(uploaded.name, uploaded.size)
        if st.session_state.template_fp != fp:
            try:
                validate_extension(uploaded.name, TEMPLATE_EXTENSIONS, "target template")
                path = _save_upload(uploaded, Path(uploaded.name).suffix, st.session_state.template_path)
                wb = load_excel_workbook(path)
                st.session_state.template_path = path
                st.session_state.template_name = uploaded.name
                st.session_state.template_size = uploaded.size
                st.session_state.template_fp = fp
                st.session_state.sheet_names = get_sheet_names(wb)
                st.session_state.ref_sheet_ok = REFERENCE_SHEET_NAME in wb.sheetnames
                detected = detect_target_sheet(wb)
                st.session_state.target_sheet = detected
                st.session_state.template_error = None
                st.session_state.template_error_detail = None
                st.session_state.mapping_confirmed = False
                st.session_state.custom_map = {}
                st.session_state.template_parsed_key = None
                _clear_results()
                wb.close()
            except Exception as exc:
                user_msg, detail = humanize_exception(exc, "template")
                remove_temp_file(st.session_state.get("template_path"))
                st.session_state.template_path = None
                st.session_state.template_fp = fp
                st.session_state.ref_sheet_ok = False
                st.session_state.template_error = user_msg
                st.session_state.template_error_detail = detail
                _clear_results()

    if st.session_state.template_error:
        st.error(st.session_state.template_error)
        _issue_expander("Technical details", st.session_state.template_error_detail or "")
        _nav(None, 2, next_disabled=True)
        return

    if not st.session_state.template_path:
        st.markdown('<div class="mda-idle-hint"><span class="mda-idle-dot"></span>Waiting for template file upload...</div>', unsafe_allow_html=True)
        _nav(None, 2, next_disabled=True)
        return

    options = [s for s in st.session_state.sheet_names if s != REFERENCE_SHEET_NAME]
    if not options:
        st.error("The template has no consumer-master worksheet besides the reference sheet. Upload a complete template.")
        _nav(None, 2, next_disabled=True)
        return

    current = st.session_state.target_sheet if st.session_state.target_sheet in options else options[0]
    selected = st.selectbox("Target worksheet", options, index=options.index(current))
    if selected != st.session_state.target_sheet:
        st.session_state.target_sheet = selected
        st.session_state.mapping_confirmed = False
        _clear_results()

    parsed_key = (st.session_state.template_fp, selected)
    try:
        if st.session_state.template_parsed_key != parsed_key:
            wb = load_excel_workbook(st.session_state.template_path)
            st.session_state.template_headers = get_sheet_headers(wb, selected)
            schema = detect_template_schema(wb, selected)
            st.session_state.ref_category_count = len(schema.reference_categories)
            st.session_state.ref_sheet_ok = REFERENCE_SHEET_NAME in wb.sheetnames
            st.session_state.template_parsed_key = parsed_key
            wb.close()
    except Exception as exc:
        user_msg, detail = humanize_exception(exc, "template")
        st.error(user_msg)
        _issue_expander("Technical details", detail)
        _nav(None, 2, next_disabled=True)
        return

    if st.session_state.ref_sheet_ok:
        _file_card(
            st.session_state.template_name,
            st.session_state.template_size,
            "Excel workbook",
            "File validated",
            "mda-ok",
            f"Target sheet: {selected}",
        )
        st.info(
            f"Required reference worksheet `{REFERENCE_SHEET_NAME}` identified. "
            f"{len(st.session_state.template_headers)} target columns detected."
        )
    else:
        _file_card(
            st.session_state.template_name,
            st.session_state.template_size,
            "Excel workbook",
            "File rejected",
            "mda-fail",
        )
        st.error(
            f"The uploaded template is missing `{REFERENCE_SHEET_NAME}`. "
            "This worksheet supplies reference UUIDs, including Main Meter. "
            "Upload a template that includes that sheet."
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("Worksheets", len(st.session_state.sheet_names))
    c2.metric("Target columns", len(st.session_state.template_headers))
    c3.metric("Reference sheet", "Found" if st.session_state.ref_sheet_ok else "Missing")
    st.caption("Sheets: " + ", ".join(st.session_state.sheet_names))
    with st.expander("Target column order"):
        _preview_frame(
            pd.DataFrame({"column": st.session_state.template_headers}),
            "Template headers.",
        )
    _nav(None, 2, next_disabled=not (st.session_state.template_path and st.session_state.target_sheet and st.session_state.ref_sheet_ok))


def step_operation() -> None:
    _step_heading(
        2,
        "Select the processing operation",
        "Choose the source and destination meter categories for this processing run.",
    )
    current = st.session_state.operation_id
    left, right = st.columns(2)
    sections = list(UI_OPERATION_GROUPS.items())
    for col, (group_name, ids) in zip((left, right), sections):
        with col:
            count = len(ids)
            st.markdown(
                f"""
                <div class="mda-op-section-header">
                  <span class="mda-op-section-title">{html.escape(group_name)}</span>
                  <span class="mda-op-badge-count">{count} operations</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            for op_id in ids:
                op = OPERATIONS[op_id]
                selected = current == op_id
                label = f"{'●' if selected else '○'}  {op.label}"
                btn_type = "primary" if selected else "secondary"
                if st.button(label, key=f"op_{op_id}", use_container_width=True, type=btn_type):
                    if st.session_state.operation_id != op_id:
                        _clear_results()
                    st.session_state.operation_id = op_id
                    st.rerun()

    op = OPERATIONS.get(st.session_state.operation_id)
    if op:
        lookup_ref = f"{op.destination.value} METER MAKE" if not op.is_same_category else f"{op.source.value} Reference"
        rule_text = "Rule A: ESC auto-resolves to TOD" if (op.source.value == "WC" and op.destination.value in ("HTCT", "LTCT")) else ("Rule B: Source identification preserved" if op.source.value in ("HTCT", "LTCT") and op.destination.value == "WC" else "Standard Field Reconciliation")
        
        st.markdown(
            f"""
            <div class="mda-op-config-card">
              <div class="mda-op-config-title">Active Configuration: {html.escape(op.label)}</div>
              <div class="mda-op-config-grid">
                <div class="mda-op-config-item">
                  <span class="mda-op-config-label">Source Category</span>
                  <span class="mda-op-config-val">{html.escape(op.source.value)}</span>
                </div>
                <div class="mda-op-config-item">
                  <span class="mda-op-config-label">Destination Category</span>
                  <span class="mda-op-config-val">{html.escape(op.destination.value)}</span>
                </div>
                <div class="mda-op-config-item">
                  <span class="mda-op-config-label">Lookup Reference</span>
                  <span class="mda-op-config-val">{html.escape(lookup_ref)}</span>
                </div>
                <div class="mda-op-config-item">
                  <span class="mda-op-config-label">Business Rule</span>
                  <span class="mda-op-config-val">{html.escape(rule_text)}</span>
                </div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(f"Source {op.source.value} → destination {op.destination.value}. {operation_explanation(op)}")
    else:
        st.markdown('<div class="mda-idle-hint"><span class="mda-idle-dot"></span>No operation selected yet</div>', unsafe_allow_html=True)
    _nav(1, 3, next_disabled=st.session_state.operation_id is None)



def step_hierarchy() -> None:
    _step_heading(
        3,
        "Upload the network hierarchy report",
        "Upload the report used to resolve DISCOM through DT identifiers.",
    )
    uploaded = st.file_uploader("Hierarchy report (CSV or Excel)", type=["csv", "xlsx", "xls"], key="hier")
    _upload_caption(".csv, .xlsx, .xls — UTF-8 CSV recommended")
    if uploaded:
        fp = file_fingerprint(uploaded.name, uploaded.size)
        if st.session_state.hierarchy_fp != fp:
            try:
                validate_extension(uploaded.name, HIERARCHY_EXTENSIONS, "hierarchy report")
                path = _save_upload(uploaded, Path(uploaded.name).suffix, st.session_state.hierarchy_path)
                df = read_hierarchy_file(path)
                mapper = HierarchyMapper()
                summary = mapper.load_hierarchy(df)
                st.session_state.hierarchy_path = path
                st.session_state.hierarchy_name = uploaded.name
                st.session_state.hierarchy_size = uploaded.size
                st.session_state.hierarchy_fp = fp
                st.session_state.hierarchy_summary = summary
                st.session_state.hierarchy_warnings = mapper.validate_hierarchy_columns()
                st.session_state.hierarchy_error = None
                st.session_state.hierarchy_error_detail = None
                _clear_results()
            except Exception as exc:
                user_msg, detail = humanize_exception(exc, "hierarchy")
                remove_temp_file(st.session_state.get("hierarchy_path"))
                st.session_state.hierarchy_path = None
                st.session_state.hierarchy_summary = None
                st.session_state.hierarchy_fp = fp
                st.session_state.hierarchy_error = user_msg
                st.session_state.hierarchy_error_detail = detail
                _clear_results()

    if st.session_state.hierarchy_error:
        st.error(st.session_state.hierarchy_error)
        _issue_expander("Technical details", st.session_state.hierarchy_error_detail or "")
        _nav(2, 4, next_disabled=True)
        return

    if not st.session_state.hierarchy_path:
        st.markdown('<div class="mda-idle-hint"><span class="mda-idle-dot"></span>Waiting for hierarchy file upload...</div>', unsafe_allow_html=True)
        _nav(2, 4, next_disabled=True)
        return

    summary = st.session_state.hierarchy_summary or {}
    needs_review = bool(summary.get("duplicates") or summary.get("missing_levels"))
    status = "File requires review" if needs_review else "File validated"
    klass = "mda-warn" if needs_review else "mda-ok"
    _file_card(
        st.session_state.hierarchy_name,
        st.session_state.hierarchy_size,
        "Hierarchy CSV" if Path(st.session_state.hierarchy_name or "").suffix.lower() == ".csv" else "Hierarchy Excel",
        status,
        klass,
        f"{summary.get('total_rows', 0)} rows",
    )
    if not needs_review:
        st.info("Hierarchy columns detected. Duplicate codes will be flagged during processing, not guessed.")
    st.caption("Detected levels: " + (", ".join(summary.get("available_levels") or []) or "(none)"))
    if summary.get("missing_levels"):
        st.warning(
            "The hierarchy report is missing one or more network levels: "
            + ", ".join(summary["missing_levels"])
            + ". Records that depend on those identifiers may fail validation."
        )
    if summary.get("duplicates"):
        st.error(
            "Duplicate hierarchy keys were found. Ambiguous codes will be marked for review "
            "and will not be auto-selected."
        )
        st.json(summary["duplicates"])
    with st.expander("Detected hierarchy columns"):
        st.json(summary.get("detected_columns") or {})
    for msg in st.session_state.get("hierarchy_warnings") or []:
        st.warning(msg)
    _nav(2, 4, next_disabled=not st.session_state.hierarchy_path)


def _apply_source_inspection(inspection, selected_sheet: str | None) -> None:
    headers = list(inspection.headers)
    st.session_state.source_headers = headers
    st.session_state.source_row_count = inspection.row_count
    st.session_state.source_preview = inspection.preview
    key_col = find_source_key_column(headers)
    st.session_state.source_key_column = key_col
    mapping = create_column_mapping(headers, st.session_state.template_headers)
    st.session_state.mapped_column_count = len(mapping.mappings)
    st.session_state.unmapped_source_count = len(mapping.unmapped_source)
    st.session_state.source_parsed_key = (st.session_state.source_fp, selected_sheet)
    st.session_state.source_dup_warning = inspection.duplicate_key_count


def step_source() -> None:
    _step_heading(
        4,
        "Upload consumer source data",
        "Upload the Excel workbook or CSV file containing the consumer records to be transformed.",
    )
    uploaded = st.file_uploader(
        "Consumer source data (CSV or Excel)",
        type=["csv", "xlsx", "xls"],
        key="src",
    )
    _upload_caption(".csv, .xlsx, .xls — UTF-8 CSV recommended")
    if uploaded:
        fp = file_fingerprint(uploaded.name, uploaded.size)
        if st.session_state.source_fp != fp:
            try:
                validate_extension(uploaded.name, SOURCE_EXTENSIONS, "consumer source data")
                path = _save_upload(uploaded, Path(uploaded.name).suffix, st.session_state.source_path)
                is_csv = Path(uploaded.name).suffix.lower() == ".csv"
                st.session_state.source_path = path
                st.session_state.source_name = uploaded.name
                st.session_state.source_size = uploaded.size
                st.session_state.source_fp = fp
                st.session_state.source_is_csv = is_csv
                st.session_state.source_error = None
                st.session_state.source_error_detail = None
                st.session_state.mapping_confirmed = False
                st.session_state.custom_map = {}
                st.session_state.source_parsed_key = None
                if is_csv:
                    st.session_state.source_sheets = []
                    st.session_state.source_sheet = None
                else:
                    wb = load_excel_workbook(path)
                    sheets = get_sheet_names(wb)
                    detected = detect_source_sheet(wb)
                    wb.close()
                    st.session_state.source_sheets = sheets
                    st.session_state.source_sheet = detected if detected in sheets else (sheets[0] if sheets else None)
                _clear_results()
            except Exception as exc:
                user_msg, detail = humanize_exception(exc, "source")
                remove_temp_file(st.session_state.get("source_path"))
                st.session_state.source_path = None
                st.session_state.source_fp = fp
                st.session_state.source_is_csv = False
                st.session_state.source_error = user_msg
                st.session_state.source_error_detail = detail
                _clear_results()

    if st.session_state.source_error:
        st.error(st.session_state.source_error)
        _issue_expander("Technical details", st.session_state.source_error_detail or "")
        _nav(3, 5, next_disabled=True)
        return

    if not st.session_state.source_path:
        st.markdown('<div class="mda-idle-hint"><span class="mda-idle-dot"></span>Waiting for consumer source file upload...</div>', unsafe_allow_html=True)
        _nav(3, 5, next_disabled=True)
        return

    selected = None
    if st.session_state.source_is_csv:
        st.info("CSV files contain a single table, so worksheet selection is not required.")
    else:
        sheets = st.session_state.get("source_sheets") or []
        if not sheets:
            st.error("No worksheets were found in the consumer source workbook.")
            _nav(3, 5, next_disabled=True)
            return
        current = st.session_state.source_sheet if st.session_state.source_sheet in sheets else sheets[0]
        selected = st.selectbox("Source worksheet", sheets, index=sheets.index(current))
        if selected != st.session_state.source_sheet:
            st.session_state.source_sheet = selected
            st.session_state.mapping_confirmed = False
            st.session_state.source_parsed_key = None
            _clear_results()
        selected = st.session_state.source_sheet

    parsed_key = (st.session_state.source_fp, selected)
    try:
        if st.session_state.source_parsed_key != parsed_key:
            if st.session_state.source_is_csv:
                inspection = inspect_csv_file(st.session_state.source_path)
            else:
                inspection = inspect_excel_sheet(st.session_state.source_path, selected)
            _apply_source_inspection(inspection, selected)
    except Exception as exc:
        user_msg, detail = humanize_exception(exc, "source")
        st.error(user_msg)
        _issue_expander("Technical details", detail)
        _nav(3, 5, next_disabled=True)
        return

    key_col = st.session_state.source_key_column
    kind = "CSV file" if st.session_state.source_is_csv else "Excel workbook"
    extra = f"{st.session_state.source_row_count} records"
    if selected:
        extra += f" · sheet {selected}"
    if key_col:
        _file_card(
            st.session_state.source_name,
            st.session_state.source_size,
            kind,
            "File validated",
            "mda-ok",
            extra,
        )
        st.info(f"Required consumer columns detected. Business key: `{key_col}`.")
    else:
        _file_card(
            st.session_state.source_name,
            st.session_state.source_size,
            kind,
            "File requires review",
            "mda-warn",
            extra,
        )
        st.error(
            "No consumer number or account number column was detected in the source file. "
            "Add one of those columns, then upload the file again."
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("Source records", st.session_state.source_row_count or 0)
    c2.metric("Mapped columns", st.session_state.mapped_column_count or 0)
    c3.metric("Unmatched source columns", st.session_state.unmapped_source_count or 0)
    if st.session_state.unmapped_source_count:
        st.caption("Unmatched source columns are ignored unless required to derive a target value.")
    preview = st.session_state.source_preview
    if preview is not None and not getattr(preview, "empty", True):
        _preview_frame(preview, "Source preview.")
    if st.session_state.source_dup_warning:
        st.warning(
            f"{st.session_state.source_dup_warning} duplicate identifier(s) in `{key_col}`. "
            "Duplicate keys will be flagged during validation."
        )
    source_ready = bool(st.session_state.source_path) and (
        st.session_state.source_is_csv or bool(st.session_state.source_sheet)
    )
    _nav(3, 5, next_disabled=not source_ready)


def step_mappings() -> None:
    _step_heading(
        5,
        "Review source-to-target mappings",
        "Confirm how source columns map to the template. Hierarchy, reference, and business-rule fields are filled by the engine.",
    )
    if not (st.session_state.source_path and st.session_state.template_headers):
        st.warning("Complete the previous steps first.")
        _nav(4, None)
        return
    source_headers = st.session_state.source_headers or []
    if not source_headers:
        st.warning("Source columns could not be detected. Return to Consumer Data and re-upload the file.")
        _nav(4, None)
        return
    mapping = create_column_mapping(
        source_headers,
        st.session_state.template_headers,
        custom_map=st.session_state.custom_map or None,
    )
    rows = []
    from .config import REFERENCE_CATEGORY_MAP
    for m in mapping.mappings:
        tgt = m.target_column
        lookup = REFERENCE_CATEGORY_MAP.get(tgt.lower().strip(), "")
        if tgt.lower().strip() == "meter_make":
            op = OPERATIONS[st.session_state.operation_id]
            lookup = f"{op.destination.value} METER MAKE"
        if tgt.lower().strip() == "meter_category":
            lookup = "METER CATEGORY (always Main Meter)"
        rows.append({
            "Source column": m.source_column,
            "Target column": tgt,
            "Transformation": m.transformation,
            "Lookup category": lookup,
            "Confidence": m.confidence,
            "Status": "OK" if m.is_confirmed else "Needs confirmation",
        })
    _preview_frame(pd.DataFrame(rows), "Column mappings.")
    st.markdown("**Unmapped source columns**")
    st.write(mapping.unmapped_source or "(none)")
    st.markdown("**Target columns without a direct source mapping**")
    st.write(mapping.unmapped_target or "(none)")
    st.caption(
        "meter_category is always resolved to Main Meter from the template reference sheet. "
        "It is not copied from the source workbook."
    )

    with st.expander("Override a source → target mapping"):
        src = st.selectbox("Source column", source_headers)
        tgt = st.selectbox("Target column", ["(none)"] + st.session_state.template_headers)
        if st.button("Apply override"):
            cmap = dict(st.session_state.custom_map)
            if tgt == "(none)":
                cmap.pop(src, None)
            else:
                cmap[src] = tgt
            st.session_state.custom_map = cmap
            st.session_state.mapping_confirmed = False
            _clear_results()
            st.rerun()

    mandatory_unmapped = [
        t for t in mapping.unmapped_target
        if t.lower().strip() in ENGINE_MANDATORY_SOURCE_TARGETS
        and t.lower().strip() not in ENGINE_POPULATED_TARGETS
    ]
    st.session_state.mapping_blockers = mandatory_unmapped
    if mandatory_unmapped:
        st.error(
            "Mandatory target fields still lack a source mapping: "
            + ", ".join(mandatory_unmapped)
            + ". Map them before continuing."
        )
        st.session_state.mapping_confirmed = False
    else:
        st.session_state.mapping_confirmed = st.checkbox(
            "I confirm the mapping is correct for this run",
            value=st.session_state.mapping_confirmed,
        )
    _nav(4, 6, next_disabled=not st.session_state.mapping_confirmed)


def _run_summary() -> None:
    op = OPERATIONS.get(st.session_state.operation_id)
    rows = [
        ("Selected operation", op.label if op else "Not selected"),
        ("Template", st.session_state.template_name or "Not uploaded"),
        ("Target worksheet", st.session_state.target_sheet or "Not selected"),
        ("Hierarchy report", st.session_state.hierarchy_name or "Not uploaded"),
        ("Consumer source", st.session_state.source_name or "Not uploaded"),
        (
            "Source worksheet",
            "CSV (single table)" if st.session_state.source_is_csv else (st.session_state.source_sheet or "Not selected"),
        ),
    ]
    if st.session_state.source_row_count is not None:
        rows.append(("Detected source records", str(st.session_state.source_row_count)))
    mapping_state = "Confirmed" if st.session_state.mapping_confirmed else "Not confirmed"
    rows.append(("Mapping status", mapping_state))
    st.dataframe(pd.DataFrame(rows, columns=["Item", "Value"]), use_container_width=True, hide_index=True)


def step_dry_run() -> None:
    _step_heading(
        6,
        "Validate the transformation",
        "Review the run summary, then process records without writing the final master workbook until you continue to Export.",
    )
    op = OPERATIONS[st.session_state.operation_id]
    st.markdown("**Run summary**")
    _run_summary()

    estimated_memory = estimate_peak_memory_bytes(
        st.session_state.source_size,
        st.session_state.source_is_csv,
        st.session_state.hierarchy_size,
        st.session_state.template_size,
    )
    available_memory = available_memory_bytes()
    memory_text = f"Estimated peak memory: {format_file_size(estimated_memory)}"
    if available_memory is not None:
        memory_text += f" · currently available: {format_file_size(available_memory)}"
    st.caption(memory_text)
    if available_memory is not None and estimated_memory > available_memory:
        st.warning(
            "This run may exceed currently available memory. The estimate is conservative; "
            "close other applications or use smaller input files if processing fails."
        )

    busy = bool(st.session_state.processing)
    if st.button("Run validation", type="primary", disabled=busy):
        _start_pipeline_job()
        st.rerun()

    if st.session_state.pipeline_job:
        _render_pipeline_job()
        return
    if st.session_state.pipeline_error:
        st.error(st.session_state.pipeline_error)
        if st.session_state.pipeline_error_detail:
            _issue_expander("Technical details", st.session_state.pipeline_error_detail)

    result = st.session_state.pipeline_result
    if result:
        _render_results(result)
        _render_hierarchy_correction_review(result)
        if st.session_state.audit_bytes:
            st.download_button(
                "Download audit report",
                data=st.session_state.audit_bytes,
                file_name=st.session_state.audit_name or "audit_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Per-record status, lookups, and hierarchy outcomes for this run.",
            )
        if st.session_state.error_bytes:
            st.download_button(
                "Download error report",
                data=st.session_state.error_bytes,
                file_name=st.session_state.error_name or "error_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Failed, review, and warning records from this run.",
            )
    _nav(5, 7, next_disabled=st.session_state.pipeline_result is None or st.session_state.processing)


def _render_results(result) -> None:
    summary = result.audit.get_summary()
    counts = action_counts(result.audit.entries)
    status = results_status_label(result)
    if status == "Complete":
        st.success("Complete — all required validations passed.")
    elif status == "Complete with warnings":
        st.warning("Complete with warnings — output can be generated, but noncritical issues remain.")
    elif status == "Failed":
        st.error("Failed — processing could not complete successfully for the submitted records.")
    else:
        st.error("Action required — unresolved issues prevent treating this as a reliable master update.")

    metrics = [
        ("Input records", result.source_row_count),
        ("Successfully processed", summary["successful"]),
        ("Processed with warnings", summary["warnings"]),
        ("Requires review", summary["requires_review"]),
        ("Failed records", summary["failed"]),
    ]
    if counts["inserted"] or counts["updated"]:
        metrics.extend([
            ("Inserted", counts["inserted"]),
            ("Updated", counts["updated"]),
        ])
    cols = st.columns(len(metrics))
    for col, (label, value) in zip(cols, metrics):
        col.metric(label, value)

    tab_issues, tab_audit, tab_info = st.tabs([
        f"Validation Issues & Quality ({len(result.validation.errors) + len(result.validation.warnings)})",
        "Record Audit Log",
        "Informational Messages",
    ])

    with tab_issues:
        has_issues = False
        if result.validation.errors:
            has_issues = True
            st.markdown("**Errors**")
            err_rows = [{
                "File / area": e.category,
                "Column": e.field,
                "Row": e.row if e.row and e.row > 0 else "",
                "Issue": humanize_validation_issue(e.category, e.message, e.field, e.row),
            } for e in result.validation.errors]
            _preview_frame(pd.DataFrame(err_rows), "Validation errors.")
        if result.validation.warnings:
            has_issues = True
            st.markdown("**Warnings**")
            warn_rows = [{
                "File / area": w.category,
                "Column": w.field,
                "Row": w.row if w.row and w.row > 0 else "",
                "Issue": humanize_validation_issue(w.category, w.message, w.field, w.row),
            } for w in result.validation.warnings]
            _preview_frame(pd.DataFrame(warn_rows), "Validation warnings.")
        if not has_issues:
            st.info("No validation issues or warnings detected. All records comply with schema and reference rules.")

    with tab_audit:
        st.markdown("**Record-level audit stream**")
        _preview_frame(result.audit.to_dataframe(), "Record-level audit.")
        st.markdown("**Field-level value traceability**")
        _preview_frame(result.audit.to_field_dataframe(), "Field-level audit.")

    with tab_info:
        infos = [i for i in result.validation.issues if i.severity == "INFO"]
        if infos:
            for item in infos[:50]:
                st.caption(item.message)
        else:
            st.caption("No informational messages.")


def step_output() -> None:
    _step_heading(
        7,
        "Generate and download output",
        "Write a new workbook from the uploaded template. Previous-run files are not reused after inputs change.",
    )
    result = st.session_state.pipeline_result
    if not result:
        st.warning("Run validation first.")
        _nav(6, None)
        return
    op = OPERATIONS[st.session_state.operation_id]
    _render_results(result)
    if not result.ready_for_output:
        st.error(
            "This run is not a successful full master update. Unresolved and failed rows will be "
            "kept out of the master workbook and listed in the review report."
        )
    unresolved = [
        (record, entry)
        for record, entry in zip(result.records, result.audit.entries)
        if entry.status in {"FAILED", "REQUIRES_REVIEW"}
    ]
    eligible = [
        record for record, entry in zip(result.records, result.audit.entries)
        if entry.status not in {"FAILED", "REQUIRES_REVIEW"}
    ]
    st.info(
        f"Master workbook: {len(eligible):,} eligible rows. "
        f"Rows needing review: {len(unresolved):,}."
    )
    make_review_workbook = st.checkbox(
        "Also create a separate review-only workbook for failed and unresolved rows",
        value=False,
        disabled=not unresolved,
    )
    busy = bool(st.session_state.processing)
    if st.button(
        "Generate output workbook",
        type="primary",
        disabled=busy or (not eligible and not make_review_workbook),
    ):
        st.session_state.processing = True
        try:
            with st.status("Writing output workbook…", expanded=True) as status:
                stamp = st.session_state.run_stamp or datetime.now()
                out_name = output_workbook_name(op, stamp)
                if eligible:
                    status.write("Writing the eligible master workbook")
                    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
                    try:
                        write_output_workbook(
                            template_path=st.session_state.template_path,
                            target_sheet=st.session_state.target_sheet,
                            output_records=eligible,
                            output_path=tmp_out,
                            target_headers=result.template_schema.headers,
                            preserve_type_hint_row=result.template_schema.has_type_hint_row,
                        )
                        st.session_state.output_bytes = Path(tmp_out).read_bytes()
                        st.session_state.output_name = out_name
                    finally:
                        remove_temp_file(tmp_out)
                else:
                    st.session_state.output_bytes = None
                    st.session_state.output_name = None
                st.session_state.review_bytes = None
                st.session_state.review_name = None
                if make_review_workbook and unresolved:
                    status.write("Writing the separate review-only workbook")
                    review_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
                    try:
                        write_output_workbook(
                            template_path=st.session_state.template_path,
                            target_sheet=st.session_state.target_sheet,
                            output_records=[record for record, _ in unresolved],
                            output_path=review_path,
                            target_headers=result.template_schema.headers,
                            preserve_type_hint_row=result.template_schema.has_type_hint_row,
                        )
                        st.session_state.review_bytes = Path(review_path).read_bytes()
                        st.session_state.review_name = out_name.replace(
                            ".xlsx", "_REVIEW_ONLY.xlsx"
                        )
                    finally:
                        remove_temp_file(review_path)
                status.write("Preparing download")
                status.update(label="Workbook ready", state="complete")
                if eligible:
                    st.success(f"Created {out_name} with {len(eligible):,} eligible data rows.")
                else:
                    st.success("No rows qualified for the master workbook; the review-only workbook is ready.")
        except MemoryError as exc:
            st.error(
                "The current environment does not have enough memory to write the output workbook. "
                "Free memory and try again."
            )
            _issue_expander("Technical details", str(exc))
            st.session_state.output_bytes = None
            st.session_state.output_name = None
            st.session_state.review_bytes = None
            st.session_state.review_name = None
        except OSError as exc:
            user_msg, detail = humanize_exception(exc, "output")
            st.error(user_msg)
            _issue_expander("Technical details", detail)
            st.session_state.output_bytes = None
            st.session_state.output_name = None
            st.session_state.review_bytes = None
            st.session_state.review_name = None
        except Exception as exc:
            user_msg, detail = humanize_exception(exc, "output")
            st.error(user_msg)
            _issue_expander("Technical details", detail)
            st.session_state.output_bytes = None
            st.session_state.output_name = None
            st.session_state.review_bytes = None
            st.session_state.review_name = None
        finally:
            st.session_state.processing = False

    st.markdown("**Downloads for this run**")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown(
            """
            <div class="mda-export-card mda-export-card-hero">
              <div class="mda-export-title">Master Ingestion Workbook (.xlsx)</div>
              <div class="mda-export-desc">Template-compliant consumer master workbook for the current run.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.session_state.output_bytes:
            st.download_button(
                "Final Excel workbook",
                data=st.session_state.output_bytes,
                file_name=st.session_state.output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Template-compliant consumer master workbook for the current run.",
                use_container_width=True,
            )
            st.caption(f"{st.session_state.output_name} ({format_file_size(len(st.session_state.output_bytes))})")
        else:
            st.caption("Eligible master workbook is available after generation.")
        if st.session_state.review_bytes:
            st.download_button(
                "Review-only workbook",
                data=st.session_state.review_bytes,
                file_name=st.session_state.review_name or "review_only.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Failed and unresolved rows, separated from the master workbook.",
                use_container_width=True,
                key="dl_review_only",
            )
    with d2:
        st.markdown(
            """
            <div class="mda-export-card">
              <div class="mda-export-title">Audit Trail Report (.xlsx)</div>
              <div class="mda-export-desc">Full per-record audit for the current validation run.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.session_state.audit_bytes:
            st.download_button(
                "Audit report",
                data=st.session_state.audit_bytes,
                file_name=st.session_state.audit_name or "audit_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Full per-record audit for the current validation run.",
                use_container_width=True,
                key="dl_audit2",
            )
            st.caption(f"{st.session_state.audit_name or ''} ({format_file_size(len(st.session_state.audit_bytes))})")
    with d3:
        st.markdown(
            """
            <div class="mda-export-card">
              <div class="mda-export-title">Error Report (.xlsx)</div>
              <div class="mda-export-desc">Unresolved issues from the current validation run.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.session_state.error_bytes:
            st.download_button(
                "Error report",
                data=st.session_state.error_bytes,
                file_name=st.session_state.error_name or "error_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Unresolved issues from the current validation run.",
                use_container_width=True,
                key="dl_err2",
            )
            st.caption(f"{st.session_state.error_name or ''} ({format_file_size(len(st.session_state.error_bytes))})")
    _nav(6, None)


def main() -> None:
    st.set_page_config(
        page_title="Master Data Automation",
        page_icon="▦",
        layout="wide",
    )
    _inject_css()
    _init_state()
    st.markdown(
        f"""
        <div class="mda-topbar">
          <div class="mda-topbar-brand">
            <div class="mda-mark">MDA</div>
            <div class="mda-title-wrap">
              <div class="mda-title-row">
                <h1 class="mda-title">Master Data Automation</h1>
                <span class="mda-version-pill">v{__version__}</span>
              </div>
              <p class="mda-kicker">Consumer Master Update System &middot; Transform consumer records, resolve network hierarchy, validate reference data</p>
            </div>
          </div>
          <div class="mda-topbar-meta">
            <span class="mda-runtime-pill"><span class="mda-pulsing-dot"></span>System Ready &middot; Python 3.11+</span>
            <span class="mda-profile-pill">Grid Admin</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    statuses = compute_step_statuses(st.session_state, st.session_state.step)
    st.markdown(stepper_html(statuses), unsafe_allow_html=True)
    dispatch = {
        1: step_template,
        2: step_operation,
        3: step_hierarchy,
        4: step_source,
        5: step_mappings,
        6: step_dry_run,
        7: step_output,
    }
    dispatch[st.session_state.step]()


if __name__ == "__main__":
    main()

