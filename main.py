# main.py
# Unnati Warranty Management Dashboard (No Login / No Change Password / No Logout)
# Features:
# - Loads all Excel files from local Windows paths OR from Render repo directory
# - Shows tables for Credit, Debit, Arbitration, Current Month, Compensation, PR Approval
# - Excel Export works (uses FileResponse + temp file)
# - Includes "Detailed Data" sheets for Credit/Debit/Arbitration (and Pending Arb for Arbitration)
# - Includes Summary + Details for Current Month / Compensation / PR Approval

import os
import io
import socket
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple

import numpy as np
import pandas as pd
import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.background import BackgroundTasks

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# =========================
# ENV / PATHS
# =========================
IS_RENDER = os.getenv("RENDER", "").lower() == "true"
PORT = int(os.getenv("PORT", "8001"))

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", "/mnt/data"))  # Render disk if used

# Your local folder paths (used if files exist there)
LOCAL_FOLDER_1 = Path(r"D:\Power BI New\Warranty Debit")
LOCAL_FOLDER_2 = Path(r"D:\Power BI New\warranty dashboard render")

# Files used by dashboard
FILE_WARRANTY_DEBIT = "Warranty Debit.xlsx"
FILE_TRANSIT = "Transit_Claims_Merged.xlsx"
FILE_PENDING = "Pending Warranty Claim Details.xlsx"
FILE_PR_APPROVAL = "Pr_Approval_Claims_Merged.xlsx"


def find_data_file(filename: str) -> Path:
    """
    Finds a file in a robust way for:
    - Render repo (/opt/render/project/src => BASE_DIR)
    - Optional Render disk (/mnt/data)
    - Local folders you use on Windows
    """
    candidates = [
        BASE_DIR / filename,                  # Render repo root
        BASE_DIR / "data" / filename,
        BASE_DIR / "Data" / filename,
        DATA_DIR / filename,                  # Render disk
        DATA_DIR / "data" / filename,
        LOCAL_FOLDER_1 / filename,            # Local Windows folder
        LOCAL_FOLDER_2 / filename,            # Local Windows folder
        Path(filename),                       # Current working dir fallback
    ]

    for p in candidates:
        try:
            if p.exists():
                print(f"[FOUND] {filename} => {p}")
                return p
        except Exception:
            pass

    raise FileNotFoundError(
        f"{filename} not found. Searched: {[str(x) for x in candidates]}"
    )


# =========================
# GLOBAL DATA STORE
# =========================
WARRANTY_DATA = {
    "credit_df": None,
    "debit_df": None,
    "arbitration_df": None,
    "source_df": None,

    "current_month_df": None,
    "current_month_source_df": None,

    "compensation_df": None,
    "compensation_source_df": None,

    "pr_approval_df": None,
    "pr_approval_source_df": None,
}


# =========================
# HELPERS
# =========================
def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def month_short(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    return s[:3].title() if len(s) >= 3 else s.title()


def format_claim_no(x) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return ""
    try:
        return str(int(float(x)))
    except Exception:
        return str(x).strip()


def format_ro_id_with_prefix(x) -> str:
    if pd.isna(x) or str(x).strip() == "":
        return ""
    try:
        return f"RO{str(int(float(x)))}"
    except Exception:
        v = str(x).strip()
        return v if v.startswith("RO") else f"RO{v}"


def is_empty_or_hyphen(value) -> bool:
    if pd.isna(value):
        return True
    v = str(value).strip()
    if v == "" or v == "-" or v.upper() == "NAN":
        return True
    return False


def has_valid_arb_id(value) -> bool:
    if pd.isna(value):
        return False
    v = str(value).strip().upper()
    return v.startswith("ARB") and v not in ("NAN", "")


# =========================
# PROCESSING FUNCTIONS
# =========================
def process_warranty_data() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Reads Warranty Debit.xlsx and produces:
    - credit_df summary
    - debit_df summary
    - arbitration_df summary (debit amount where Claim arbitration ID starts with ARB)
    - source_df raw
    """
    try:
        path = find_data_file(FILE_WARRANTY_DEBIT)
        df = pd.read_excel(path, sheet_name="Sheet1")
        print(f"[OK] Loaded {FILE_WARRANTY_DEBIT}: rows={len(df)} cols={len(df.columns)}")

        required = ["Dealer Location", "Fiscal Month", "Credit Note Amount", "Debit Note Amount", "Total Claim Amount", "Claim arbitration ID"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in {FILE_WARRANTY_DEBIT}: {missing}")

        dealer_mapping = {
            "AMRAVATI": "AMT",
            "CHAUFULA_SZZ": "CHA",
            "CHIKHALI": "CHI",
            "KOLHAPUR_WS": "KOL",
            "NAGPUR_KAMPTHEE ROAD": "HO",
            "NAGPUR_WARDHAMAN NGR": "CITY",
            "SHIKRAPUR_SZS": "SHI",
            "WAGHOLI": "WAG",
            "YAVATMAL": "YAT",
            "NAGPUR_WARDHAMAN NGR_CQ": "CQ",
        }

        df["Credit Note Amount"] = safe_numeric(df["Credit Note Amount"])
        df["Debit Note Amount"] = safe_numeric(df["Debit Note Amount"])
        df["Total Claim Amount"] = safe_numeric(df["Total Claim Amount"])

        df["Dealer_Code"] = df["Dealer Location"].map(dealer_mapping).fillna(df["Dealer Location"])
        df["Month"] = df["Fiscal Month"].apply(month_short)

        # Ensure arbitration id column is normalized
        df["Claim arbitration ID"] = df["Claim arbitration ID"].astype(str)
        df.loc[df["Claim arbitration ID"].str.lower().isin(["nan", "none"]), "Claim arbitration ID"] = ""

        months = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
        dealers = sorted([d for d in df["Dealer_Code"].dropna().unique().tolist()])

        # CREDIT
        credit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m]
            if md.empty:
                credit_df[f"Credit Note {m}"] = 0
            else:
                s = md.groupby("Dealer_Code")["Credit Note Amount"].sum().reset_index()
                s.columns = ["Division", f"Credit Note {m}"]
                credit_df = credit_df.merge(s, on="Division", how="left")
        credit_df = credit_df.fillna(0)
        credit_cols = [f"Credit Note {m}" for m in months]
        credit_df["Total Credit"] = credit_df[credit_cols].sum(axis=1)

        gt = {"Division": "Grand Total"}
        for c in credit_df.columns[1:]:
            gt[c] = float(credit_df[c].sum())
        credit_df = pd.concat([credit_df, pd.DataFrame([gt])], ignore_index=True)

        # DEBIT
        debit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m]
            if md.empty:
                debit_df[f"Debit Note {m}"] = 0
            else:
                s = md.groupby("Dealer_Code")["Debit Note Amount"].sum().reset_index()
                s.columns = ["Division", f"Debit Note {m}"]
                debit_df = debit_df.merge(s, on="Division", how="left")
        debit_df = debit_df.fillna(0)
        debit_cols = [f"Debit Note {m}" for m in months]
        debit_df["Total Debit"] = debit_df[debit_cols].sum(axis=1)

        gt = {"Division": "Grand Total"}
        for c in debit_df.columns[1:]:
            gt[c] = float(debit_df[c].sum())
        debit_df = pd.concat([debit_df, pd.DataFrame([gt])], ignore_index=True)

        # ARBITRATION
        arbitration_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m].copy()
            if md.empty:
                arbitration_df[f"Claim Arbitration {m}"] = 0
                continue

            md["Is_ARB"] = md["Claim arbitration ID"].apply(has_valid_arb_id)
            md["Arbitration_Amount"] = np.where(md["Is_ARB"], md["Debit Note Amount"], 0.0)
            s = md.groupby("Dealer_Code")["Arbitration_Amount"].sum().reset_index()
            s.columns = ["Division", f"Claim Arbitration {m}"]
            arbitration_df = arbitration_df.merge(s, on="Division", how="left")

        arbitration_df = arbitration_df.fillna(0)

        arb_cols = [f"Claim Arbitration {m}" for m in months]
        total_debit_by_dealer = debit_df[debit_df["Division"] != "Grand Total"][["Division", "Total Debit"]].copy()
        arbitration_df = arbitration_df.merge(total_debit_by_dealer, on="Division", how="left")
        arbitration_df["Pending Claim Arbitration"] = arbitration_df["Total Debit"] - arbitration_df[arb_cols].sum(axis=1)
        arbitration_df = arbitration_df.drop(columns=["Total Debit"])

        gt = {"Division": "Grand Total"}
        for c in arbitration_df.columns[1:]:
            gt[c] = float(arbitration_df[c].sum())
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([gt])], ignore_index=True)

        return credit_df, debit_df, arbitration_df, df

    except Exception as e:
        print(f"[ERROR] process_warranty_data: {e}")
        return None, None, None, None


def process_current_month_warranty() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Reads Pending Warranty Claim Details.xlsx (sheet: Pending Warranty Claim Details)
    Returns summary_df + full source_df
    """
    try:
        path = find_data_file(FILE_PENDING)
        df = pd.read_excel(path, sheet_name="Pending Warranty Claim Details")
        print(f"[OK] Loaded {FILE_PENDING}: rows={len(df)} cols={len(df.columns)}")

        required = ["Division", "Pending Claims Spares", "Pending Claims Labour"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in {FILE_PENDING}: {missing}")

        df["Division"] = df["Division"].astype(str).str.strip()
        df = df[df["Division"].notna() & (df["Division"] != "") & (df["Division"].str.lower() != "nan")]

        summary = []
        for div in sorted(df["Division"].unique()):
            dd = df[df["Division"] == div]
            spares_count = dd["Pending Claims Spares"].notna().sum()
            labour_count = dd["Pending Claims Labour"].notna().sum()
            summary.append({
                "Division": div,
                "Pending Claims Spares Count": int(spares_count),
                "Pending Claims Labour Count": int(labour_count),
                "Total Pending Claims": int(spares_count + labour_count),
            })

        summary_df = pd.DataFrame(summary)
        gt = {
            "Division": "Grand Total",
            "Pending Claims Spares Count": int(summary_df["Pending Claims Spares Count"].sum()) if not summary_df.empty else 0,
            "Pending Claims Labour Count": int(summary_df["Pending Claims Labour Count"].sum()) if not summary_df.empty else 0,
            "Total Pending Claims": int(summary_df["Total Pending Claims"].sum()) if not summary_df.empty else 0,
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([gt])], ignore_index=True)

        return summary_df, df

    except Exception as e:
        print(f"[ERROR] process_current_month_warranty: {e}")
        return None, None


def process_compensation_claim() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Reads Transit_Claims_Merged.xlsx
    Returns summary_df + filtered_df (details source for export)
    """
    try:
        path = find_data_file(FILE_TRANSIT)
        df = pd.read_excel(path)
        print(f"[OK] Loaded {FILE_TRANSIT}: rows={len(df)} cols={len(df.columns)}")

        required_columns = [
            "Division", "RO Id.", "Registration No.", "RO Date", "RO Bill Date",
            "Chassis No.", "Model Group", "Claim Amount", "Request Status",
            "Claim Approved Amt.", "No. of Days"
        ]

        available_columns = [c for c in required_columns if c in df.columns]
        if not available_columns:
            raise ValueError("No required columns found in Transit_Claims_Merged.xlsx")

        df_filtered = df[available_columns].copy()

        if "Division" in df_filtered.columns:
            df_filtered["Division"] = df_filtered["Division"].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered["Division"].notna() & (df_filtered["Division"] != "") & (df_filtered["Division"].str.lower() != "nan")]

        if "RO Id." in df_filtered.columns:
            df_filtered["RO Id."] = df_filtered["RO Id."].apply(format_ro_id_with_prefix)

        for c in ["Claim Amount", "Claim Approved Amt.", "No. of Days"]:
            if c in df_filtered.columns:
                df_filtered[c] = safe_numeric(df_filtered[c])

        summary = []
        if "Division" in df_filtered.columns:
            for div in sorted(df_filtered["Division"].unique()):
                dd = df_filtered[df_filtered["Division"] == div]
                row = {"Division": div, "Total Claims": int(len(dd))}
                if "Claim Amount" in df_filtered.columns:
                    row["Total Claim Amount"] = float(dd["Claim Amount"].sum())
                if "Claim Approved Amt." in df_filtered.columns:
                    row["Total Approved Amount"] = float(dd["Claim Approved Amt."].sum())
                if "No. of Days" in df_filtered.columns:
                    row["Avg No. of Days"] = float(dd["No. of Days"].mean()) if len(dd) else 0.0
                summary.append(row)

        summary_df = pd.DataFrame(summary)
        gt = {"Division": "Grand Total"}
        if not summary_df.empty:
            for c in summary_df.columns:
                if c != "Division":
                    if pd.api.types.is_numeric_dtype(summary_df[c]):
                        gt[c] = float(summary_df[c].sum()) if c != "Avg No. of Days" else float(summary_df[c].mean())
        summary_df = pd.concat([summary_df, pd.DataFrame([gt])], ignore_index=True)

        return summary_df, df_filtered

    except Exception as e:
        print(f"[ERROR] process_compensation_claim: {e}")
        return None, None


def process_pr_approval() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Reads Pr_Approval_Claims_Merged.xlsx
    Returns summary_df + full source df
    """
    try:
        path = find_data_file(FILE_PR_APPROVAL)
        df = pd.read_excel(path)
        print(f"[OK] Loaded {FILE_PR_APPROVAL}: rows={len(df)} cols={len(df.columns)}")

        summary_columns = ["Division", "PA Request No.", "PA Date", "Request Type", "App. Claim Amt from M&M"]
        available_cols = [c for c in summary_columns if c in df.columns]
        if not available_cols:
            raise ValueError("No required columns found in Pr_Approval_Claims_Merged.xlsx")

        display_df = df[available_cols].copy()

        if "Division" in display_df.columns:
            display_df["Division"] = display_df["Division"].astype(str).str.strip()
            display_df = display_df[display_df["Division"].notna() & (display_df["Division"] != "") & (display_df["Division"].str.lower() != "nan")]

        if "App. Claim Amt from M&M" in display_df.columns:
            display_df["App. Claim Amt from M&M"] = safe_numeric(display_df["App. Claim Amt from M&M"])

        summary = []
        if "Division" in display_df.columns:
            for div in sorted(display_df["Division"].unique()):
                dd = display_df[display_df["Division"] == div]
                row = {"Division": div, "Total Requests": int(len(dd))}
                if "App. Claim Amt from M&M" in display_df.columns:
                    row["Total Approved Amount"] = float(dd["App. Claim Amt from M&M"].sum())
                if "Request Type" in display_df.columns:
                    counts = dd["Request Type"].value_counts(dropna=True).to_dict()
                    for k, v in counts.items():
                        kk = str(k).strip()
                        if kk:
                            row[f"{kk} Count"] = int(v)
                summary.append(row)

        summary_df = pd.DataFrame(summary)
        gt = {"Division": "Grand Total"}
        if not summary_df.empty:
            for c in summary_df.columns:
                if c != "Division" and pd.api.types.is_numeric_dtype(summary_df[c]):
                    gt[c] = float(summary_df[c].sum())
        summary_df = pd.concat([summary_df, pd.DataFrame([gt])], ignore_index=True)

        return summary_df, df

    except Exception as e:
        print(f"[ERROR] process_pr_approval: {e}")
        return None, None


# =========================
# EXCEL WRITER HELPERS
# =========================
def wb_styles():
    header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    return header_fill, header_font, border


def write_df_to_sheet(ws, df: pd.DataFrame, header_fill, header_font, border, number_format="#,##0.00"):
    # headers
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # data
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
                cell.value = float(value)
                cell.number_format = number_format
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(value, (datetime, pd.Timestamp)):
                cell.value = value
                cell.number_format = "DD-MM-YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = "" if pd.isna(value) else str(value)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

    # widths
    for col_idx, col_name in enumerate(df.columns, 1):
        series = df[col_name].astype(str).fillna("")
        max_len = max([len(str(col_name))] + series.map(len).tolist()) if len(series) else len(str(col_name))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)


def save_wb_as_temp_file(wb: Workbook, filename: str, background_tasks: BackgroundTasks) -> FileResponse:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp_path = tmp.name
    tmp.close()

    wb.save(tmp_path)

    # cleanup after response
    def _cleanup(path: str):
        try:
            os.remove(path)
        except Exception:
            pass

    background_tasks.add_task(_cleanup, tmp_path)

    return FileResponse(
        tmp_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================
# DASHBOARD HTML (NO LOGIN)
# =========================
DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Unnati Motors Warranty Management Dashboard</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%);
      min-height: 100vh;
    }
    .navbar {
      background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
      color: #fff;
      padding: 16px 0;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      position: sticky;
      top: 0;
      z-index: 100;
    }
    .navbar .container-fluid {
      max-width: 1400px;
      margin: 0 auto;
      display:flex;
      justify-content:space-between;
      align-items:center;
      padding: 0 30px;
    }
    .navbar-brand { font-size: 22px; font-weight: 800; }
    .status { font-weight: 700; opacity: 0.95; }

    .container {
      max-width: 1400px;
      margin: 30px auto;
      padding: 0 20px;
    }

    .dashboard-content {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      padding: 30px;
    }

    .nav-tabs {
      border-bottom: 2px solid #FF8C00;
      margin-bottom: 20px;
      display:flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .nav-tabs .nav-link {
      color: #666;
      font-weight: 700;
      border: none;
      border-bottom: 3px solid transparent;
      padding: 10px 14px;
      cursor: pointer;
      background: transparent;
      transition: all 0.2s ease;
      border-radius: 8px 8px 0 0;
    }
    .nav-tabs .nav-link:hover {
      color: #FF8C00;
      border-bottom-color: #FF8C00;
    }
    .nav-tabs .nav-link.active {
      color: #FF8C00;
      border-bottom-color: #FF8C00;
    }

    .tab-content { display:none; }
    .tab-content.active { display:block; }

    .table-title {
      font-size: 15px;
      font-weight: 800;
      color: #FF8C00;
      margin-bottom: 10px;
    }

    .table-wrapper { overflow-x:auto; }

    .data-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
      font-size: 12px;
    }
    .data-table thead th {
      background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
      color: #fff;
      padding: 10px 10px;
      text-align: center;
      font-weight: 800;
      font-size: 11px;
      border: none;
      white-space: nowrap;
    }
    .data-table tbody td {
      padding: 9px 10px;
      border-bottom: 1px solid #eee;
      text-align: right;
      white-space: nowrap;
    }
    .data-table tbody td:first-child {
      text-align: left;
      font-weight: 800;
      color: #333;
    }
    .data-table tbody tr:hover { background: #fafafa; }
    .data-table tbody tr:last-child {
      background: #fff8f3;
      font-weight: 900;
      border-top: 2px solid #FF8C00;
      border-bottom: 2px solid #FF8C00;
    }
    .data-table tbody tr:last-child td { color: #FF8C00; }

    .export-section {
      margin: 18px 0 26px 0;
      padding: 16px;
      background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%);
      border-radius: 10px;
      border-left: 5px solid #FF8C00;
      box-shadow: 0 2px 8px rgba(255,140,0,0.1);
    }
    .export-section h3 {
      color: #FF8C00;
      margin-bottom: 10px;
      font-size: 15px;
      font-weight: 900;
    }
    .export-controls {
      display:flex;
      gap: 12px;
      align-items:center;
      flex-wrap:wrap;
      background:#fff;
      padding: 12px;
      border-radius: 8px;
    }
    .export-control-group {
      display:flex;
      gap: 8px;
      align-items:center;
    }
    .export-control-group label {
      font-weight: 800;
      color:#333;
      font-size: 13px;
      min-width: 75px;
    }
    .export-control-group select {
      padding: 8px 10px;
      border: 2px solid #FF8C00;
      border-radius: 6px;
      cursor: pointer;
      background: #fff;
      font-size: 13px;
      min-width: 170px;
      font-weight: 700;
    }
    .export-btn {
      padding: 10px 18px;
      background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
      color:#fff;
      border:none;
      border-radius: 6px;
      cursor:pointer;
      font-weight: 900;
      font-size: 13px;
    }
    .export-btn:disabled {
      background:#ccc;
      cursor:not-allowed;
    }

    .loading {
      text-align:center;
      padding: 24px;
      color:#666;
      font-weight: 700;
    }
    .error {
      text-align:center;
      padding: 20px;
      color:#c62828;
      font-weight: 800;
    }
  </style>
</head>
<body>
  <nav class="navbar">
    <div class="container-fluid">
      <div class="navbar-brand">Unnati Motors Warranty Management Dashboard</div>
      <div class="status" id="statusText">Ready</div>
    </div>
  </nav>

  <div class="container">
    <div class="dashboard-content">
      <div class="loading" id="loadingBox">Loading warranty data...</div>
      <div class="error" id="errorBox" style="display:none;"></div>

      <div id="warrantyTabs" style="display:none;">
        <div class="nav-tabs">
          <button class="nav-link active" onclick="switchTab(event,'credit')">Warranty Credit</button>
          <button class="nav-link" onclick="switchTab(event,'debit')">Warranty Debit</button>
          <button class="nav-link" onclick="switchTab(event,'arbitration')">Claim Arbitration</button>
          <button class="nav-link" onclick="switchTab(event,'currentmonth')">Current Month Warranty</button>
          <button class="nav-link" onclick="switchTab(event,'compensation')">Compensation Claim</button>
          <button class="nav-link" onclick="switchTab(event,'pr_approval')">PR Approval</button>
        </div>

        <div class="export-section">
          <h3>Export to Excel</h3>
          <div class="export-controls">
            <div class="export-control-group">
              <label for="divisionFilter">Division</label>
              <select id="divisionFilter">
                <option value="All">All Divisions</option>
              </select>
            </div>

            <div class="export-control-group">
              <label for="exportType">Type</label>
              <select id="exportType">
                <option value="credit">Credit Note</option>
                <option value="debit">Debit Note</option>
                <option value="arbitration">Claim Arbitration</option>
                <option value="currentmonth">Current Month Warranty</option>
                <option value="compensation">Compensation Claim</option>
                <option value="pr_approval">PR Approval</option>
              </select>
            </div>

            <button onclick="exportToExcel()" class="export-btn" id="exportBtn">Export to Excel</button>
          </div>
        </div>

        <div id="credit" class="tab-content active">
          <div class="table-title">Warranty Credit Note by Division & Month</div>
          <div class="table-wrapper">
            <table class="data-table" id="creditTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

        <div id="debit" class="tab-content">
          <div class="table-title">Warranty Debit Note by Division & Month</div>
          <div class="table-wrapper">
            <table class="data-table" id="debitTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

        <div id="arbitration" class="tab-content">
          <div class="table-title">Warranty Claim Arbitration by Division</div>
          <div class="table-wrapper">
            <table class="data-table" id="arbitrationTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

        <div id="currentmonth" class="tab-content">
          <div class="table-title">Current Month Warranty - Pending Claims Summary</div>
          <div class="table-wrapper">
            <table class="data-table" id="currentMonthTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

        <div id="compensation" class="tab-content">
          <div class="table-title">Compensation Claim - Transit Claims Summary</div>
          <div class="table-wrapper">
            <table class="data-table" id="compensationTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

        <div id="pr_approval" class="tab-content">
          <div class="table-title">PR Approval - Claims Summary</div>
          <div class="table-wrapper">
            <table class="data-table" id="prApprovalTable"><thead></thead><tbody></tbody></table>
          </div>
        </div>

      </div>
    </div>
  </div>

<script>
  let warrantyData = {};

  function switchTab(ev, tabName) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.nav-link').forEach(b => b.classList.remove('active'));
    document.getElementById(tabName).classList.add('active');
    ev.target.classList.add('active');
  }

  function renderTable(tableId, data, decimals=0) {
    const table = document.getElementById(tableId);
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    if (!data || data.length === 0) {
      thead.innerHTML = '';
      tbody.innerHTML = '<tr><td style="text-align:left;" colspan="50">No data</td></tr>';
      return;
    }

    const headers = Object.keys(data[0]);
    thead.innerHTML = '<tr>' + headers.map(h => '<th>' + h + '</th>').join('') + '</tr>';

    tbody.innerHTML = data.map(row => {
      return '<tr>' + headers.map(h => {
        const v = row[h];
        if (typeof v === 'number') {
          return '<td>' + v.toLocaleString('en-IN', { maximumFractionDigits: decimals }) + '</td>';
        }
        return '<td style="text-align:left;">' + (v ?? '') + '</td>';
      }).join('') + '</tr>';
    }).join('');
  }

  function loadDivisions() {
    const divisionSelect = document.getElementById('divisionFilter');
    const exportType = document.getElementById('exportType').value;

    let dataSource = warrantyData.credit || [];
    if (exportType === 'debit') dataSource = warrantyData.debit || [];
    if (exportType === 'arbitration') dataSource = warrantyData.arbitration || [];
    if (exportType === 'currentmonth') dataSource = warrantyData.currentMonth || [];
    if (exportType === 'compensation') dataSource = warrantyData.compensation || [];
    if (exportType === 'pr_approval') dataSource = warrantyData.prApproval || [];

    const divs = new Set();
    dataSource.forEach(r => {
      if (r.Division && r.Division !== 'Grand Total') divs.add(r.Division);
    });

    const current = divisionSelect.value;
    divisionSelect.innerHTML = '<option value="All">All Divisions</option>';
    Array.from(divs).sort().forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      divisionSelect.appendChild(opt);
    });

    if (current && divisionSelect.querySelector('option[value="' + current + '"]')) {
      divisionSelect.value = current;
    }
  }

  document.getElementById('exportType').addEventListener('change', loadDivisions);

  async function loadDashboard() {
    const loadingBox = document.getElementById('loadingBox');
    const errorBox = document.getElementById('errorBox');
    const tabs = document.getElementById('warrantyTabs');

    loadingBox.style.display = 'block';
    errorBox.style.display = 'none';
    tabs.style.display = 'none';

    try {
      const res = await fetch('/api/warranty-data', { method: 'GET' });
      if (!res.ok) {
        const txt = await res.text();
        throw new Error('Failed to load data: HTTP ' + res.status + ' ' + txt);
      }
      warrantyData = await res.json();

      renderTable('creditTable', warrantyData.credit, 0);
      renderTable('debitTable', warrantyData.debit, 0);
      renderTable('arbitrationTable', warrantyData.arbitration, 0);
      renderTable('currentMonthTable', warrantyData.currentMonth, 0);
      renderTable('compensationTable', warrantyData.compensation, 2);
      renderTable('prApprovalTable', warrantyData.prApproval, 2);

      loadDivisions();

      loadingBox.style.display = 'none';
      tabs.style.display = 'block';
      document.getElementById('statusText').textContent = 'Ready';
    } catch (e) {
      loadingBox.style.display = 'none';
      errorBox.style.display = 'block';
      errorBox.textContent = e.message || String(e);
      document.getElementById('statusText').textContent = 'Error';
    }
  }

  async function exportToExcel() {
    const division = document.getElementById('divisionFilter').value || 'All';
    const type = document.getElementById('exportType').value;
    const btn = document.getElementById('exportBtn');

    btn.disabled = true;
    btn.textContent = 'Exporting...';

    try {
      const res = await fetch('/api/export-to-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ division, type })
      });

      if (!res.ok) {
        const msg = await res.text();
        throw new Error('Export failed: HTTP ' + res.status + ' ' + msg);
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = type + '_' + division + '_' + new Date().toISOString().split('T')[0] + '.xlsx';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (e) {
      alert(e.message || String(e));
    } finally {
      btn.disabled = false;
      btn.textContent = 'Export to Excel';
    }
  }

  window.onload = loadDashboard;
</script>
</body>
</html>
"""


# =========================
# FASTAPI APP
# =========================
app = FastAPI()


@app.get("/")
async def root():
    return HTMLResponse(DASHBOARD_HTML)


@app.get("/api/warranty-data")
async def api_warranty_data():
    """
    No login required. Returns all summary tables.
    """
    # if data not loaded, return empty arrays
    def to_records(df):
        if df is None or df.empty:
            return []
        recs = df.to_dict("records")
        # replace NaN with 0 or ""
        for r in recs:
            for k, v in list(r.items()):
                if pd.isna(v):
                    r[k] = 0
        return recs

    return JSONResponse({
        "credit": to_records(WARRANTY_DATA["credit_df"]),
        "debit": to_records(WARRANTY_DATA["debit_df"]),
        "arbitration": to_records(WARRANTY_DATA["arbitration_df"]),
        "currentMonth": to_records(WARRANTY_DATA["current_month_df"]),
        "compensation": to_records(WARRANTY_DATA["compensation_df"]),
        "prApproval": to_records(WARRANTY_DATA["pr_approval_df"]),
    })


@app.post("/api/export-to-excel")
async def export_to_excel(request: Request, background_tasks: BackgroundTasks):
    """
    Export selected division data to Excel with summary and detailed sheets.
    Supports types:
    credit, debit, arbitration, currentmonth, compensation, pr_approval
    """
    body = await request.json()
    selected_division = body.get("division", "All")
    export_type = body.get("type", "credit")

    if export_type not in ["credit", "debit", "arbitration", "currentmonth", "compensation", "pr_approval"]:
        raise HTTPException(status_code=400, detail="Invalid export type")

    # Dispatch for special exports
    if export_type == "currentmonth":
        return await export_current_month(selected_division, background_tasks)
    if export_type == "compensation":
        return await export_compensation(selected_division, background_tasks)
    if export_type == "pr_approval":
        return await export_prapproval(selected_division, background_tasks)

    # Warranty summary exports
    if export_type == "credit":
        df = WARRANTY_DATA["credit_df"]
    elif export_type == "debit":
        df = WARRANTY_DATA["debit_df"]
    else:
        df = WARRANTY_DATA["arbitration_df"]

    if df is None or df.empty:
        raise HTTPException(status_code=500, detail="No data available for export")

    # Filter for selected division (keep grand total row)
    if selected_division not in ("All", "Grand Total"):
        df_export = df[df["Division"] == selected_division].copy()
        gt = df[df["Division"] == "Grand Total"]
        if not gt.empty:
            df_export = pd.concat([df_export, gt], ignore_index=True)
    else:
        df_export = df.copy()

    wb = Workbook()
    header_fill, header_font, border = wb_styles()

    # Sheet 1: Summary
    ws1 = wb.active
    ws1.title = f"{export_type.capitalize()}"
    if selected_division not in ("All", "Grand Total"):
        ws1.title = f"{selected_division}-{export_type.capitalize()}"[:31]

    write_df_to_sheet(ws1, df_export, header_fill, header_font, border, number_format="#,##0.00")

    # Detailed sheets only if a specific division selected
    if selected_division not in ("All", "Grand Total"):
        source_df = WARRANTY_DATA["source_df"]
        if source_df is not None and not source_df.empty:
            dealer_mapping = {
                "AMRAVATI": "AMT",
                "CHAUFULA_SZZ": "CHA",
                "CHIKHALI": "CHI",
                "KOLHAPUR_WS": "KOL",
                "NAGPUR_KAMPTHEE ROAD": "HO",
                "NAGPUR_WARDHAMAN NGR": "CITY",
                "SHIKRAPUR_SZS": "SHI",
                "WAGHOLI": "WAG",
                "YAVATMAL": "YAT",
                "NAGPUR_WARDHAMAN NGR_CQ": "CQ",
            }
            reverse_mapping = {v: k for k, v in dealer_mapping.items()}
            dealer_location = reverse_mapping.get(selected_division)

            if dealer_location and "Dealer Location" in source_df.columns:
                detail_df = source_df[source_df["Dealer Location"] == dealer_location].copy()

                # Apply type-specific filters and columns
                required = [
                    "Fiscal Month",
                    "Dealer Location",
                    "Claim arbitration ID",
                    "Claim Invoice Date",
                    "Claim No",
                    "Claim Date",
                    "Chassis No",
                    "Ro Id",
                    "Claim Type",
                ]

                if export_type == "credit":
                    # credit: Credit Note Amount > 0 and arbitration id empty
                    if "Credit Note Amount" in detail_df.columns:
                        detail_df["Credit Note Amount"] = safe_numeric(detail_df["Credit Note Amount"])
                        detail_df = detail_df[detail_df["Credit Note Amount"] > 0].copy()
                    if "Claim arbitration ID" in detail_df.columns:
                        detail_df = detail_df[detail_df["Claim arbitration ID"].apply(is_empty_or_hyphen)].copy()
                    required.append("Credit Note Amount")

                elif export_type == "debit":
                    if "Debit Note Amount" in detail_df.columns:
                        detail_df["Debit Note Amount"] = safe_numeric(detail_df["Debit Note Amount"])
                        detail_df = detail_df[detail_df["Debit Note Amount"] > 0].copy()
                    required.append("Debit Note Amount")

                else:  # arbitration
                    if "Debit Note Amount" in detail_df.columns:
                        detail_df["Debit Note Amount"] = safe_numeric(detail_df["Debit Note Amount"])
                        detail_df = detail_df[detail_df["Debit Note Amount"] > 0].copy()
                    if "Claim arbitration ID" in detail_df.columns:
                        detail_df = detail_df[detail_df["Claim arbitration ID"].apply(has_valid_arb_id)].copy()
                    required.extend(["Debit Note Amount", "Credit Note Amount"])

                available = [c for c in required if c in detail_df.columns]
                detail_df = detail_df[available].copy()

                if "Claim No" in detail_df.columns:
                    detail_df["Claim No"] = detail_df["Claim No"].apply(format_claim_no)
                if "Ro Id" in detail_df.columns:
                    detail_df["Ro Id"] = detail_df["Ro Id"].apply(format_ro_id_with_prefix)

                # Sort by fiscal month order
                month_order = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
                if "Fiscal Month" in detail_df.columns:
                    detail_df["__m"] = detail_df["Fiscal Month"].apply(month_short)
                    detail_df["__o"] = detail_df["__m"].apply(lambda x: month_order.index(x) if x in month_order else 999)
                    detail_df = detail_df.sort_values("__o").drop(columns=["__m","__o"], errors="ignore")

                ws2 = wb.create_sheet(title=f"{selected_division}-Details"[:31])
                write_df_to_sheet(ws2, detail_df, header_fill, header_font, border, number_format="#,##0.00")

                # For arbitration export: add pending arbitration sheet
                if export_type == "arbitration":
                    pending_df = source_df[source_df["Dealer Location"] == dealer_location].copy()
                    if "Debit Note Amount" in pending_df.columns:
                        pending_df["Debit Note Amount"] = safe_numeric(pending_df["Debit Note Amount"])
                        pending_df = pending_df[pending_df["Debit Note Amount"] > 0].copy()
                    if "Claim arbitration ID" in pending_df.columns:
                        pending_df = pending_df[pending_df["Claim arbitration ID"].apply(is_empty_or_hyphen)].copy()

                    pending_cols = [
                        "Fiscal Month","Dealer Location","Claim arbitration ID","Claim Invoice Date",
                        "Claim No","Claim Date","Chassis No","Ro Id","Claim Type",
                        "Credit Note Amount","Debit Note Amount"
                    ]
                    pending_avail = [c for c in pending_cols if c in pending_df.columns]
                    pending_df = pending_df[pending_avail].copy()

                    if "Claim No" in pending_df.columns:
                        pending_df["Claim No"] = pending_df["Claim No"].apply(format_claim_no)
                    if "Ro Id" in pending_df.columns:
                        pending_df["Ro Id"] = pending_df["Ro Id"].apply(format_ro_id_with_prefix)

                    if "Debit Note Amount" in pending_df.columns:
                        pending_df = pending_df.rename(columns={"Debit Note Amount": "Pending Arbitration Amount"})

                    if "Fiscal Month" in pending_df.columns:
                        month_order = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
                        pending_df["__m"] = pending_df["Fiscal Month"].apply(month_short)
                        pending_df["__o"] = pending_df["__m"].apply(lambda x: month_order.index(x) if x in month_order else 999)
                        pending_df = pending_df.sort_values("__o").drop(columns=["__m","__o"], errors="ignore")

                    ws3 = wb.create_sheet(title=f"{selected_division}-Pending"[:31])
                    write_df_to_sheet(ws3, pending_df, header_fill, header_font, border, number_format="#,##0.00")

    filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return save_wb_as_temp_file(wb, filename, background_tasks)


async def export_current_month(selected_division: str, background_tasks: BackgroundTasks):
    summary_df = WARRANTY_DATA["current_month_df"]
    source_df = WARRANTY_DATA["current_month_source_df"]

    if summary_df is None or summary_df.empty:
        raise HTTPException(status_code=500, detail="No current month warranty data available")

    if selected_division not in ("All", "Grand Total"):
        df_export = summary_df[summary_df["Division"] == selected_division].copy()
        gt = summary_df[summary_df["Division"] == "Grand Total"]
        if not gt.empty:
            df_export = pd.concat([df_export, gt], ignore_index=True)
    else:
        df_export = summary_df.copy()

    wb = Workbook()
    header_fill, header_font, border = wb_styles()

    ws1 = wb.active
    ws1.title = "CurrentMonth" if selected_division in ("All", "Grand Total") else f"{selected_division}-Summary"[:31]
    write_df_to_sheet(ws1, df_export, header_fill, header_font, border, number_format="#,##0")

    # spares + labour detail sheets
    if source_df is not None and not source_df.empty:
        df_src = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            df_src = df_src[df_src["Division"] == selected_division].copy()

        # Spares
        if "Pending Claims Spares" in df_src.columns:
            spares_df = df_src[df_src["Pending Claims Spares"].notna()].copy()
            if not spares_df.empty:
                ws2 = wb.create_sheet(title=("Spares" if selected_division in ("All","Grand Total") else f"{selected_division}-Spares")[:31])
                write_df_to_sheet(ws2, spares_df, header_fill, header_font, border, number_format="#,##0.00")

        # Labour
        if "Pending Claims Labour" in df_src.columns:
            labour_df = df_src[df_src["Pending Claims Labour"].notna()].copy()
            if not labour_df.empty:
                ws3 = wb.create_sheet(title=("Labour" if selected_division in ("All","Grand Total") else f"{selected_division}-Labour")[:31])
                write_df_to_sheet(ws3, labour_df, header_fill, header_font, border, number_format="#,##0.00")

    filename = f"{selected_division}_CurrentMonthWarranty_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return save_wb_as_temp_file(wb, filename, background_tasks)


async def export_compensation(selected_division: str, background_tasks: BackgroundTasks):
    summary_df = WARRANTY_DATA["compensation_df"]
    source_df = WARRANTY_DATA["compensation_source_df"]

    if summary_df is None or summary_df.empty:
        raise HTTPException(status_code=500, detail="No compensation claim data available")

    if selected_division not in ("All", "Grand Total"):
        df_export = summary_df[summary_df["Division"] == selected_division].copy()
        gt = summary_df[summary_df["Division"] == "Grand Total"]
        if not gt.empty:
            df_export = pd.concat([df_export, gt], ignore_index=True)
    else:
        df_export = summary_df.copy()

    wb = Workbook()
    header_fill, header_font, border = wb_styles()

    ws1 = wb.active
    ws1.title = "CompSummary" if selected_division in ("All", "Grand Total") else f"{selected_division}-Summary"[:31]
    write_df_to_sheet(ws1, df_export, header_fill, header_font, border, number_format="#,##0.00")

    if source_df is not None and not source_df.empty:
        detail_df = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            detail_df = detail_df[detail_df["Division"] == selected_division].copy()

        ws2 = wb.create_sheet(title=("CompDetails" if selected_division in ("All","Grand Total") else f"{selected_division}-Details")[:31])
        write_df_to_sheet(ws2, detail_df, header_fill, header_font, border, number_format="#,##0.00")

    filename = f"{selected_division}_CompensationClaim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return save_wb_as_temp_file(wb, filename, background_tasks)


async def export_prapproval(selected_division: str, background_tasks: BackgroundTasks):
    summary_df = WARRANTY_DATA["pr_approval_df"]
    source_df = WARRANTY_DATA["pr_approval_source_df"]

    if summary_df is None or summary_df.empty:
        raise HTTPException(status_code=500, detail="No PR Approval data available")

    if selected_division not in ("All", "Grand Total"):
        df_export = summary_df[summary_df["Division"] == selected_division].copy()
        gt = summary_df[summary_df["Division"] == "Grand Total"]
        if not gt.empty:
            df_export = pd.concat([df_export, gt], ignore_index=True)
    else:
        df_export = summary_df.copy()

    wb = Workbook()
    header_fill, header_font, border = wb_styles()

    ws1 = wb.active
    ws1.title = "PRSummary" if selected_division in ("All", "Grand Total") else f"{selected_division}-Summary"[:31]
    write_df_to_sheet(ws1, df_export, header_fill, header_font, border, number_format="#,##0.00")

    if source_df is not None and not source_df.empty:
        detail_df = source_df.copy()
        if selected_division not in ("All", "Grand Total") and "Division" in detail_df.columns:
            detail_df = detail_df[detail_df["Division"].astype(str).str.strip() == selected_division].copy()

        ws2 = wb.create_sheet(title=("PRDetails" if selected_division in ("All","Grand Total") else f"{selected_division}-Details")[:31])
        write_df_to_sheet(ws2, detail_df, header_fill, header_font, border, number_format="#,##0.00")

    filename = f"{selected_division}_PrApproval_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return save_wb_as_temp_file(wb, filename, background_tasks)


# =========================
# STARTUP LOAD
# =========================
def load_all_data():
    print("=" * 90)
    print("LOADING WARRANTY DATA FILES")
    print("=" * 90)

    WARRANTY_DATA["credit_df"], WARRANTY_DATA["debit_df"], WARRANTY_DATA["arbitration_df"], WARRANTY_DATA["source_df"] = process_warranty_data()
    WARRANTY_DATA["current_month_df"], WARRANTY_DATA["current_month_source_df"] = process_current_month_warranty()
    WARRANTY_DATA["compensation_df"], WARRANTY_DATA["compensation_source_df"] = process_compensation_claim()
    WARRANTY_DATA["pr_approval_df"], WARRANTY_DATA["pr_approval_source_df"] = process_pr_approval()

    def status(name, df):
        if df is None:
            print(f"[FAIL] {name}: None")
        else:
            print(f"[OK]   {name}: rows={len(df)}")

    status("Credit Summary", WARRANTY_DATA["credit_df"])
    status("Debit Summary", WARRANTY_DATA["debit_df"])
    status("Arbitration Summary", WARRANTY_DATA["arbitration_df"])
    status("Current Month Summary", WARRANTY_DATA["current_month_df"])
    status("Compensation Summary", WARRANTY_DATA["compensation_df"])
    status("PR Approval Summary", WARRANTY_DATA["pr_approval_df"])

    print("=" * 90)


load_all_data()


if __name__ == "__main__":
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = "127.0.0.1"

    print("=" * 90)
    print("SERVER READY - WARRANTY DASHBOARD (NO LOGIN)")
    print("=" * 90)
    print(f"PORT: {PORT}")
    print(f"Local URL:   http://localhost:{PORT}/")
    print(f"Network URL: http://{local_ip}:{PORT}/")
    print("=" * 90)

    uvicorn.run(app, host="0.0.0.0", port=PORT)
