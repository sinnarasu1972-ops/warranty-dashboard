import pandas as pd
import numpy as np
from datetime import datetime
import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
import os
import socket
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ==================== WARRANTY DATA PROCESSING ====================

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

def process_pr_approval():
    input_path = r"D:\Power BI New\warranty dashboard render\Pr_Approval_Claims_Merged.xlsx"
    try:
        df = pd.read_excel(input_path)

        summary_columns = [
            "Division", "PA Request No.", "PA Date", "Request Type", "App. Claim Amt from M&M"
        ]
        available_summary_columns = [c for c in summary_columns if c in df.columns]
        if not available_summary_columns:
            return None, None

        df_summary_display = df[available_summary_columns].copy()

        if "Division" in df_summary_display.columns:
            df_summary_display["Division"] = df_summary_display["Division"].astype(str).str.strip()
            df_summary_display = df_summary_display[
                df_summary_display["Division"].notna()
                & (df_summary_display["Division"] != "")
                & (df_summary_display["Division"].str.lower() != "nan")
            ]

        if "App. Claim Amt from M&M" in df_summary_display.columns:
            df_summary_display["App. Claim Amt from M&M"] = pd.to_numeric(
                df_summary_display["App. Claim Amt from M&M"], errors="coerce"
            ).fillna(0)

        summary_data = []
        if "Division" in df_summary_display.columns:
            for division in sorted(df_summary_display["Division"].unique()):
                div_data = df_summary_display[df_summary_display["Division"] == division]
                row = {"Division": division}
                row["Total Requests"] = len(div_data)

                if "App. Claim Amt from M&M" in df_summary_display.columns:
                    row["Total Approved Amount"] = float(div_data["App. Claim Amt from M&M"].sum())

                if "Request Type" in df_summary_display.columns:
                    request_types = div_data["Request Type"].value_counts(dropna=True).to_dict()
                    for req_type, count in request_types.items():
                        if pd.notna(req_type) and str(req_type).strip() != "":
                            row[f"{str(req_type).strip()} Count"] = int(count)

                summary_data.append(row)

            summary_df = pd.DataFrame(summary_data)

            grand_total = {"Division": "Grand Total"}
            for col in summary_df.columns:
                if col != "Division" and pd.api.types.is_numeric_dtype(summary_df[col]):
                    grand_total[col] = float(summary_df[col].sum())

            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        return summary_df, df

    except FileNotFoundError:
        return None, None
    except Exception:
        return None, None


def process_compensation_claim():
    input_path = r"D:\Power BI New\Warranty Debit\Transit_Claims_Merged.xlsx"
    try:
        df = pd.read_excel(input_path)

        required_columns = [
            "Division", "RO Id.", "Registration No.", "RO Date", "RO Bill Date",
            "Chassis No.", "Model Group", "Claim Amount", "Request Status",
            "Claim Approved Amt.", "No. of Days"
        ]
        available_columns = [c for c in required_columns if c in df.columns]
        if not available_columns:
            return None, None

        df_filtered = df[available_columns].copy()

        if "Division" in df_filtered.columns:
            df_filtered["Division"] = df_filtered["Division"].astype(str).str.strip()
            df_filtered = df_filtered[
                df_filtered["Division"].notna()
                & (df_filtered["Division"] != "")
                & (df_filtered["Division"].str.lower() != "nan")
            ]

        if "RO Id." in df_filtered.columns:
            def format_ro_id(x):
                if pd.isna(x) or str(x).strip() == "":
                    return ""
                try:
                    return f"RO{str(int(float(x)))}"
                except Exception:
                    s = str(x).strip()
                    return s if s.startswith("RO") else f"RO{s}"
            df_filtered["RO Id."] = df_filtered["RO Id."].apply(format_ro_id)

        for col in ["Claim Amount", "Claim Approved Amt.", "No. of Days"]:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce").fillna(0)

        summary_data = []
        if "Division" in df_filtered.columns:
            for division in sorted(df_filtered["Division"].unique()):
                div_data = df_filtered[df_filtered["Division"] == division]
                row = {"Division": division}
                row["Total Claims"] = int(len(div_data))
                if "Claim Amount" in df_filtered.columns:
                    row["Total Claim Amount"] = float(div_data["Claim Amount"].sum())
                if "Claim Approved Amt." in df_filtered.columns:
                    row["Total Approved Amount"] = float(div_data["Claim Approved Amt."].sum())
                if "No. of Days" in df_filtered.columns:
                    row["Avg No. of Days"] = float(div_data["No. of Days"].mean()) if len(div_data) else 0
                summary_data.append(row)

            summary_df = pd.DataFrame(summary_data)
            grand_total = {"Division": "Grand Total"}
            if "Total Claims" in summary_df.columns:
                grand_total["Total Claims"] = int(summary_df["Total Claims"].sum())
            if "Total Claim Amount" in summary_df.columns:
                grand_total["Total Claim Amount"] = float(summary_df["Total Claim Amount"].sum())
            if "Total Approved Amount" in summary_df.columns:
                grand_total["Total Approved Amount"] = float(summary_df["Total Approved Amount"].sum())
            if "Avg No. of Days" in summary_df.columns:
                grand_total["Avg No. of Days"] = float(summary_df["Avg No. of Days"].mean())
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        return summary_df, df_filtered

    except FileNotFoundError:
        return None, None
    except Exception:
        return None, None


def process_current_month_warranty():
    input_path = r"D:\Power BI New\Warranty Debit\Pending Warranty Claim Details.xlsx"
    try:
        df = pd.read_excel(input_path, sheet_name="Pending Warranty Claim Details")

        required_columns = ["Division", "Pending Claims Spares", "Pending Claims Labour"]
        if any(c not in df.columns for c in required_columns):
            return None, None

        df["Division"] = df["Division"].astype(str).str.strip()
        df = df[
            df["Division"].notna()
            & (df["Division"] != "")
            & (df["Division"].str.lower() != "nan")
        ]

        summary_data = []
        for division in sorted(df["Division"].unique()):
            div_data = df[df["Division"] == division]
            spares_count = int(div_data["Pending Claims Spares"].notna().sum())
            labour_count = int(div_data["Pending Claims Labour"].notna().sum())
            summary_data.append({
                "Division": division,
                "Pending Claims Spares Count": spares_count,
                "Pending Claims Labour Count": labour_count,
                "Total Pending Claims": spares_count + labour_count
            })

        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            "Division": "Grand Total",
            "Pending Claims Spares Count": int(summary_df["Pending Claims Spares Count"].sum()) if not summary_df.empty else 0,
            "Pending Claims Labour Count": int(summary_df["Pending Claims Labour Count"].sum()) if not summary_df.empty else 0,
            "Total Pending Claims": int(summary_df["Total Pending Claims"].sum()) if not summary_df.empty else 0
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

        return summary_df, df

    except FileNotFoundError:
        return None, None
    except Exception:
        return None, None


def process_warranty_data():
    input_path = r"D:\Power BI New\Warranty Debit\Warranty Debit.xlsx"
    try:
        df = pd.read_excel(input_path, sheet_name="Sheet1")

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

        for col in ["Total Claim Amount", "Credit Note Amount", "Debit Note Amount"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = 0

        df["Dealer_Code"] = df["Dealer Location"].map(dealer_mapping).fillna(df["Dealer Location"])
        df["Month"] = df["Fiscal Month"].astype(str).str.strip().str[:3]
        df["Claim arbitration ID"] = df["Claim arbitration ID"].astype(str).replace("nan", "").replace("", np.nan)

        dealers = sorted(df["Dealer_Code"].dropna().unique())
        months = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

        # CREDIT
        credit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            month_data = df[df["Month"] == m]
            if not month_data.empty:
                summary = month_data.groupby("Dealer_Code")["Credit Note Amount"].sum().reset_index()
                summary.columns = ["Division", f"Credit Note {m}"]
                credit_df = credit_df.merge(summary, on="Division", how="left")
            else:
                credit_df[f"Credit Note {m}"] = 0
        credit_df = credit_df.fillna(0)
        credit_cols = [f"Credit Note {m}" for m in months]
        credit_df["Total Credit"] = credit_df[credit_cols].sum(axis=1)

        grand_total = {"Division": "Grand Total"}
        for col in credit_df.columns[1:]:
            grand_total[col] = float(credit_df[col].sum())
        credit_df = pd.concat([credit_df, pd.DataFrame([grand_total])], ignore_index=True)

        # DEBIT
        debit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            month_data = df[df["Month"] == m]
            if not month_data.empty:
                summary = month_data.groupby("Dealer_Code")["Debit Note Amount"].sum().reset_index()
                summary.columns = ["Division", f"Debit Note {m}"]
                debit_df = debit_df.merge(summary, on="Division", how="left")
            else:
                debit_df[f"Debit Note {m}"] = 0
        debit_df = debit_df.fillna(0)
        debit_cols = [f"Debit Note {m}" for m in months]
        debit_df["Total Debit"] = debit_df[debit_cols].sum(axis=1)

        grand_total = {"Division": "Grand Total"}
        for col in debit_df.columns[1:]:
            grand_total[col] = float(debit_df[col].sum())
        debit_df = pd.concat([debit_df, pd.DataFrame([grand_total])], ignore_index=True)

        # ARBITRATION
        arbitration_df = pd.DataFrame({"Division": dealers})

        def is_arbitration(value):
            if pd.isna(value):
                return False
            v = str(value).strip().upper()
            return v.startswith("ARB") and v not in ("", "NAN")

        for m in months:
            md = df[df["Month"] == m].copy()
            md["Is_ARB"] = md["Claim arbitration ID"].apply(is_arbitration)
            md["Arbitration_Amount"] = np.where(md["Is_ARB"], md["Debit Note Amount"], 0)
            arb_summary = md.groupby("Dealer_Code")["Arbitration_Amount"].sum().reset_index()
            arb_summary.columns = ["Division", f"Claim Arbitration {m}"]
            arbitration_df = arbitration_df.merge(arb_summary, on="Division", how="left")

        arbitration_df = arbitration_df.fillna(0)
        arbitration_cols = [f"Claim Arbitration {m}" for m in months]

        total_debit_by_dealer = debit_df[debit_df["Division"] != "Grand Total"][["Division", "Total Debit"]].copy()
        arbitration_df = arbitration_df.merge(total_debit_by_dealer, on="Division", how="left")
        arbitration_df["Pending Claim Arbitration"] = arbitration_df["Total Debit"] - arbitration_df[arbitration_cols].sum(axis=1)
        arbitration_df = arbitration_df.drop(columns=["Total Debit"])

        grand_total = {"Division": "Grand Total"}
        for col in arbitration_df.columns[1:]:
            grand_total[col] = float(arbitration_df[col].sum())
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([grand_total])], ignore_index=True)

        return credit_df, debit_df, arbitration_df, df

    except FileNotFoundError:
        return None, None, None, None
    except Exception:
        return None, None, None, None


# ==================== EXCEL EXPORT HELPERS ====================

def _styles():
    header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    return header_fill, header_font, border

def _autosize(ws, df, max_width=40):
    for col_idx, col in enumerate(df.columns, 1):
        try:
            series = df[col].astype(str)
            max_len = max(series.map(len).max(), len(str(col))) + 2
        except Exception:
            max_len = len(str(col)) + 2
        max_len = min(max_len, max_width)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = max_len

def _write_df(ws, df, header_fill, header_font, border, num_format="#,##0.00"):
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=str(col_name))
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, row in enumerate(df.itertuples(index=False), start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c)
            if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
                cell.value = float(val)
                cell.number_format = num_format
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (datetime, pd.Timestamp)):
                cell.value = val
                cell.number_format = "DD-MM-YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = "" if pd.isna(val) else str(val)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border


def _workbook_bytes(wb: Workbook) -> bytes:
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ==================== DASHBOARD HTML (NO LOGIN / NO LOGOUT / NO CHANGE PASSWORD) ====================

DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Unnati Warranty Management Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%);
            min-height: 100vh;
        }

        .navbar {
            background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
            color: white;
            padding: 15px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            position: sticky;
            top: 0;
            z-index: 100;
        }

        .navbar .container-fluid {
            max-width: 1400px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 30px;
        }

        .navbar-brand {
            font-size: 22px;
            font-weight: 700;
        }

        .container {
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 20px;
        }

        .dashboard-content {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 30px;
        }

        .nav-tabs {
            border-bottom: 2px solid #FF8C00;
            margin-bottom: 20px;
        }

        .nav-tabs .nav-link {
            color: #666;
            font-weight: 600;
            border: none;
            border-bottom: 3px solid transparent;
            padding: 12px 20px;
            cursor: pointer;
            transition: all 0.3s ease;
            background: transparent;
        }

        .nav-tabs .nav-link:hover {
            color: #FF8C00;
            border-bottom-color: #FF8C00;
        }

        .nav-tabs .nav-link.active {
            color: #FF8C00;
            border-bottom-color: #FF8C00;
        }

        .tab-content { display: none; }
        .tab-content.active { display: block; }

        .table-title {
            font-size: 16px;
            font-weight: 700;
            color: #FF8C00;
            margin: 10px 0 10px 0;
        }

        .table-wrapper { overflow-x: auto; }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
            font-size: 12px;
        }

        .data-table thead th {
            background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
            color: white;
            padding: 12px;
            text-align: center;
            font-weight: 600;
            border: none;
            font-size: 11px;
        }

        .data-table tbody td {
            padding: 10px 12px;
            border-bottom: 1px solid #e0e0e0;
            text-align: right;
        }

        .data-table tbody td:first-child {
            text-align: left;
            font-weight: 600;
            color: #333;
        }

        .data-table tbody tr:hover { background: #f9f9f9; }

        .data-table tbody tr:last-child {
            background: #fff8f3;
            font-weight: 700;
            border-top: 2px solid #FF8C00;
            border-bottom: 2px solid #FF8C00;
        }

        .data-table tbody tr:last-child td { color: #FF8C00; }

        .loading-spinner { text-align: center; padding: 40px; }
        .spinner {
            border: 4px solid rgba(255,140,0,0.2);
            border-top: 4px solid #FF8C00;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg);} 100% { transform: rotate(360deg);} }

        .export-section {
            margin: 18px 0 22px 0;
            padding: 16px;
            background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%);
            border-radius: 8px;
            border-left: 5px solid #FF8C00;
            box-shadow: 0 2px 8px rgba(255,140,0,0.1);
        }

        .export-section h3 {
            color: #FF8C00;
            margin-bottom: 12px;
            font-size: 16px;
            font-weight: 700;
        }

        .export-controls {
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
            background: white;
            padding: 12px;
            border-radius: 6px;
        }

        .export-control-group {
            display: flex;
            gap: 8px;
            align-items: center;
        }

        .export-control-group label {
            font-weight: 600;
            color: #333;
            font-size: 14px;
            min-width: 80px;
        }

        .export-control-group select {
            padding: 8px 12px;
            border: 2px solid #FF8C00;
            border-radius: 4px;
            cursor: pointer;
            background: white;
            font-size: 13px;
            min-width: 170px;
        }

        .export-btn {
            padding: 10px 25px;
            background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 700;
            font-size: 14px;
        }

        .export-btn:disabled { background: #ccc; cursor: not-allowed; }
    </style>
</head>
<body>
    <nav class="navbar">
        <div class="container-fluid">
            <span class="navbar-brand">Unnati Motors Warranty Management Dashboard</span>
            <span id="statusText" style="font-weight:600;"></span>
        </div>
    </nav>

    <div class="container">
        <div class="dashboard-content">
            <div class="loading-spinner" id="loadingSpinner">
                <div class="spinner"></div>
                <p style="margin-top: 15px; color: #666;">Loading warranty data...</p>
            </div>

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
                            <label for="divisionFilter">Division:</label>
                            <select id="divisionFilter">
                                <option value="">-- Select Division --</option>
                                <option value="All">All Divisions</option>
                            </select>
                        </div>

                        <div class="export-control-group">
                            <label for="exportType">Export Type:</label>
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
                    <div class="table-title">Warranty Credit Note by Division and Month</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="creditTable"><thead></thead><tbody></tbody></table>
                    </div>
                </div>

                <div id="debit" class="tab-content">
                    <div class="table-title">Warranty Debit Note by Division and Month</div>
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

    function renderTable(tableId, data, maxFrac=0) {
        if (!data || data.length === 0) return;
        const table = document.getElementById(tableId);
        const headers = Object.keys(data[0]);

        table.querySelector('thead').innerHTML = headers.map(h => `<th>${h}</th>`).join('');
        table.querySelector('tbody').innerHTML = data.map(row => {
            return '<tr>' + headers.map(h => {
                let v = row[h];
                if (typeof v === 'number') {
                    v = v.toLocaleString('en-IN', {maximumFractionDigits: maxFrac});
                }
                return `<td>${v}</td>`;
            }).join('') + '</tr>';
        }).join('');
    }

    function loadDivisions() {
        const divisions = new Set();
        const type = document.getElementById('exportType').value;

        let dataSource = warrantyData.credit;
        if (type === 'debit') dataSource = warrantyData.debit;
        if (type === 'arbitration') dataSource = warrantyData.arbitration;
        if (type === 'currentmonth') dataSource = warrantyData.currentMonth;
        if (type === 'compensation') dataSource = warrantyData.compensation;
        if (type === 'pr_approval') dataSource = warrantyData.prApproval;

        if (dataSource && dataSource.length > 0) {
            dataSource.forEach(r => {
                if (r.Division && r.Division !== 'Grand Total') divisions.add(r.Division);
            });
        }

        const sel = document.getElementById('divisionFilter');
        const keep = sel.value;

        sel.innerHTML = '<option value="">-- Select Division --</option><option value="All">All Divisions</option>';
        Array.from(divisions).sort().forEach(d => {
            const opt = document.createElement('option');
            opt.value = d;
            opt.textContent = d;
            sel.appendChild(opt);
        });

        if (keep && sel.querySelector(`option[value="${keep}"]`)) sel.value = keep;
    }

    document.getElementById('exportType').addEventListener('change', loadDivisions);

    async function exportToExcel() {
        const division = document.getElementById('divisionFilter').value;
        const type = document.getElementById('exportType').value;
        const btn = document.getElementById('exportBtn');

        if (!division) { alert('Please select a division'); return; }

        btn.disabled = true;
        btn.textContent = 'Exporting...';

        try {
            const resp = await fetch('/api/export-to-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ division, type })
            });

            if (!resp.ok) {
                const err = await resp.json().catch(() => ({detail: 'Export failed'}));
                throw new Error(err.detail || 'Export failed');
            }

            const blob = await resp.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${type}_${division}_${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            alert('Export completed successfully');
        } catch (e) {
            alert('Export failed: ' + e.message);
        } finally {
            btn.disabled = false;
            btn.textContent = 'Export to Excel';
        }
    }

    async function loadDashboard() {
        const spinner = document.getElementById('loadingSpinner');
        const tabs = document.getElementById('warrantyTabs');
        const statusText = document.getElementById('statusText');

        spinner.style.display = 'block';
        tabs.style.display = 'none';
        statusText.textContent = '';

        try {
            const resp = await fetch('/api/warranty-data', { method: 'GET' });
            if (!resp.ok) throw new Error('Failed to load data');
            warrantyData = await resp.json();

            renderTable('creditTable', warrantyData.credit, 0);
            renderTable('debitTable', warrantyData.debit, 0);
            renderTable('arbitrationTable', warrantyData.arbitration, 0);
            renderTable('currentMonthTable', warrantyData.currentMonth, 0);
            renderTable('compensationTable', warrantyData.compensation, 2);
            renderTable('prApprovalTable', warrantyData.prApproval, 2);

            loadDivisions();

            spinner.style.display = 'none';
            tabs.style.display = 'block';
            statusText.textContent = 'Ready';
        } catch (e) {
            spinner.innerHTML = '<p style="color:red; padding:20px; text-align:center;">Error loading warranty data<br><br><button onclick="location.reload();" style="padding:10px 20px; background:#FF8C00; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:600;">Refresh</button></p>';
        }
    }

    window.onload = loadDashboard;
</script>
</body>
</html>
"""

# ==================== FASTAPI SETUP ====================

app = FastAPI()

@app.get("/")
async def root():
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/api/warranty-data")
async def get_warranty_data():
    if WARRANTY_DATA["credit_df"] is None:
        return JSONResponse(
            content={
                "credit": [],
                "debit": [],
                "arbitration": [],
                "currentMonth": [],
                "compensation": [],
                "prApproval": [],
            },
            status_code=200,
        )

    def df_records(key):
        df = WARRANTY_DATA.get(key)
        if df is None or df.empty:
            return []
        recs = df.to_dict("records")
        for r in recs:
            for k in list(r.keys()):
                if pd.isna(r[k]):
                    r[k] = 0
        return recs

    return {
        "credit": df_records("credit_df"),
        "debit": df_records("debit_df"),
        "arbitration": df_records("arbitration_df"),
        "currentMonth": df_records("current_month_df"),
        "compensation": df_records("compensation_df"),
        "prApproval": df_records("pr_approval_df"),
    }

@app.post("/api/export-to-excel")
async def export_to_excel(request: Request):
    body = await request.json()
    selected_division = body.get("division", "All")
    export_type = body.get("type", "credit")

    if export_type not in ["credit", "debit", "arbitration", "currentmonth", "compensation", "pr_approval"]:
        raise HTTPException(status_code=400, detail="Invalid export type")

    if export_type == "currentmonth":
        return await export_current_month_warranty(selected_division)
    if export_type == "compensation":
        return await export_compensation_claim(selected_division)
    if export_type == "pr_approval":
        return await export_pr_approval(selected_division)

    df = WARRANTY_DATA["credit_df"] if export_type == "credit" else WARRANTY_DATA["debit_df"] if export_type == "debit" else WARRANTY_DATA["arbitration_df"]
    if df is None or df.empty:
        raise HTTPException(status_code=500, detail="No data available for export")

    # filter summary
    if selected_division not in ("All", "Grand Total"):
        df_export = df[df["Division"] == selected_division].copy()
        gt = df[df["Division"] == "Grand Total"]
        if not gt.empty:
            df_export = pd.concat([df_export, gt], ignore_index=True)
    else:
        df_export = df.copy()

    wb = Workbook()
    header_fill, header_font, border = _styles()

    ws1 = wb.active
    ws1.title = f"{export_type.capitalize()}" if selected_division in ("All", "Grand Total") else f"{selected_division} - {export_type.capitalize()}"
    _write_df(ws1, df_export, header_fill, header_font, border, num_format="#,##0.00")
    _autosize(ws1, df_export, max_width=35)

    # detailed sheets only for single division (like your local logic)
    if selected_division not in ("All", "Grand Total"):
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
        source_df = WARRANTY_DATA.get("source_df")

        if dealer_location and source_df is not None and not source_df.empty:
            detail_df = source_df[source_df["Dealer Location"] == dealer_location].copy()

            def is_empty_or_hyphen(value):
                if pd.isna(value):
                    return True
                v = str(value).strip()
                return v in ("", "-", "NAN", "nan")

            def has_valid_arb_id(value):
                if pd.isna(value):
                    return False
                v = str(value).strip().upper()
                return v.startswith("ARB") and v not in ("", "NAN")

            required_columns = [
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
                if "Credit Note Amount" in detail_df.columns:
                    detail_df = detail_df[detail_df["Credit Note Amount"] > 0].copy()
                detail_df = detail_df[detail_df["Claim arbitration ID"].apply(is_empty_or_hyphen)].copy()
                required_columns.append("Credit Note Amount")

            elif export_type == "debit":
                if "Debit Note Amount" in detail_df.columns:
                    detail_df = detail_df[detail_df["Debit Note Amount"] > 0].copy()
                required_columns.append("Debit Note Amount")
                if "Total Claim Amount" in detail_df.columns:
                    required_columns.append("Total Claim Amount")

            else:  # arbitration
                detail_df = detail_df[detail_df["Claim arbitration ID"].apply(has_valid_arb_id)].copy()
                required_columns.append("Debit Note Amount")

            available_cols = [c for c in required_columns if c in detail_df.columns]
            detail_df = detail_df[available_cols].copy()

            if "Claim No" in detail_df.columns:
                def format_claim_no(x):
                    if pd.isna(x) or str(x).strip() == "":
                        return ""
                    try:
                        return str(int(float(x)))
                    except Exception:
                        return str(x).strip()
                detail_df["Claim No"] = detail_df["Claim No"].apply(format_claim_no)

            if "Ro Id" in detail_df.columns:
                def format_ro_id(x):
                    if pd.isna(x) or str(x).strip() == "":
                        return ""
                    try:
                        return f"RO{str(int(float(x)))}"
                    except Exception:
                        s = str(x).strip()
                        return s if s.startswith("RO") else f"RO{s}"
                detail_df["Ro Id"] = detail_df["Ro Id"].apply(format_ro_id)

            month_order = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
            if "Fiscal Month" in detail_df.columns:
                detail_df["Month"] = detail_df["Fiscal Month"].astype(str).str.strip().str[:3]
                detail_df["Month_Order"] = detail_df["Month"].apply(lambda x: month_order.index(x) if x in month_order else 999)
                detail_df = detail_df.sort_values("Month_Order").drop(columns=["Month", "Month_Order"])

            ws2 = wb.create_sheet(title=f"{selected_division} - Detailed Data")
            _write_df(ws2, detail_df, header_fill, header_font, border, num_format="#,##0.00")
            _autosize(ws2, detail_df, max_width=40)

            # Pending arbitration sheet only for arbitration export
            if export_type == "arbitration":
                pending_df = source_df[source_df["Dealer Location"] == dealer_location].copy()
                if "Debit Note Amount" in pending_df.columns:
                    pending_df = pending_df[pending_df["Debit Note Amount"] > 0].copy()
                pending_df = pending_df[pending_df["Claim arbitration ID"].apply(is_empty_or_hyphen)].copy()

                pending_cols = [
                    "Fiscal Month", "Dealer Location", "Claim arbitration ID", "Claim Invoice Date",
                    "Claim No", "Claim Date", "Chassis No", "Ro Id", "Claim Type",
                    "Credit Note Amount", "Debit Note Amount"
                ]
                pending_cols = [c for c in pending_cols if c in pending_df.columns]
                pending_df = pending_df[pending_cols].copy()

                if "Claim No" in pending_df.columns:
                    pending_df["Claim No"] = pending_df["Claim No"].apply(lambda x: "" if pd.isna(x) else str(int(float(x))) if str(x).strip() != "" and str(x).replace(".0","").isdigit() else str(x).strip())
                if "Ro Id" in pending_df.columns:
                    def _ro(x):
                        if pd.isna(x) or str(x).strip() == "":
                            return ""
                        try:
                            return f"RO{str(int(float(x)))}"
                        except Exception:
                            s = str(x).strip()
                            return s if s.startswith("RO") else f"RO{s}"
                    pending_df["Ro Id"] = pending_df["Ro Id"].apply(_ro)

                if "Debit Note Amount" in pending_df.columns:
                    pending_df = pending_df.rename(columns={"Debit Note Amount": "Pending Arbitration Amount"})

                if "Fiscal Month" in pending_df.columns:
                    pending_df["Month"] = pending_df["Fiscal Month"].astype(str).str.strip().str[:3]
                    pending_df["Month_Order"] = pending_df["Month"].apply(lambda x: month_order.index(x) if x in month_order else 999)
                    pending_df = pending_df.sort_values("Month_Order").drop(columns=["Month", "Month_Order"])

                ws3 = wb.create_sheet(title=f"{selected_division} - Pending Arb")
                _write_df(ws3, pending_df, header_fill, header_font, border, num_format="#,##0.00")
                _autosize(ws3, pending_df, max_width=40)

    file_bytes = _workbook_bytes(wb)
    filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # IMPORTANT: return BytesIO stream (more reliable in browsers)
    out = io.BytesIO(file_bytes)
    out.seek(0)

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

async def export_current_month_warranty(selected_division: str):
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
    header_fill, header_font, border = _styles()

    ws1 = wb.active
    ws1.title = "Current Month Summary" if selected_division in ("All", "Grand Total") else f"{selected_division} - Summary"
    _write_df(ws1, df_export, header_fill, header_font, border, num_format="#,##0")
    _autosize(ws1, df_export, max_width=35)

    if source_df is not None and not source_df.empty and "Division" in source_df.columns:
        # Spares
        spares_df = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            spares_df = spares_df[spares_df["Division"] == selected_division].copy()
        if "Pending Claims Spares" in spares_df.columns:
            spares_df = spares_df[spares_df["Pending Claims Spares"].notna()].copy()

        if not spares_df.empty:
            ws2 = wb.create_sheet(title="Pending Spares Claims" if selected_division in ("All", "Grand Total") else f"{selected_division} - Spares")
            _write_df(ws2, spares_df, header_fill, header_font, border, num_format="#,##0.00")
            _autosize(ws2, spares_df, max_width=45)

        # Labour
        labour_df = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            labour_df = labour_df[labour_df["Division"] == selected_division].copy()
        if "Pending Claims Labour" in labour_df.columns:
            labour_df = labour_df[labour_df["Pending Claims Labour"].notna()].copy()

        if not labour_df.empty:
            ws3 = wb.create_sheet(title="Pending Labour Claims" if selected_division in ("All", "Grand Total") else f"{selected_division} - Labour")
            _write_df(ws3, labour_df, header_fill, header_font, border, num_format="#,##0.00")
            _autosize(ws3, labour_df, max_width=45)

    file_bytes = _workbook_bytes(wb)
    filename = f"{selected_division}_CurrentMonthWarranty_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out = io.BytesIO(file_bytes)
    out.seek(0)

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

async def export_compensation_claim(selected_division: str):
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
    header_fill, header_font, border = _styles()

    ws1 = wb.active
    ws1.title = "Compensation Summary" if selected_division in ("All", "Grand Total") else f"{selected_division} - Summary"
    _write_df(ws1, df_export, header_fill, header_font, border, num_format="#,##0.00")
    _autosize(ws1, df_export, max_width=35)

    if source_df is not None and not source_df.empty and "Division" in source_df.columns:
        detail_df = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            detail_df = detail_df[detail_df["Division"] == selected_division].copy()

        if not detail_df.empty:
            ws2 = wb.create_sheet(title="Compensation Details" if selected_division in ("All", "Grand Total") else f"{selected_division} - Details")
            _write_df(ws2, detail_df, header_fill, header_font, border, num_format="#,##0.00")
            _autosize(ws2, detail_df, max_width=45)

    file_bytes = _workbook_bytes(wb)
    filename = f"{selected_division}_CompensationClaim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out = io.BytesIO(file_bytes)
    out.seek(0)

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

async def export_pr_approval(selected_division: str):
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
    header_fill, header_font, border = _styles()

    ws1 = wb.active
    ws1.title = "PR Approval Summary" if selected_division in ("All", "Grand Total") else f"{selected_division} - Summary"
    _write_df(ws1, df_export, header_fill, header_font, border, num_format="#,##0.00")
    _autosize(ws1, df_export, max_width=35)

    if source_df is not None and not source_df.empty and "Division" in source_df.columns:
        detail_df = source_df.copy()
        if selected_division not in ("All", "Grand Total"):
            detail_df = detail_df[detail_df["Division"] == selected_division].copy()

        if not detail_df.empty:
            ws2 = wb.create_sheet(title="PR Approval Details" if selected_division in ("All", "Grand Total") else f"{selected_division} - Details")
            _write_df(ws2, detail_df, header_fill, header_font, border, num_format="#,##0.00")
            _autosize(ws2, detail_df, max_width=45)

    file_bytes = _workbook_bytes(wb)
    filename = f"{selected_division}_PrApproval_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out = io.BytesIO(file_bytes)
    out.seek(0)

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ==================== STARTUP LOAD ====================

print("\n" + "=" * 80)
print("STARTING WARRANTY DASHBOARD (NO LOGIN) - PORT 8001")
print("=" * 80)

print("Loading warranty data...")
WARRANTY_DATA["credit_df"], WARRANTY_DATA["debit_df"], WARRANTY_DATA["arbitration_df"], WARRANTY_DATA["source_df"] = process_warranty_data()

print("Loading current month warranty...")
WARRANTY_DATA["current_month_df"], WARRANTY_DATA["current_month_source_df"] = process_current_month_warranty()

print("Loading compensation claim...")
WARRANTY_DATA["compensation_df"], WARRANTY_DATA["compensation_source_df"] = process_compensation_claim()

print("Loading PR approval...")
WARRANTY_DATA["pr_approval_df"], WARRANTY_DATA["pr_approval_source_df"] = process_pr_approval()

if __name__ == "__main__":
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = "127.0.0.1"

    port = 8001
    print("\n" + "=" * 80)
    print("SERVER READY")
    print("=" * 80)
    print(f"PORT: {port}")
    print(f"Local URL:   http://localhost:{port}/")
    print(f"Network URL: http://{local_ip}:{port}/")
    print("=" * 80 + "\n")

    uvicorn.run(app, host="0.0.0.0", port=port)
