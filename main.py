import os
import io
import socket
import secrets
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd
import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ==================== GLOBAL STORAGE ====================

WARRANTY_DATA = {
    "credit_df": None,
    "debit_df": None,
    "arbitration_df": None,  # Pending Arbitration month-wise summary
    "source_df": None,

    "current_month_df": None,
    "current_month_source_df": None,

    "compensation_df": None,
    "compensation_source_df": None,

    "pr_approval_df": None,
    "pr_approval_source_df": None,
}

# ==================== FILE HELPERS ====================

def find_data_file(filename: str):
    possible_paths = [
        f"/mnt/data/{filename}",
        filename,
        f"./{filename}",
        f"Data/{filename}",
        f"data/{filename}",
    ]
    for path in possible_paths:
        if os.path.exists(path):
            print(f"  Found: {filename} at {path}")
            return path
    print(f"  WARNING: {filename} not found. Checked: {possible_paths}")
    return None

# ==================== PR APPROVAL ====================

def process_pr_approval():
    input_path = find_data_file("Pr_Approval_Claims_Merged.xlsx")
    if input_path is None:
        print("  PR Approval file not found - returning empty data")
        return None, None

    try:
        df = pd.read_excel(input_path)
        print("  PR Approval data loaded successfully")
        print(f"  Total rows in source data: {len(df)}")

        summary_columns = [
            "Division",
            "PA Request No.",
            "PA Date",
            "Request Type",
            "App. Claim Amt from M&M",
        ]
        available_summary_columns = [c for c in summary_columns if c in df.columns]
        if not available_summary_columns:
            print("  No required columns found in PR Approval file")
            return None, df

        df_display = df[available_summary_columns].copy()

        if "Division" in df_display.columns:
            df_display["Division"] = df_display["Division"].astype(str).str.strip()
            df_display = df_display[
                df_display["Division"].notna()
                & (df_display["Division"] != "")
                & (df_display["Division"] != "nan")
            ]

        if "App. Claim Amt from M&M" in df_display.columns:
            df_display["App. Claim Amt from M&M"] = pd.to_numeric(
                df_display["App. Claim Amt from M&M"], errors="coerce"
            ).fillna(0)

        summary_data = []
        if "Division" in df_display.columns and not df_display.empty:
            for division in sorted(df_display["Division"].unique()):
                div_data = df_display[df_display["Division"] == division]
                row = {"Division": division}
                row["Total Requests"] = len(div_data)
                if "App. Claim Amt from M&M" in df_display.columns:
                    row["Total Approved Amount"] = float(div_data["App. Claim Amt from M&M"].sum())
                if "Request Type" in df_display.columns:
                    request_types = div_data["Request Type"].value_counts().to_dict()
                    for k, v in request_types.items():
                        if pd.notna(k) and str(k).strip() != "":
                            row[f"{k} Count"] = int(v)
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

    except Exception as e:
        import traceback
        print(f"  Error processing PR Approval data: {e}")
        traceback.print_exc()
        return None, None

# ==================== COMPENSATION CLAIM ====================

def process_compensation_claim():
    input_path = find_data_file("Transit_Claims_Merged.xlsx")
    if input_path is None:
        print("  Compensation Claim file not found - returning empty data")
        return None, None

    try:
        df = pd.read_excel(input_path)
        print("  Compensation Claim data loaded successfully")
        print(f"  Total rows in source data: {len(df)}")

        required_columns = [
            "Division",
            "RO Id.",
            "Registration No.",
            "RO Date",
            "RO Bill Date",
            "Chassis No.",
            "Model Group",
            "Claim Amount",
            "Request Status",
            "Claim Approved Amt.",
            "No. of Days",
        ]
        available_columns = [c for c in required_columns if c in df.columns]
        if not available_columns:
            print("  No required columns found in Compensation Claim file")
            return None, df

        df_filtered = df[available_columns].copy()

        if "Division" in df_filtered.columns:
            df_filtered["Division"] = df_filtered["Division"].astype(str).str.strip()
            df_filtered = df_filtered[
                df_filtered["Division"].notna()
                & (df_filtered["Division"] != "")
                & (df_filtered["Division"] != "nan")
            ]

        if "RO Id." in df_filtered.columns:
            def format_ro_id(x):
                if pd.isna(x) or str(x).strip() == "":
                    return ""
                try:
                    return f"RO{str(int(float(x)))}"
                except Exception:
                    s = str(x).strip()
                    return s if s.upper().startswith("RO") else f"RO{s}"
            df_filtered["RO Id."] = df_filtered["RO Id."].apply(format_ro_id)

        for col in ["Claim Amount", "Claim Approved Amt.", "No. of Days"]:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce").fillna(0)

        summary_data = []
        if "Division" in df_filtered.columns and not df_filtered.empty:
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

    except Exception as e:
        import traceback
        print(f"  Error processing compensation claim data: {e}")
        traceback.print_exc()
        return None, None

# ==================== CURRENT MONTH WARRANTY ====================

def process_current_month_warranty():
    input_path = find_data_file("Pending Warranty Claim Details.xlsx")
    if input_path is None:
        print("  Current Month Warranty file not found - returning empty data")
        return None, None

    try:
        df = pd.read_excel(input_path, sheet_name="Pending Warranty Claim Details")
        print("  Current Month Warranty data loaded successfully")
        print(f"  Total rows in source data: {len(df)}")

        required_columns = ["Division", "Pending Claims Spares", "Pending Claims Labour"]
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            print(f"  Missing columns in Current Month Warranty: {missing}")
            return None, df

        df["Division"] = df["Division"].astype(str).str.strip()
        df = df[df["Division"].notna() & (df["Division"] != "") & (df["Division"] != "nan")]

        summary_data = []
        for division in sorted(df["Division"].unique()):
            div_data = df[df["Division"] == division]
            spares_count = int(div_data["Pending Claims Spares"].notna().sum())
            labour_count = int(div_data["Pending Claims Labour"].notna().sum())
            summary_data.append(
                {
                    "Division": division,
                    "Pending Claims Spares Count": spares_count,
                    "Pending Claims Labour Count": labour_count,
                    "Total Pending Claims": spares_count + labour_count,
                }
            )

        summary_df = pd.DataFrame(summary_data)
        if not summary_df.empty:
            grand_total = {
                "Division": "Grand Total",
                "Pending Claims Spares Count": int(summary_df["Pending Claims Spares Count"].sum()),
                "Pending Claims Labour Count": int(summary_df["Pending Claims Labour Count"].sum()),
                "Total Pending Claims": int(summary_df["Total Pending Claims"].sum()),
            }
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

        return summary_df, df

    except Exception as e:
        import traceback
        print(f"  Error processing current month warranty data: {e}")
        traceback.print_exc()
        return None, None

# ==================== WARRANTY CREDIT/DEBIT + PENDING ARB MONTHWISE ====================

def process_warranty_data():
    input_path = find_data_file("Warranty Debit.xlsx")
    if input_path is None:
        print("  Warranty Debit file not found - returning empty data")
        return None, None, None, None

    try:
        df = pd.read_excel(input_path, sheet_name="Sheet1")
        print("  Warranty data loaded successfully")
        print(f"  Total rows in source data: {len(df)}")

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

        numeric_columns = ["Total Claim Amount", "Credit Note Amount", "Debit Note Amount"]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = 0

        if "Dealer Location" not in df.columns:
            raise ValueError("Column 'Dealer Location' missing in Warranty Debit.xlsx Sheet1")
        if "Fiscal Month" not in df.columns:
            raise ValueError("Column 'Fiscal Month' missing in Warranty Debit.xlsx Sheet1")

        df["Dealer_Code"] = df["Dealer Location"].map(dealer_mapping).fillna(df["Dealer Location"].astype(str))
        df["Month"] = df["Fiscal Month"].astype(str).str.strip().str[:3]

        if "Claim arbitration ID" in df.columns:
            df["Claim arbitration ID"] = df["Claim arbitration ID"].astype(str).replace("nan", "").replace("", np.nan)
        else:
            df["Claim arbitration ID"] = np.nan

        dealers = sorted(df["Dealer_Code"].dropna().unique().tolist())
        months = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

        # CREDIT
        credit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m]
            summary = md.groupby("Dealer_Code")["Credit Note Amount"].sum().reset_index()
            summary.columns = ["Division", f"Credit Note {m}"]
            credit_df = credit_df.merge(summary, on="Division", how="left")
        credit_df = credit_df.fillna(0)
        credit_cols = [f"Credit Note {m}" for m in months]
        credit_df["Total Credit"] = credit_df[credit_cols].sum(axis=1)

        gt = {"Division": "Grand Total"}
        for col in credit_df.columns[1:]:
            gt[col] = float(credit_df[col].sum())
        credit_df = pd.concat([credit_df, pd.DataFrame([gt])], ignore_index=True)

        # DEBIT
        debit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m]
            summary = md.groupby("Dealer_Code")["Debit Note Amount"].sum().reset_index()
            summary.columns = ["Division", f"Debit Note {m}"]
            debit_df = debit_df.merge(summary, on="Division", how="left")
        debit_df = debit_df.fillna(0)
        debit_cols = [f"Debit Note {m}" for m in months]
        debit_df["Total Debit"] = debit_df[debit_cols].sum(axis=1)

        gt = {"Division": "Grand Total"}
        for col in debit_df.columns[1:]:
            gt[col] = float(debit_df[col].sum())
        debit_df = pd.concat([debit_df, pd.DataFrame([gt])], ignore_index=True)

        # PENDING ARBITRATION (month-wise) = Debit Note Amount where Claim arbitration ID is blank/-/nan
        def is_empty_or_hyphen(v):
            if pd.isna(v):
                return True
            s = str(v).strip()
            return s == "" or s == "-" or s.upper() == "NAN"

        arbitration_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m].copy()
            md["Is_Pending_ARB"] = md["Claim arbitration ID"].apply(is_empty_or_hyphen)
            md["Pending_Arb_Amount"] = np.where(md["Is_Pending_ARB"], md["Debit Note Amount"], 0)
            summary = md.groupby("Dealer_Code")["Pending_Arb_Amount"].sum().reset_index()
            summary.columns = ["Division", f"Pending Arbitration {m}"]
            arbitration_df = arbitration_df.merge(summary, on="Division", how="left")

        arbitration_df = arbitration_df.fillna(0)
        pend_cols = [f"Pending Arbitration {m}" for m in months]
        arbitration_df["Total Pending Arbitration"] = arbitration_df[pend_cols].sum(axis=1)

        gt = {"Division": "Grand Total"}
        for col in arbitration_df.columns[1:]:
            gt[col] = float(arbitration_df[col].sum())
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([gt])], ignore_index=True)

        return credit_df, debit_df, arbitration_df, df

    except Exception as e:
        import traceback
        print(f"  Error processing warranty data: {e}")
        traceback.print_exc()
        return None, None, None, None

# ==================== DASHBOARD HTML ====================

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Unnati Warranty Management Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body{
            font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%);
            min-height:100vh;
        }
        .navbar{
            background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
            color:white; padding:15px 0;
            box-shadow:0 2px 8px rgba(0,0,0,0.15);
            position:sticky; top:0; z-index:100;
        }
        .navbar .container-fluid{
            max-width:1400px; margin:0 auto;
            display:flex; justify-content:center; align-items:center;
            padding:0 30px;
        }
        .navbar-brand{ font-size:24px; font-weight:700; }
        .container{ max-width:1400px; margin:30px auto; padding:0 20px; }
        .dashboard-content{
            background:white; border-radius:12px;
            box-shadow:0 2px 10px rgba(0,0,0,0.1);
            padding:30px;
        }
        .nav-tabs{ border-bottom:2px solid #FF8C00; margin-bottom:30px; }
        .nav-tabs .nav-link{
            color:#666; font-weight:600; border:none;
            border-bottom:3px solid transparent;
            padding:12px 20px; cursor:pointer;
            background:transparent;
        }
        .nav-tabs .nav-link:hover{ color:#FF8C00; border-bottom-color:#FF8C00; }
        .nav-tabs .nav-link.active{ color:#FF8C00; border-bottom-color:#FF8C00; background:transparent; }
        .tab-content{ display:none; }
        .tab-content.active{ display:block; }
        .data-table{
            width:100%; border-collapse:collapse;
            margin-top:20px; font-size:12px;
        }
        .data-table thead th{
            background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
            color:white; padding:12px; text-align:center;
            font-weight:600; border:none; font-size:11px;
        }
        .data-table tbody td{
            padding:10px 12px; border-bottom:1px solid #e0e0e0;
            text-align:right;
        }
        .data-table tbody td:first-child{
            text-align:left; font-weight:600; color:#333;
        }
        .data-table tbody tr:hover{ background:#f9f9f9; }
        .data-table tbody tr:last-child{
            background:#fff8f3; font-weight:700;
            border-top:2px solid #FF8C00; border-bottom:2px solid #FF8C00;
        }
        .data-table tbody tr:last-child td{ color:#FF8C00; }
        .loading-spinner{ display:none; text-align:center; padding:40px; }
        .spinner{
            border:4px solid rgba(255,140,0,0.2);
            border-top:4px solid #FF8C00;
            border-radius:50%; width:40px; height:40px;
            animation:spin 1s linear infinite; margin:0 auto;
        }
        @keyframes spin{ 0%{transform:rotate(0deg);} 100%{transform:rotate(360deg);} }
        .table-title{ font-size:16px; font-weight:700; color:#FF8C00; margin-bottom:15px; }
        .table-wrapper{ overflow-x:auto; }

        .export-section{
            margin:30px 0; padding:20px;
            background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%);
            border-radius:8px; border-left:5px solid #FF8C00;
            box-shadow:0 2px 8px rgba(255,140,0,0.1);
        }
        .export-section h3{ color:#FF8C00; margin-bottom:15px; font-size:16px; font-weight:700; }
        .export-controls{
            display:flex; gap:15px; align-items:center; flex-wrap:wrap;
            background:white; padding:15px; border-radius:6px;
        }
        .export-control-group{ display:flex; gap:8px; align-items:center; }
        .export-control-group label{ font-weight:600; color:#333; font-size:14px; min-width:80px; }
        .export-control-group select{
            padding:8px 12px; border:2px solid #FF8C00; border-radius:4px;
            cursor:pointer; background:white; font-size:13px; min-width:150px;
        }
        .export-btn{
            padding:10px 25px;
            background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            color:white; border:none; border-radius:4px;
            cursor:pointer; font-weight:700; font-size:14px;
        }
        .export-btn:disabled{ background:#ccc; cursor:not-allowed; }
    </style>
</head>
<body>
    <nav class="navbar navbar-dark">
        <div class="container-fluid">
            <span class="navbar-brand">Unnati Motors Warranty Management Dashboard</span>
        </div>
    </nav>

    <div class="container">
        <div class="dashboard-content">
            <div class="loading-spinner" id="loadingSpinner">
                <div class="spinner"></div>
                <p style="margin-top:15px; color:#666;">Loading warranty data...</p>
            </div>

            <div id="warrantyTabs" style="display:none;">
                <div class="nav-tabs">
                    <button class="nav-link active" onclick="switchTab('credit', this)">Warranty Credit</button>
                    <button class="nav-link" onclick="switchTab('debit', this)">Warranty Debit</button>
                    <button class="nav-link" onclick="switchTab('arbitration', this)">Claim Arbitration</button>
                    <button class="nav-link" onclick="switchTab('currentmonth', this)">Current Month Warranty</button>
                    <button class="nav-link" onclick="switchTab('compensation', this)">Compensation Claim</button>
                    <button class="nav-link" onclick="switchTab('pr_approval', this)">PR Approval</button>
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
                    <div class="table-title">Pending Claim Arbitration by Division (Month-wise)</div>
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

    async function loadDashboard(){
        const spinner = document.getElementById('loadingSpinner');
        const tabs = document.getElementById('warrantyTabs');

        spinner.style.display = 'block';
        tabs.style.display = 'none';

        try{
            const response = await fetch('/api/warranty-data');
            if(!response.ok){
                throw new Error('Failed to load warranty data: HTTP ' + response.status);
            }

            warrantyData = await response.json();

            displayTable('creditTable', warrantyData.credit, 0);
            displayTable('debitTable', warrantyData.debit, 0);
            displayTable('arbitrationTable', warrantyData.arbitration, 0);
            displayTable('currentMonthTable', warrantyData.currentMonth, 0);
            displayTable('compensationTable', warrantyData.compensation, 2);
            displayTable('prApprovalTable', warrantyData.prApproval, 2);

            loadDivisions();

            spinner.style.display = 'none';
            tabs.style.display = 'block';
        }catch(e){
            console.error(e);
            spinner.innerHTML = '<p style="color:red; padding:20px; text-align:center;">Error loading warranty data<br><br><button onclick="location.reload();" style="padding:10px 20px; background:#FF8C00; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:600;">Refresh</button></p>';
        }
    }

    function displayTable(tableId, data, decimals){
        const table = document.getElementById(tableId);
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');

        if(!data || data.length === 0){
            thead.innerHTML = '';
            tbody.innerHTML = '<tr><td style="text-align:left;" colspan="1">No data</td></tr>';
            return;
        }

        const headers = Object.keys(data[0]);
        thead.innerHTML = '<tr>' + headers.map(h => '<th>' + h + '</th>').join('') + '</tr>';

        tbody.innerHTML = data.map(row => {
            return '<tr>' + headers.map(h => {
                const v = row[h];
                if(typeof v === 'number'){
                    return '<td>' + v.toLocaleString('en-IN', {maximumFractionDigits: decimals}) + '</td>';
                }
                return '<td>' + (v ?? '') + '</td>';
            }).join('') + '</tr>';
        }).join('');
    }

    function switchTab(tabName, btn){
        document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.nav-link').forEach(b => b.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');
        btn.classList.add('active');
    }

    function loadDivisions(){
        const divisions = new Set();
        const type = document.getElementById('exportType').value;

        let dataSource = warrantyData.credit;
        if(type === 'debit') dataSource = warrantyData.debit;
        if(type === 'arbitration') dataSource = warrantyData.arbitration;
        if(type === 'currentmonth') dataSource = warrantyData.currentMonth;
        if(type === 'compensation') dataSource = warrantyData.compensation;
        if(type === 'pr_approval') dataSource = warrantyData.prApproval;

        if(dataSource && dataSource.length > 0){
            dataSource.forEach(r => {
                if(r.Division && r.Division !== 'Grand Total'){
                    divisions.add(r.Division);
                }
            });
        }

        const divisionSelect = document.getElementById('divisionFilter');
        const currentValue = divisionSelect.value;

        divisionSelect.innerHTML = '<option value="">-- Select Division --</option><option value="All">All Divisions</option>';

        Array.from(divisions).sort().forEach(div => {
            const opt = document.createElement('option');
            opt.value = div;
            opt.textContent = div;
            divisionSelect.appendChild(opt);
        });

        if(currentValue && divisionSelect.querySelector('option[value="' + currentValue + '"]')){
            divisionSelect.value = currentValue;
        }
    }

    document.getElementById('exportType').addEventListener('change', loadDivisions);

    async function exportToExcel(){
        const division = document.getElementById('divisionFilter').value;
        const type = document.getElementById('exportType').value;
        const exportBtn = document.getElementById('exportBtn');

        if(!division){
            alert('Please select a division');
            return;
        }

        exportBtn.disabled = true;
        exportBtn.textContent = 'Exporting...';

        try{
            const response = await fetch('/api/export-to-excel', {
                method:'POST',
                headers:{ 'Content-Type':'application/json' },
                body: JSON.stringify({ division: division, type: type })
            });

            if(!response.ok){
                const err = await response.json().catch(() => ({detail:'Export failed'}));
                throw new Error(err.detail || 'Export failed');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = type + '_' + division + '_' + new Date().toISOString().split('T')[0] + '.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            alert('Export completed successfully');
        }catch(e){
            alert('Export failed: ' + e.message);
        }finally{
            exportBtn.disabled = false;
            exportBtn.textContent = 'Export to Excel';
        }
    }

    window.onload = function(){ loadDashboard(); };
</script>

</body>
</html>
"""

# ==================== EXCEL EXPORT HELPERS ====================

def _excel_styles():
    header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    return header_fill, header_font, border

def _autosize_columns(ws, df, max_w=35, min_w=10):
    for col_idx, col in enumerate(df.columns, 1):
        try:
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
        except Exception:
            max_len = len(str(col)) + 2
        ws.column_dimensions[get_column_letter(col_idx)].width = max(min(max_len, max_w), min_w)

def _write_df(ws, df):
    header_fill, header_font, border = _excel_styles()

    for col_idx, col in enumerate(df.columns, 1):
        c = ws.cell(row=1, column=col_idx, value=col)
        c.fill = header_fill
        c.font = header_font
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
                cell.value = float(val)
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (datetime, pd.Timestamp)):
                cell.value = val
                cell.number_format = "dd/mm/yyyy"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = "" if pd.isna(val) else str(val)
                cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

    _autosize_columns(ws, df)

# ==================== FASTAPI APP ====================

app = FastAPI()

@app.get("/")
async def root():
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/dashboard")
async def dashboard():
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/api/warranty-data")
async def get_warranty_data():
    try:
        if WARRANTY_DATA["credit_df"] is None:
            return {
                "credit": [],
                "debit": [],
                "arbitration": [],
                "currentMonth": [],
                "compensation": [],
                "prApproval": [],
            }

        def clean_records(df):
            if df is None:
                return []
            recs = df.to_dict("records")
            for r in recs:
                for k in list(r.keys()):
                    if pd.isna(r[k]):
                        r[k] = 0
            return recs

        return {
            "credit": clean_records(WARRANTY_DATA["credit_df"]),
            "debit": clean_records(WARRANTY_DATA["debit_df"]),
            "arbitration": clean_records(WARRANTY_DATA["arbitration_df"]),
            "currentMonth": clean_records(WARRANTY_DATA["current_month_df"]),
            "compensation": clean_records(WARRANTY_DATA["compensation_df"]),
            "prApproval": clean_records(WARRANTY_DATA["pr_approval_df"]),
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/export-to-excel")
async def export_to_excel(request: Request):
    try:
        body = await request.json()
        selected_division = body.get("division", "All")
        export_type = body.get("type", "credit")

        if export_type not in ["credit", "debit", "arbitration", "currentmonth", "compensation", "pr_approval"]:
            raise HTTPException(status_code=400, detail="Invalid export type")

        if export_type == "currentmonth":
            summary_df = WARRANTY_DATA["current_month_df"]
            source_df = WARRANTY_DATA["current_month_source_df"]
            if summary_df is None or summary_df.empty:
                raise HTTPException(status_code=500, detail="No current month warranty data available")

            if selected_division not in ["All", "Grand Total"]:
                df_export = summary_df[summary_df["Division"] == selected_division].copy()
                gt = summary_df[summary_df["Division"] == "Grand Total"]
                if not gt.empty:
                    df_export = pd.concat([df_export, gt], ignore_index=True)
            else:
                df_export = summary_df.copy()

            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Summary"
            _write_df(ws1, df_export)

            # Optional detail sheets (spares/labour)
            if source_df is not None and not source_df.empty:
                spares_df = source_df.copy()
                labour_df = source_df.copy()
                if selected_division not in ["All", "Grand Total"]:
                    spares_df = spares_df[spares_df["Division"] == selected_division].copy()
                    labour_df = labour_df[labour_df["Division"] == selected_division].copy()

                spares_df = spares_df[spares_df.get("Pending Claims Spares").notna()].copy() if "Pending Claims Spares" in spares_df.columns else pd.DataFrame()
                labour_df = labour_df[labour_df.get("Pending Claims Labour").notna()].copy() if "Pending Claims Labour" in labour_df.columns else pd.DataFrame()

                if not spares_df.empty:
                    ws2 = wb.create_sheet("Spares")
                    _write_df(ws2, spares_df)

                if not labour_df.empty:
                    ws3 = wb.create_sheet("Labour")
                    _write_df(ws3, labour_df)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            filename = f"{selected_division}_CurrentMonthWarranty_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            return StreamingResponse(
                iter([out.getvalue()]),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename={filename}"},
            )

        if export_type == "compensation":
            summary_df = WARRANTY_DATA["compensation_df"]
            source_df = WARRANTY_DATA["compensation_source_df"]
            if summary_df is None or summary_df.empty:
                raise HTTPException(status_code=500, detail="No compensation claim data available")

            if selected_division not in ["All", "Grand Total"]:
                df_export = summary_df[summary_df["Division"] == selected_division].copy()
                gt = summary_df[summary_df["Division"] == "Grand Total"]
                if not gt.empty:
                    df_export = pd.concat([df_export, gt], ignore_index=True)
            else:
                df_export = summary_df.copy()

            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Summary"
            _write_df(ws1, df_export)

            if source_df is not None and not source_df.empty:
                detail_df = source_df.copy()
                if selected_division not in ["All", "Grand Total"]:
                    detail_df = detail_df[detail_df["Division"] == selected_division].copy()
                if not detail_df.empty:
                    ws2 = wb.create_sheet("Details")
                    _write_df(ws2, detail_df)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            filename = f"{selected_division}_CompensationClaim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            return StreamingResponse(
                iter([out.getvalue()]),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename={filename}"},
            )

        if export_type == "pr_approval":
            summary_df = WARRANTY_DATA["pr_approval_df"]
            source_df = WARRANTY_DATA["pr_approval_source_df"]
            if summary_df is None or summary_df.empty:
                raise HTTPException(status_code=500, detail="No PR Approval data available")

            if selected_division not in ["All", "Grand Total"]:
                df_export = summary_df[summary_df["Division"] == selected_division].copy()
                gt = summary_df[summary_df["Division"] == "Grand Total"]
                if not gt.empty:
                    df_export = pd.concat([df_export, gt], ignore_index=True)
            else:
                df_export = summary_df.copy()

            wb = Workbook()
            ws1 = wb.active
            ws1.title = "Summary"
            _write_df(ws1, df_export)

            if source_df is not None and not source_df.empty:
                detail_df = source_df.copy()
                if selected_division not in ["All", "Grand Total"] and "Division" in detail_df.columns:
                    detail_df = detail_df[detail_df["Division"].astype(str).str.strip() == selected_division].copy()
                if not detail_df.empty:
                    ws2 = wb.create_sheet("Details")
                    _write_df(ws2, detail_df)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            filename = f"{selected_division}_PrApproval_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            return StreamingResponse(
                iter([out.getvalue()]),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename={filename}"},
            )

        # credit / debit / arbitration summary export
        if export_type == "credit":
            df = WARRANTY_DATA["credit_df"]
        elif export_type == "debit":
            df = WARRANTY_DATA["debit_df"]
        else:
            df = WARRANTY_DATA["arbitration_df"]

        if df is None or df.empty:
            raise HTTPException(status_code=500, detail="No data available for export")

        if selected_division not in ["All", "Grand Total"]:
            df_export = df[df["Division"] == selected_division].copy()
            gt = df[df["Division"] == "Grand Total"]
            if not gt.empty:
                df_export = pd.concat([df_export, gt], ignore_index=True)
        else:
            df_export = df.copy()

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Summary"
        _write_df(ws1, df_export)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return StreamingResponse(
            iter([out.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Export error: {str(e)}")

# ==================== LOAD DATA ON START ====================

print("\n" + "=" * 100)
print("STARTING WARRANTY MANAGEMENT SYSTEM (NO LOGIN)")
print("=" * 100)

print("\nProcessing warranty data...")
WARRANTY_DATA["credit_df"], WARRANTY_DATA["debit_df"], WARRANTY_DATA["arbitration_df"], WARRANTY_DATA["source_df"] = process_warranty_data()

print("\nProcessing current month warranty data...")
WARRANTY_DATA["current_month_df"], WARRANTY_DATA["current_month_source_df"] = process_current_month_warranty()

print("\nProcessing compensation claim data...")
WARRANTY_DATA["compensation_df"], WARRANTY_DATA["compensation_source_df"] = process_compensation_claim()

print("\nProcessing PR Approval data...")
WARRANTY_DATA["pr_approval_df"], WARRANTY_DATA["pr_approval_source_df"] = process_pr_approval()

if __name__ == "__main__":
    try:
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = "127.0.0.1"

    port = int(os.getenv("PORT", "8001"))
    print("\n" + "=" * 100)
    print("SERVER READY - Warranty Dashboard (NO LOGIN)")
    print("=" * 100)
    print(f"Dashboard URL: http://localhost:{port}/")
    print(f"Network URL:  http://{local_ip}:{port}/")
    print("=" * 100 + "\n")

    uvicorn.run(app, host="0.0.0.0", port=port)
