import pandas as pd
import numpy as np
from datetime import datetime
import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import sys
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ==================== ENVIRONMENT SETUP ====================

IS_RENDER = os.getenv('RENDER') == 'true'
DATA_DIR = os.getenv('DATA_DIR', '/mnt/data')

print(f"\n{'='*80}")
print(f"🚀 WARRANTY MANAGEMENT DASHBOARD")
print(f"{'='*80}")
print(f"Environment: {'RENDER PRODUCTION' if IS_RENDER else 'LOCAL'}")
print(f"Data Directory: {DATA_DIR}")
print(f"Python Version: {sys.version.split()[0]}")
print(f"{'='*80}\n")

# Define file paths based on environment
if IS_RENDER:
    WARRANTY_FILE = os.path.join(DATA_DIR, 'Warranty Debit.xlsx')
    CURRENT_MONTH_FILE = os.path.join(DATA_DIR, 'Pending Warranty Claim Details.xlsx')
    COMPENSATION_FILE = os.path.join(DATA_DIR, 'Transit_Claims_Merged.xlsx')
    PR_APPROVAL_FILE = os.path.join(DATA_DIR, 'Pr_Approval_Claims_Merged.xlsx')
    IMAGE_FOLDER = os.path.join(DATA_DIR, 'Image')
    print(f"📂 Using Render Disk Mount: {DATA_DIR}")
else:
    WARRANTY_FILE = r'D:\Power BI New\Warranty Debit\Warranty Debit.xlsx'
    CURRENT_MONTH_FILE = r'D:\Power BI New\Warranty Debit\Pending Warranty Claim Details.xlsx'
    COMPENSATION_FILE = r'D:\Power BI New\Warranty Debit\Transit_Claims_Merged.xlsx'
    PR_APPROVAL_FILE = r'D:\Power BI New\warranty dashboard render\Pr_Approval_Claims_Merged.xlsx'
    IMAGE_FOLDER = r'D:\Power BI New\Warranty Debit\Image'
    print(f"📂 Using Local Path: D:\\Power BI New\\...")

# ==================== WARRANTY DATA STORAGE ====================

WARRANTY_DATA = {
    'credit_df': None,
    'debit_df': None,
    'arbitration_df': None,
    'source_df': None,
    'current_month_df': None,
    'current_month_source_df': None,
    'compensation_df': None,
    'compensation_source_df': None,
    'pr_approval_df': None,
    'pr_approval_source_df': None
}

# ==================== DATA PROCESSING FUNCTIONS ====================

def process_pr_approval():
    """Process PR Approval data"""
    try:
        if not os.path.exists(PR_APPROVAL_FILE):
            print(f"⚠️  PR Approval file not found: {PR_APPROVAL_FILE}")
            return None, None
            
        df = pd.read_excel(PR_APPROVAL_FILE)
        print("✓ PR Approval data loaded successfully")
        print(f"  Rows: {len(df)}")

        summary_columns = ['Division', 'PA Request No.', 'PA Date', 'Request Type', 'App. Claim Amt from M&M']
        available_columns = [col for col in summary_columns if col in df.columns]
        
        if not available_columns:
            print(f"⚠️  No required columns in PR Approval")
            return None, None

        df_summary = df[available_columns].copy()
        
        if 'Division' in df_summary.columns:
            df_summary['Division'] = df_summary['Division'].astype(str).str.strip()
            df_summary = df_summary[df_summary['Division'].notna() & (df_summary['Division'] != '') & (df_summary['Division'] != 'nan')]
        
        if 'App. Claim Amt from M&M' in df_summary.columns:
            df_summary['App. Claim Amt from M&M'] = pd.to_numeric(df_summary['App. Claim Amt from M&M'], errors='coerce').fillna(0)
        
        summary_data = []
        if 'Division' in df_summary.columns:
            for division in sorted(df_summary['Division'].unique()):
                div_data = df_summary[df_summary['Division'] == division]
                summary_row = {'Division': division, 'Total Requests': len(div_data)}
                if 'App. Claim Amt from M&M' in df_summary.columns:
                    summary_row['Total Approved Amount'] = div_data['App. Claim Amt from M&M'].sum()
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            grand_total = {'Division': 'Grand Total'}
            for col in summary_df.columns:
                if col != 'Division' and summary_df[col].dtype in ['int64', 'float64']:
                    grand_total[col] = summary_df[col].sum()
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        return summary_df, df
    except Exception as e:
        print(f"⚠️  PR Approval error: {e}")
        return None, None

def process_compensation_claim():
    """Process compensation claim data"""
    try:
        if not os.path.exists(COMPENSATION_FILE):
            print(f"⚠️  Compensation file not found: {COMPENSATION_FILE}")
            return None, None
            
        df = pd.read_excel(COMPENSATION_FILE)
        print("✓ Compensation Claim data loaded successfully")
        print(f"  Rows: {len(df)}")

        required_columns = ['Division', 'RO Id.', 'Registration No.', 'RO Date', 'Claim Amount', 'Claim Approved Amt.']
        available_columns = [col for col in required_columns if col in df.columns]
        
        if not available_columns:
            print(f"⚠️  No required columns in Compensation")
            return None, None

        df_filtered = df[available_columns].copy()
        
        if 'Division' in df_filtered.columns:
            df_filtered['Division'] = df_filtered['Division'].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered['Division'].notna() & (df_filtered['Division'] != '') & (df_filtered['Division'] != 'nan')]
        
        numeric_cols = ['Claim Amount', 'Claim Approved Amt.']
        for col in numeric_cols:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
        
        summary_data = []
        if 'Division' in df_filtered.columns:
            for division in sorted(df_filtered['Division'].unique()):
                div_data = df_filtered[df_filtered['Division'] == division]
                summary_row = {'Division': division, 'Total Claims': len(div_data)}
                if 'Claim Amount' in df_filtered.columns:
                    summary_row['Total Claim Amount'] = div_data['Claim Amount'].sum()
                if 'Claim Approved Amt.' in df_filtered.columns:
                    summary_row['Total Approved Amount'] = div_data['Claim Approved Amt.'].sum()
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            grand_total = {'Division': 'Grand Total'}
            for col in summary_df.columns:
                if col != 'Division' and summary_df[col].dtype in ['int64', 'float64']:
                    grand_total[col] = summary_df[col].sum()
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        return summary_df, df_filtered
    except Exception as e:
        print(f"⚠️  Compensation error: {e}")
        return None, None

def process_current_month_warranty():
    """Process current month warranty data"""
    try:
        if not os.path.exists(CURRENT_MONTH_FILE):
            print(f"⚠️  Current month file not found: {CURRENT_MONTH_FILE}")
            return None, None
            
        df = pd.read_excel(CURRENT_MONTH_FILE, sheet_name='Pending Warranty Claim Details')
        print("✓ Current Month Warranty data loaded successfully")
        print(f"  Rows: {len(df)}")

        required_columns = ['Division', 'Pending Claims Spares', 'Pending Claims Labour']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"⚠️  Missing columns: {missing_columns}")
            return None, None

        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]

        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            spares_count = div_data['Pending Claims Spares'].notna().sum()
            labour_count = div_data['Pending Claims Labour'].notna().sum()
            summary_data.append({
                'Division': division,
                'Pending Claims Spares Count': spares_count,
                'Pending Claims Labour Count': labour_count,
                'Total Pending Claims': spares_count + labour_count
            })
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            'Division': 'Grand Total',
            'Pending Claims Spares Count': summary_df['Pending Claims Spares Count'].sum(),
            'Pending Claims Labour Count': summary_df['Pending Claims Labour Count'].sum(),
            'Total Pending Claims': summary_df['Total Pending Claims'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)

        return summary_df, df
    except Exception as e:
        print(f"⚠️  Current Month error: {e}")
        return None, None

def process_warranty_data():
    """Process warranty data"""
    try:
        if not os.path.exists(WARRANTY_FILE):
            print(f"⚠️  Warranty file not found: {WARRANTY_FILE}")
            return None, None, None, None
            
        df = pd.read_excel(WARRANTY_FILE, sheet_name='Sheet1')
        print("✓ Warranty data loaded successfully")
        print(f"  Rows: {len(df)}")

        dealer_mapping = {
            'AMRAVATI': 'AMT', 'CHAUFULA_SZZ': 'CHA', 'CHIKHALI': 'CHI',
            'KOLHAPUR_WS': 'KOL', 'NAGPUR_KAMPTHEE ROAD': 'HO',
            'NAGPUR_WARDHAMAN NGR': 'CITY', 'SHIKRAPUR_SZS': 'SHI',
            'WAGHOLI': 'WAG', 'YAVATMAL': 'YAT', 'NAGPUR_WARDHAMAN NGR_CQ': 'CQ'
        }

        numeric_columns = ['Total Claim Amount', 'Credit Note Amount', 'Debit Note Amount']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df['Dealer_Code'] = df['Dealer Location'].map(dealer_mapping).fillna(df['Dealer Location'])
        df['Month'] = df['Fiscal Month'].astype(str).str.strip().str[:3]
        df['Claim arbitration ID'] = df['Claim arbitration ID'].astype(str).replace('nan', '').replace('', np.nan)

        dealers = sorted(df['Dealer_Code'].unique())
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov']

        # Credit Note Table
        credit_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Credit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Credit Note {month}']
                credit_df = credit_df.merge(summary, on='Division', how='left')
            else:
                credit_df[f'Credit Note {month}'] = 0
        
        credit_df = credit_df.fillna(0)
        credit_columns = [f'Credit Note {month}' for month in months]
        credit_df['Total Credit'] = credit_df[credit_columns].sum(axis=1)
        grand_total_credit = {'Division': 'Grand Total'}
        for col in credit_df.columns[1:]:
            grand_total_credit[col] = credit_df[col].sum()
        credit_df = pd.concat([credit_df, pd.DataFrame([grand_total_credit])], ignore_index=True)

        # Debit Note Table
        debit_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Debit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Debit Note {month}']
                debit_df = debit_df.merge(summary, on='Division', how='left')
            else:
                debit_df[f'Debit Note {month}'] = 0
        
        debit_df = debit_df.fillna(0)
        debit_columns = [f'Debit Note {month}' for month in months]
        debit_df['Total Debit'] = debit_df[debit_columns].sum(axis=1)
        grand_total_debit = {'Division': 'Grand Total'}
        for col in debit_df.columns[1:]:
            grand_total_debit[col] = debit_df[col].sum()
        debit_df = pd.concat([debit_df, pd.DataFrame([grand_total_debit])], ignore_index=True)

        # Arbitration Table
        arbitration_df = pd.DataFrame({'Division': dealers})
        def is_arbitration(value):
            if pd.isna(value): return False
            value = str(value).strip().upper()
            return value.startswith('ARB') and value != 'NAN'

        for month in months:
            month_data = df[df['Month'] == month].copy()
            month_data['Is_ARB'] = month_data['Claim arbitration ID'].apply(is_arbitration)
            month_data['Arbitration_Amount'] = month_data.apply(lambda row: row['Debit Note Amount'] if row['Is_ARB'] else 0, axis=1)
            arb_summary = month_data.groupby('Dealer_Code')['Arbitration_Amount'].sum().reset_index()
            arb_summary.columns = ['Division', f'Claim Arbitration {month}']
            arbitration_df = arbitration_df.merge(arb_summary, on='Division', how='left')
        
        arbitration_df = arbitration_df.fillna(0)
        arbitration_cols = [f'Claim Arbitration {m}' for m in months]
        total_debit_by_dealer = debit_df[debit_df['Division'] != 'Grand Total'][['Division', 'Total Debit']].copy()
        arbitration_df = arbitration_df.merge(total_debit_by_dealer, on='Division', how='left')
        arbitration_df['Pending Claim Arbitration'] = arbitration_df['Total Debit'] - arbitration_df[arbitration_cols].sum(axis=1)
        arbitration_df = arbitration_df.drop('Total Debit', axis=1)
        grand_total_arb = {'Division': 'Grand Total'}
        for col in arbitration_df.columns[1:]:
            grand_total_arb[col] = arbitration_df[col].sum()
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([grand_total_arb])], ignore_index=True)

        return credit_df, debit_df, arbitration_df, df
    except Exception as e:
        print(f"⚠️  Warranty data error: {e}")
        return None, None, None, None

# ==================== DASHBOARD HTML ====================

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Warranty Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%); min-height: 100vh; }
        .navbar { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 15px 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); position: sticky; top: 0; z-index: 100; }
        .navbar h1 { font-size: 24px; font-weight: 700; }
        .container { max-width: 1400px; margin: 30px auto; padding: 0 20px; }
        .dashboard-content { background: white; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; }
        .nav-tabs { border-bottom: 2px solid #FF8C00; margin-bottom: 30px; display: flex; flex-wrap: wrap; gap: 10px; }
        .nav-link { color: #666; font-weight: 600; border: none; border-bottom: 3px solid transparent; padding: 12px 20px; cursor: pointer; background: none; transition: all 0.3s; }
        .nav-link:hover { color: #FF8C00; border-bottom-color: #FF8C00; }
        .nav-link.active { color: #FF8C00; border-bottom-color: #FF8C00; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .table-wrapper { overflow-x: auto; }
        .data-table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
        .data-table thead th { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 12px; text-align: center; font-weight: 600; }
        .data-table tbody td { padding: 10px 12px; border-bottom: 1px solid #e0e0e0; text-align: right; }
        .data-table tbody td:first-child { text-align: left; font-weight: 600; color: #333; }
        .data-table tbody tr:hover { background: #f9f9f9; }
        .data-table tbody tr:last-child { background: #fff8f3; font-weight: 700; border-top: 2px solid #FF8C00; }
        .loading { text-align: center; padding: 40px; }
        .spinner { border: 4px solid rgba(255,140,0,0.2); border-top: 4px solid #FF8C00; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .export-section { margin: 30px 0; padding: 20px; background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%); border-radius: 8px; border-left: 5px solid #FF8C00; }
        .export-section h3 { color: #FF8C00; margin-bottom: 15px; font-weight: 700; }
        .export-controls { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; background: white; padding: 15px; border-radius: 6px; }
        .export-controls select { padding: 8px 12px; border: 2px solid #FF8C00; border-radius: 4px; font-size: 13px; }
        .export-btn { padding: 10px 25px; background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 700; transition: all 0.3s; }
        .export-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(76,175,80,0.3); }
        .error-msg { background: #ffebee; color: #c62828; padding: 15px; border-radius: 6px; margin: 20px 0; }
    </style>
</head>
<body>
    <nav class="navbar">
        <h1>📊 Warranty Management Dashboard</h1>
    </nav>
    
    <div class="container">
        <div class="dashboard-content">
            <div class="loading" id="loadingSpinner">
                <div class="spinner"></div>
                <p style="margin-top: 15px; color: #666;">Loading warranty data...</p>
            </div>
            
            <div id="warrantyTabs" style="display: none;">
                <div class="nav-tabs">
                    <button class="nav-link active" onclick="switchTab('credit')">💰 Credit</button>
                    <button class="nav-link" onclick="switchTab('debit')">💳 Debit</button>
                    <button class="nav-link" onclick="switchTab('arbitration')">⚖️ Arbitration</button>
                    <button class="nav-link" onclick="switchTab('currentmonth')">📅 Current Month</button>
                    <button class="nav-link" onclick="switchTab('compensation')">🚚 Compensation</button>
                    <button class="nav-link" onclick="switchTab('pr_approval')">✅ PR Approval</button>
                </div>

                <div class="export-section">
                    <h3>📥 Export to Excel</h3>
                    <div class="export-controls">
                        <select id="divisionFilter"><option value="">-- Select Division --</option><option value="All">All Divisions</option></select>
                        <select id="exportType"><option value="credit">Credit Note</option><option value="debit">Debit Note</option><option value="arbitration">Arbitration</option><option value="currentmonth">Current Month</option><option value="compensation">Compensation</option><option value="pr_approval">PR Approval</option></select>
                        <button onclick="exportToExcel()" class="export-btn">📥 Export</button>
                    </div>
                </div>

                <div id="credit" class="tab-content active"><div class="table-wrapper"><table class="data-table" id="creditTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="debit" class="tab-content"><div class="table-wrapper"><table class="data-table" id="debitTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="arbitration" class="tab-content"><div class="table-wrapper"><table class="data-table" id="arbitrationTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="currentmonth" class="tab-content"><div class="table-wrapper"><table class="data-table" id="currentMonthTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="compensation" class="tab-content"><div class="table-wrapper"><table class="data-table" id="compensationTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="pr_approval" class="tab-content"><div class="table-wrapper"><table class="data-table" id="prApprovalTable"><thead></thead><tbody></tbody></table></div></div>
            </div>
        </div>
    </div>
    
    <script>
        let warrantyData = {};
        
        async function loadDashboard() {
            try {
                const response = await fetch('/api/warranty-data');
                warrantyData = await response.json();
                
                const spinner = document.getElementById('loadingSpinner');
                const tabs = document.getElementById('warrantyTabs');
                
                displayTable('creditTable', warrantyData.credit);
                displayTable('debitTable', warrantyData.debit);
                displayTable('arbitrationTable', warrantyData.arbitration);
                displayTable('currentMonthTable', warrantyData.currentMonth);
                displayTable('compensationTable', warrantyData.compensation);
                displayTable('prApprovalTable', warrantyData.prApproval);
                loadDivisions();
                
                spinner.style.display = 'none';
                tabs.style.display = 'block';
            } catch (error) {
                document.getElementById('loadingSpinner').innerHTML = '<div class="error-msg">⚠️ Error loading data. Check if Excel files are uploaded.</div>';
            }
        }
        
        function displayTable(tableId, data) {
            if (!data || data.length === 0) {
                document.getElementById(tableId).innerHTML = '<tr><td>No data available</td></tr>';
                return;
            }
            const table = document.getElementById(tableId);
            const headers = Object.keys(data[0]);
            table.querySelector('thead').innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            table.querySelector('tbody').innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 2}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function switchTab(tabName) {
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.nav-link').forEach(btn => btn.classList.remove('active'));
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
        }

        function loadDivisions() {
            const divisions = new Set();
            if (warrantyData.credit && warrantyData.credit.length > 0) {
                warrantyData.credit.forEach(row => {
                    if (row.Division && row.Division !== 'Grand Total') divisions.add(row.Division);
                });
            }
            const divisionSelect = document.getElementById('divisionFilter');
            divisionSelect.innerHTML = '<option value="">-- Select Division --</option><option value="All">All Divisions</option>';
            Array.from(divisions).sort().forEach(div => {
                const option = document.createElement('option');
                option.value = div;
                option.textContent = div;
                divisionSelect.appendChild(option);
            });
        }

        async function exportToExcel() {
            const division = document.getElementById('divisionFilter').value;
            const type = document.getElementById('exportType').value;
            if (!division) { alert('Select a division'); return; }
            
            try {
                const response = await fetch('/api/export-to-excel', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({division: division, type: type})
                });
                
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${type}_${division}.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (error) {
                alert('Export failed: ' + error.message);
            }
        }
        
        window.onload = function() { loadDashboard(); };
    </script>
</body>
</html>
"""

# ==================== FASTAPI APPLICATION ====================

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==================== API ENDPOINTS ====================

@app.post("/api/export-to-excel")
async def export_to_excel(request_data: dict):
    """Export data to Excel"""
    try:
        selected_division = request_data.get('division', 'All')
        export_type = request_data.get('type', 'credit')
        
        if export_type == 'credit':
            df = WARRANTY_DATA['credit_df']
        elif export_type == 'debit':
            df = WARRANTY_DATA['debit_df']
        elif export_type == 'arbitration':
            df = WARRANTY_DATA['arbitration_df']
        elif export_type == 'currentmonth':
            df = WARRANTY_DATA['current_month_df']
        elif export_type == 'compensation':
            df = WARRANTY_DATA['compensation_df']
        else:
            df = WARRANTY_DATA['pr_approval_df']
        
        if df is None or df.empty:
            raise HTTPException(status_code=500, detail="No data available")
        
        if selected_division != 'All' and selected_division != 'Grand Total':
            df_export = df[df['Division'] == selected_division].copy()
            grand_total = df[df['Division'] == 'Grand Total']
            if not grand_total.empty:
                df_export = pd.concat([df_export, grand_total], ignore_index=True)
        else:
            df_export = df.copy()
        
        wb = Workbook()
        ws = wb.active
        ws.title = export_type[:15]
        
        for col_idx, column in enumerate(df_export.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=column)
            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if isinstance(value, (int, float)):
                    cell.number_format = '#,##0.00'
        
        for col_idx, column in enumerate(df_export.columns, 1):
            max_length = min(max(df_export[column].astype(str).map(len).max(), len(str(column))) + 2, 30)
            ws.column_dimensions[chr(64 + col_idx)].width = max_length
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{export_type}_{selected_division}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/warranty-data")
async def get_warranty_data():
    """Get warranty data as JSON"""
    try:
        credit_records = WARRANTY_DATA['credit_df'].to_dict('records') if WARRANTY_DATA['credit_df'] is not None else []
        debit_records = WARRANTY_DATA['debit_df'].to_dict('records') if WARRANTY_DATA['debit_df'] is not None else []
        arbitration_records = WARRANTY_DATA['arbitration_df'].to_dict('records') if WARRANTY_DATA['arbitration_df'] is not None else []
        current_month_records = WARRANTY_DATA['current_month_df'].to_dict('records') if WARRANTY_DATA['current_month_df'] is not None else []
        compensation_records = WARRANTY_DATA['compensation_df'].to_dict('records') if WARRANTY_DATA['compensation_df'] is not None else []
        pr_approval_records = WARRANTY_DATA['pr_approval_df'].to_dict('records') if WARRANTY_DATA['pr_approval_df'] is not None else []
        
        for records in [credit_records, debit_records, arbitration_records, current_month_records, compensation_records, pr_approval_records]:
            for record in records:
                for key in record:
                    if pd.isna(record[key]):
                        record[key] = 0
        
        return {
            "credit": credit_records,
            "debit": debit_records,
            "arbitration": arbitration_records,
            "currentMonth": current_month_records,
            "compensation": compensation_records,
            "prApproval": pr_approval_records
        }
    except Exception as e:
        print(f"Error: {e}")
        return {"error": str(e)}

@app.get("/")
async def root():
    """Serve dashboard"""
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/health")
async def health():
    """Health check endpoint"""
    return {"status": "ok", "environment": "render" if IS_RENDER else "local"}

# ==================== STARTUP ====================

print("\n📂 Checking for Excel files...")
print(f"  Warranty File: {WARRANTY_FILE} {'✓' if os.path.exists(WARRANTY_FILE) else '❌'}")
print(f"  Current Month File: {CURRENT_MONTH_FILE} {'✓' if os.path.exists(CURRENT_MONTH_FILE) else '❌'}")
print(f"  Compensation File: {COMPENSATION_FILE} {'✓' if os.path.exists(COMPENSATION_FILE) else '❌'}")
print(f"  PR Approval File: {PR_APPROVAL_FILE} {'✓' if os.path.exists(PR_APPROVAL_FILE) else '❌'}")

print("\n⚙️  Processing data...")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'], WARRANTY_DATA['source_df'] = process_warranty_data()
WARRANTY_DATA['current_month_df'], WARRANTY_DATA['current_month_source_df'] = process_current_month_warranty()
WARRANTY_DATA['compensation_df'], WARRANTY_DATA['compensation_source_df'] = process_compensation_claim()
WARRANTY_DATA['pr_approval_df'], WARRANTY_DATA['pr_approval_source_df'] = process_pr_approval()

print(f"\n{'='*80}")
print(f"✅ DASHBOARD READY")
print(f"{'='*80}")
print(f"📊 Data Loaded: {sum(1 for v in WARRANTY_DATA.values() if v is not None)} modules")
print(f"🌐 Running on: http://0.0.0.0:$PORT")
print(f"{'='*80}\n")

if __name__ == "__main__":
    port = int(os.getenv('PORT', 8000))
    print(f"🚀 Starting server on port {port}...")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
