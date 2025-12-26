import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import sys
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from PIL import Image, ImageDraw, ImageFont
import secrets
import base64

# ==================== ENVIRONMENT SETUP ====================

IS_RENDER = os.getenv('RENDER') == 'true'
DATA_DIR = os.getenv('DATA_DIR', '/mnt/data')
PORT = int(os.getenv('PORT', 8001))

print(f"\n{'='*100}")
print(f"🚀 WARRANTY MANAGEMENT DASHBOARD - NO LOGIN")
print(f"{'='*100}")
print(f"Environment: {'RENDER ☁️' if IS_RENDER else 'LOCAL 💻'}")
print(f"Data Directory: {DATA_DIR}")
print(f"Port: {PORT}")
print(f"Access Mode: DIRECT - No Login Required")
print(f"{'='*100}\n")

# ==================== FILE PATH CONFIGURATION ====================

def get_file_path(filename):
    """Get file path based on environment"""
    if IS_RENDER:
        return os.path.join(DATA_DIR, filename)
    else:
        base_paths = {
            'Warranty Debit.xlsx': r'D:\Power BI New\Warranty Debit\Warranty Debit.xlsx',
            'Pending Warranty Claim Details.xlsx': r'D:\Power BI New\Warranty Debit\Pending Warranty Claim Details.xlsx',
            'Transit_Claims_Merged.xlsx': r'D:\Power BI New\Warranty Debit\Transit_Claims_Merged.xlsx',
            'Pr_Approval_Claims_Merged.xlsx': r'D:\Power BI New\warranty dashboard render\Pr_Approval_Claims_Merged.xlsx',
            'Image': r'D:\Power BI New\Warranty Debit\Image'
        }
        return base_paths.get(filename, filename)

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
    input_path = get_file_path('Pr_Approval_Claims_Merged.xlsx')
    
    try:
        if not os.path.exists(input_path):
            print(f"⚠️  PR Approval file not found: {input_path}")
            return None, None
            
        df = pd.read_excel(input_path)
        print("✓ PR Approval data loaded successfully")
        print(f"  Total rows: {len(df)}")

        summary_columns = ['Division', 'PA Request No.', 'PA Date', 'Request Type', 'App. Claim Amt from M&M']
        available_columns = [col for col in summary_columns if col in df.columns]
        
        if not available_columns:
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
        print(f"❌ Error processing PR Approval: {e}")
        return None, None

def process_compensation_claim():
    """Process compensation claim data"""
    input_path = get_file_path('Transit_Claims_Merged.xlsx')
    
    try:
        if not os.path.exists(input_path):
            print(f"⚠️  Compensation file not found: {input_path}")
            return None, None
            
        df = pd.read_excel(input_path)
        print("✓ Compensation Claim data loaded successfully")
        print(f"  Total rows: {len(df)}")

        required_columns = ['Division', 'RO Id.', 'Registration No.', 'RO Date', 'RO Bill Date', 'Chassis No.', 'Model Group', 'Claim Amount', 'Request Status', 'Claim Approved Amt.', 'No. of Days']
        available_columns = [col for col in required_columns if col in df.columns]
        
        if not available_columns:
            return None, None

        df_filtered = df[available_columns].copy()
        
        if 'Division' in df_filtered.columns:
            df_filtered['Division'] = df_filtered['Division'].astype(str).str.strip()
            df_filtered = df_filtered[df_filtered['Division'].notna() & (df_filtered['Division'] != '') & (df_filtered['Division'] != 'nan')]
        
        if 'RO Id.' in df_filtered.columns:
            def format_ro_id(x):
                if pd.isna(x) or str(x).strip() == '':
                    return ''
                try:
                    return f"RO{str(int(float(x)))}"
                except (ValueError, TypeError):
                    value_str = str(x).strip()
                    if not value_str.startswith('RO'):
                        return f"RO{value_str}"
                    return value_str
            df_filtered['RO Id.'] = df_filtered['RO Id.'].apply(format_ro_id)
        
        numeric_cols = ['Claim Amount', 'Claim Approved Amt.', 'No. of Days']
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
                if 'No. of Days' in df_filtered.columns:
                    summary_row['Avg No. of Days'] = div_data['No. of Days'].mean()
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
        print(f"❌ Error processing Compensation: {e}")
        return None, None

def process_current_month_warranty():
    """Process current month warranty data"""
    input_path = get_file_path('Pending Warranty Claim Details.xlsx')
    
    try:
        if not os.path.exists(input_path):
            print(f"⚠️  Current month file not found: {input_path}")
            return None, None
            
        df = pd.read_excel(input_path, sheet_name='Pending Warranty Claim Details')
        print("✓ Current Month Warranty data loaded successfully")
        print(f"  Total rows: {len(df)}")

        required_columns = ['Division', 'Pending Claims Spares', 'Pending Claims Labour']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
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
        print(f"❌ Error processing Current Month: {e}")
        return None, None

def process_warranty_data():
    """Process warranty data"""
    input_path = get_file_path('Warranty Debit.xlsx')
    
    try:
        if not os.path.exists(input_path):
            print(f"⚠️  Warranty file not found: {input_path}")
            return None, None, None, None
            
        df = pd.read_excel(input_path, sheet_name='Sheet1')
        print("✓ Warranty data loaded successfully")
        print(f"  Total rows: {len(df)}")

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
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        # CREDIT NOTE TABLE
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

        # DEBIT NOTE TABLE
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

        # ARBITRATION TABLE
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
        print(f"❌ Error processing warranty data: {e}")
        return None, None, None, None

# ==================== DASHBOARD HTML ====================

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Warranty Management Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; background: #f5f5f5; }
        .navbar { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); position: sticky; top: 0; z-index: 100; }
        .navbar h1 { font-size: 28px; font-weight: 700; }
        .container { max-width: 1400px; margin: 30px auto; padding: 0 20px; }
        .dashboard { background: white; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; border-bottom: 2px solid #FF8C00; }
        .tab-btn { padding: 12px 20px; border: none; background: transparent; cursor: pointer; font-weight: 600; color: #666; border-bottom: 3px solid transparent; transition: all 0.3s; }
        .tab-btn:hover, .tab-btn.active { color: #FF8C00; border-bottom-color: #FF8C00; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
        th { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 12px; text-align: center; font-weight: 600; }
        td { padding: 10px; border-bottom: 1px solid #eee; text-align: right; }
        td:first-child { text-align: left; font-weight: 600; }
        tr:hover { background: #f9f9f9; }
        tr:last-child { background: #fff8f3; font-weight: 700; border-top: 2px solid #FF8C00; }
        .loading { text-align: center; padding: 40px; }
        .spinner { border: 4px solid #ddd; border-top: 4px solid #FF8C00; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .export-section { background: #fff8f3; padding: 20px; border-radius: 8px; border-left: 5px solid #FF8C00; margin-bottom: 20px; }
        .export-section h3 { color: #FF8C00; margin-bottom: 15px; }
        .export-controls { display: flex; gap: 15px; flex-wrap: wrap; background: white; padding: 15px; border-radius: 6px; }
        .export-controls select { padding: 8px 12px; border: 2px solid #FF8C00; border-radius: 4px; }
        .export-btn { padding: 10px 25px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 700; transition: all 0.3s; }
        .export-btn:hover { background: #45a049; transform: translateY(-2px); }
    </style>
</head>
<body>
    <div class="navbar">
        <h1>📊 Warranty Management Dashboard</h1>
    </div>
    
    <div class="container">
        <div class="dashboard">
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>Loading data...</p>
            </div>
            
            <div id="content" style="display: none;">
                <div class="tabs">
                    <button class="tab-btn active" onclick="switchTab('credit')">💰 Credit</button>
                    <button class="tab-btn" onclick="switchTab('debit')">💳 Debit</button>
                    <button class="tab-btn" onclick="switchTab('arbitration')">⚖️ Arbitration</button>
                    <button class="tab-btn" onclick="switchTab('currentmonth')">📅 Current Month</button>
                    <button class="tab-btn" onclick="switchTab('compensation')">🚚 Compensation</button>
                    <button class="tab-btn" onclick="switchTab('pr_approval')">✅ PR Approval</button>
                </div>

                <div class="export-section">
                    <h3>📥 Export to Excel</h3>
                    <div class="export-controls">
                        <select id="divisionFilter"><option>Select Division</option></select>
                        <select id="exportType">
                            <option value="credit">Credit</option>
                            <option value="debit">Debit</option>
                            <option value="arbitration">Arbitration</option>
                            <option value="currentmonth">Current Month</option>
                            <option value="compensation">Compensation</option>
                            <option value="pr_approval">PR Approval</option>
                        </select>
                        <button onclick="exportData()" class="export-btn">📥 Export</button>
                    </div>
                </div>

                <div id="credit" class="tab-content active"><table id="creditTable"><thead></thead><tbody></tbody></table></div>
                <div id="debit" class="tab-content"><table id="debitTable"><thead></thead><tbody></tbody></table></div>
                <div id="arbitration" class="tab-content"><table id="arbitrationTable"><thead></thead><tbody></tbody></table></div>
                <div id="currentmonth" class="tab-content"><table id="currentMonthTable"><thead></thead><tbody></tbody></table></div>
                <div id="compensation" class="tab-content"><table id="compensationTable"><thead></thead><tbody></tbody></table></div>
                <div id="pr_approval" class="tab-content"><table id="prApprovalTable"><thead></thead><tbody></tbody></table></div>
            </div>
        </div>
    </div>
    
    <script>
        let allData = {};
        
        async function loadData() {
            try {
                const res = await fetch('/api/data');
                allData = await res.json();
                
                renderTable('creditTable', allData.credit);
                renderTable('debitTable', allData.debit);
                renderTable('arbitrationTable', allData.arbitration);
                renderTable('currentMonthTable', allData.currentMonth);
                renderTable('compensationTable', allData.compensation);
                renderTable('prApprovalTable', allData.prApproval);
                loadDivisions();
                
                document.getElementById('loading').style.display = 'none';
                document.getElementById('content').style.display = 'block';
            } catch (e) {
                document.getElementById('loading').innerHTML = '<p style="color: red;">⚠️ Error loading data</p>';
            }
        }
        
        function renderTable(tableId, data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById(tableId);
            const headers = Object.keys(data[0]);
            table.querySelector('thead').innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            table.querySelector('tbody').innerHTML = data.map(row => '<tr>' + 
                headers.map(h => '<td>' + (typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 2}) : row[h]) + '</td>').join('') + 
                '</tr>').join('');
        }
        
        function switchTab(tab) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            document.getElementById(tab).classList.add('active');
            event.target.classList.add('active');
        }
        
        function loadDivisions() {
            const divisions = new Set();
            (allData.credit || []).forEach(row => {
                if (row.Division && row.Division !== 'Grand Total') divisions.add(row.Division);
            });
            const select = document.getElementById('divisionFilter');
            select.innerHTML = '<option value="All">All Divisions</option>';
            Array.from(divisions).sort().forEach(div => {
                const opt = document.createElement('option');
                opt.value = div;
                opt.textContent = div;
                select.appendChild(opt);
            });
        }
        
        async function exportData() {
            const division = document.getElementById('divisionFilter').value;
            const type = document.getElementById('exportType').value;
            
            try {
                const res = await fetch('/api/export', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({division, type})
                });
                const blob = await res.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = type + '_' + division + '.xlsx';
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (e) {
                alert('Export failed: ' + e);
            }
        }
        
        loadData();
    </script>
</body>
</html>
"""

# ==================== FASTAPI SETUP ====================

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==================== API ENDPOINTS ====================

@app.get("/")
async def root():
    """Serve dashboard directly"""
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/api/data")
async def get_data():
    """Get all warranty data"""
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
        print(f"❌ Error: {e}")
        return {"error": str(e)}

@app.post("/api/export")
async def export_data(request: Request):
    """Export data to Excel"""
    try:
        data = await request.json()
        division = data.get('division', 'All')
        export_type = data.get('type', 'credit')
        
        # Select data
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
            return JSONResponse({"error": "No data"}, status_code=400)
        
        # Filter by division
        if division != 'All' and division != 'Grand Total':
            df_export = df[(df['Division'] == division) | (df['Division'] == 'Grand Total')].copy()
        else:
            df_export = df.copy()
        
        # Create Excel
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
        
        for col_idx in range(1, len(df_export.columns) + 1):
            max_length = min(20, max(len(str(cell.value)) for cell in ws.col_dimensions) + 2)
            ws.column_dimensions[chr(64 + col_idx)].width = max_length
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{export_type}_{division}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

# ==================== STARTUP ====================

print("📂 Checking files...")
for name in ['Warranty Debit.xlsx', 'Pending Warranty Claim Details.xlsx', 'Transit_Claims_Merged.xlsx', 'Pr_Approval_Claims_Merged.xlsx']:
    path = get_file_path(name)
    status = "✓" if os.path.exists(path) else "❌"
    print(f"  {status} {name}")

print("\n⚙️  Loading data...")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'], WARRANTY_DATA['source_df'] = process_warranty_data()
WARRANTY_DATA['current_month_df'], WARRANTY_DATA['current_month_source_df'] = process_current_month_warranty()
WARRANTY_DATA['compensation_df'], WARRANTY_DATA['compensation_source_df'] = process_compensation_claim()
WARRANTY_DATA['pr_approval_df'], WARRANTY_DATA['pr_approval_source_df'] = process_pr_approval()

print(f"\n{'='*100}")
print(f"✅ DASHBOARD READY")
print(f"{'='*100}")
if IS_RENDER:
    print(f"🌐 URL: https://warranty-management-dashboard.onrender.com")
else:
    print(f"🌐 URL: http://localhost:{PORT}")
print(f"✨ No login required - Opens directly to dashboard")
print(f"{'='*100}\n")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
