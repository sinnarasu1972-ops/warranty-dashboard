import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import uvicorn
from fastapi import FastAPI, Request, HTTPException, Cookie
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import sys
from typing import Optional
import hashlib
import secrets
import string
from PIL import Image, ImageDraw, ImageFont
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# ==================== ENVIRONMENT DETECTION ====================

IS_RENDER = os.getenv('RENDER') == 'true'
DATA_DIR = os.getenv('DATA_DIR', '/mnt/data')
PORT = int(os.getenv('PORT', 8001))

print(f"\n{'='*80}")
print(f"🚀 WARRANTY MANAGEMENT DASHBOARD - RENDER")
print(f"{'='*80}")
print(f"Environment: {'RENDER ☁️' if IS_RENDER else 'LOCAL 💻'}")
print(f"Data Directory: {DATA_DIR}")
print(f"Python: {sys.version.split()[0]}")
print(f"Port: {PORT}")
print(f"{'='*80}\n")

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
            'UserID.xlsx': r'D:\Power BI New\Warranty Debit\UserID.xlsx',
            'Image': r'D:\Power BI New\Warranty Debit\Image'
        }
        return base_paths.get(filename, filename)

# ==================== DATA STORAGE ====================

WARRANTY_DATA = {
    'credit_df': None,
    'debit_df': None,
    'arbitration_df': None,
    'current_month_df': None,
    'compensation_df': None,
    'pr_approval_df': None
}

# ==================== DATA PROCESSING ====================

def process_warranty_data():
    """Process warranty data - Credit, Debit, Arbitration"""
    try:
        filepath = get_file_path('Warranty Debit.xlsx')
        if not os.path.exists(filepath):
            print(f"⚠️ Warranty file not found: {filepath}")
            return None, None, None
        
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        print(f"✓ Warranty data loaded ({len(df)} rows)")
        
        # Dealer mapping
        dealer_mapping = {
            'AMRAVATI': 'AMT', 'CHAUFULA_SZZ': 'CHA', 'CHIKHALI': 'CHI',
            'KOLHAPUR_WS': 'KOL', 'NAGPUR_KAMPTHEE ROAD': 'HO',
            'NAGPUR_WARDHAMAN NGR': 'CITY', 'SHIKRAPUR_SZS': 'SHI',
            'WAGHOLI': 'WAG', 'YAVATMAL': 'YAT', 'NAGPUR_WARDHAMAN NGR_CQ': 'CQ'
        }
        
        # Clean numeric columns
        for col in ['Total Claim Amount', 'Credit Note Amount', 'Debit Note Amount']:
            if col in df.columns:
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
                summary.columns = ['Division', f'Credit {month}']
                credit_df = credit_df.merge(summary, on='Division', how='left')
            else:
                credit_df[f'Credit {month}'] = 0
        
        credit_df = credit_df.fillna(0)
        credit_df['Total'] = credit_df.iloc[:, 1:].sum(axis=1)
        grand_total = {'Division': 'TOTAL'}
        for col in credit_df.columns[1:]:
            grand_total[col] = credit_df[col].sum()
        credit_df = pd.concat([credit_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        # Debit Note Table
        debit_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Debit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Debit {month}']
                debit_df = debit_df.merge(summary, on='Division', how='left')
            else:
                debit_df[f'Debit {month}'] = 0
        
        debit_df = debit_df.fillna(0)
        debit_df['Total'] = debit_df.iloc[:, 1:].sum(axis=1)
        grand_total = {'Division': 'TOTAL'}
        for col in debit_df.columns[1:]:
            grand_total[col] = debit_df[col].sum()
        debit_df = pd.concat([debit_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        # Arbitration Table
        arbitration_df = pd.DataFrame({'Division': dealers})
        for month in months:
            month_data = df[df['Month'] == month].copy()
            month_data['Is_ARB'] = month_data['Claim arbitration ID'].apply(
                lambda x: str(x).upper().startswith('ARB') if pd.notna(x) else False
            )
            month_data['Arb_Amount'] = month_data.apply(
                lambda row: row['Debit Note Amount'] if row['Is_ARB'] else 0, axis=1
            )
            arb_summary = month_data.groupby('Dealer_Code')['Arb_Amount'].sum().reset_index()
            arb_summary.columns = ['Division', f'Arb {month}']
            arbitration_df = arbitration_df.merge(arb_summary, on='Division', how='left')
        
        arbitration_df = arbitration_df.fillna(0)
        arbitration_df['Total'] = arbitration_df.iloc[:, 1:].sum(axis=1)
        grand_total = {'Division': 'TOTAL'}
        for col in arbitration_df.columns[1:]:
            grand_total[col] = arbitration_df[col].sum()
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        return credit_df, debit_df, arbitration_df
        
    except Exception as e:
        print(f"❌ Error processing warranty: {e}")
        return None, None, None

def process_current_month_warranty():
    """Process current month warranty"""
    try:
        filepath = get_file_path('Pending Warranty Claim Details.xlsx')
        if not os.path.exists(filepath):
            print(f"⚠️ Current month file not found: {filepath}")
            return None
        
        df = pd.read_excel(filepath, sheet_name='Pending Warranty Claim Details')
        print(f"✓ Current month data loaded ({len(df)} rows)")
        
        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]
        
        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            summary_data.append({
                'Division': division,
                'Pending Spares': div_data['Pending Claims Spares'].notna().sum(),
                'Pending Labour': div_data['Pending Claims Labour'].notna().sum()
            })
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            'Division': 'TOTAL',
            'Pending Spares': summary_df['Pending Spares'].sum(),
            'Pending Labour': summary_df['Pending Labour'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        return summary_df
        
    except Exception as e:
        print(f"❌ Error processing current month: {e}")
        return None

def process_compensation_claim():
    """Process compensation claims"""
    try:
        filepath = get_file_path('Transit_Claims_Merged.xlsx')
        if not os.path.exists(filepath):
            print(f"⚠️ Compensation file not found: {filepath}")
            return None
        
        df = pd.read_excel(filepath)
        print(f"✓ Compensation data loaded ({len(df)} rows)")
        
        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]
        
        for col in ['Claim Amount', 'Claim Approved Amt.']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            summary_data.append({
                'Division': division,
                'Total Claims': len(div_data),
                'Claim Amount': div_data.get('Claim Amount', pd.Series()).sum() if 'Claim Amount' in df.columns else 0,
                'Approved Amount': div_data.get('Claim Approved Amt.', pd.Series()).sum() if 'Claim Approved Amt.' in df.columns else 0
            })
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            'Division': 'TOTAL',
            'Total Claims': summary_df['Total Claims'].sum(),
            'Claim Amount': summary_df['Claim Amount'].sum(),
            'Approved Amount': summary_df['Approved Amount'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        return summary_df
        
    except Exception as e:
        print(f"❌ Error processing compensation: {e}")
        return None

def process_pr_approval():
    """Process PR approval"""
    try:
        filepath = get_file_path('Pr_Approval_Claims_Merged.xlsx')
        if not os.path.exists(filepath):
            print(f"⚠️ PR Approval file not found: {filepath}")
            return None
        
        df = pd.read_excel(filepath)
        print(f"✓ PR Approval data loaded ({len(df)} rows)")
        
        if 'Division' not in df.columns:
            return None
        
        df['Division'] = df['Division'].astype(str).str.strip()
        df = df[df['Division'].notna() & (df['Division'] != '') & (df['Division'] != 'nan')]
        
        if 'App. Claim Amt from M&M' in df.columns:
            df['App. Claim Amt from M&M'] = pd.to_numeric(df['App. Claim Amt from M&M'], errors='coerce').fillna(0)
        
        summary_data = []
        for division in sorted(df['Division'].unique()):
            div_data = df[df['Division'] == division]
            summary_data.append({
                'Division': division,
                'Total Requests': len(div_data),
                'Approved Amount': div_data['App. Claim Amt from M&M'].sum() if 'App. Claim Amt from M&M' in df.columns else 0
            })
        
        summary_df = pd.DataFrame(summary_data)
        grand_total = {
            'Division': 'TOTAL',
            'Total Requests': summary_df['Total Requests'].sum(),
            'Approved Amount': summary_df['Approved Amount'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        
        return summary_df
        
    except Exception as e:
        print(f"❌ Error processing PR approval: {e}")
        return None

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
    """Serve dashboard"""
    return HTMLResponse(get_dashboard_html())

@app.get("/api/data")
async def get_data():
    """Get all warranty data"""
    return {
        "credit": WARRANTY_DATA['credit_df'].to_dict('records') if WARRANTY_DATA['credit_df'] is not None else [],
        "debit": WARRANTY_DATA['debit_df'].to_dict('records') if WARRANTY_DATA['debit_df'] is not None else [],
        "arbitration": WARRANTY_DATA['arbitration_df'].to_dict('records') if WARRANTY_DATA['arbitration_df'] is not None else [],
        "currentMonth": WARRANTY_DATA['current_month_df'].to_dict('records') if WARRANTY_DATA['current_month_df'] is not None else [],
        "compensation": WARRANTY_DATA['compensation_df'].to_dict('records') if WARRANTY_DATA['compensation_df'] is not None else [],
        "prApproval": WARRANTY_DATA['pr_approval_df'].to_dict('records') if WARRANTY_DATA['pr_approval_df'] is not None else []
    }

@app.post("/api/export")
async def export_data(request: Request):
    """Export data to Excel"""
    try:
        data = await request.json()
        sheet_type = data.get('type', 'credit')
        division_filter = data.get('division', 'All')
        
        # Select data based on type
        if sheet_type == 'credit':
            df = WARRANTY_DATA['credit_df']
        elif sheet_type == 'debit':
            df = WARRANTY_DATA['debit_df']
        elif sheet_type == 'arbitration':
            df = WARRANTY_DATA['arbitration_df']
        elif sheet_type == 'currentMonth':
            df = WARRANTY_DATA['current_month_df']
        elif sheet_type == 'compensation':
            df = WARRANTY_DATA['compensation_df']
        else:
            df = WARRANTY_DATA['pr_approval_df']
        
        if df is None or df.empty:
            return JSONResponse({"error": "No data"}, status_code=400)
        
        # Filter by division
        if division_filter != 'All':
            df = df[(df['Division'] == division_filter) | (df['Division'] == 'TOTAL')]
        
        # Create Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_type[:20]
        
        # Headers
        for col_idx, column in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=column)
            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")
        
        # Data
        for row_idx, row in enumerate(df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if isinstance(value, (int, float)):
                    cell.number_format = '#,##0.00'
        
        # Auto width
        for col in ws.columns:
            max_length = min(20, max(len(str(cell.value)) for cell in col))
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
        
        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=warranty_{sheet_type}.xlsx"}
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

def get_dashboard_html():
    """Get dashboard HTML"""
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Warranty Dashboard</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; background: #f5f5f5; }
            .navbar { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
            .navbar h1 { font-size: 28px; font-weight: 700; }
            .container { max-width: 1400px; margin: 20px auto; padding: 0 20px; }
            .tabs { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
            .tab-btn { padding: 10px 20px; border: none; background: #ddd; cursor: pointer; border-radius: 4px; font-weight: 600; transition: all 0.3s; }
            .tab-btn.active { background: #FF8C00; color: white; }
            .tab-btn:hover { background: #FF8C00; color: white; }
            .tab-content { display: none; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            .tab-content.active { display: block; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th { background: #FF8C00; color: white; padding: 10px; text-align: center; font-weight: 600; }
            td { padding: 8px; border-bottom: 1px solid #eee; text-align: right; }
            td:first-child { text-align: left; font-weight: 600; }
            tr:last-child { background: #fff8f3; border-top: 2px solid #FF8C00; font-weight: 700; }
            .loading { text-align: center; padding: 40px; }
            .spinner { border: 4px solid #ddd; border-top: 4px solid #FF8C00; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
            @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
            .export-section { background: #fff8f3; padding: 15px; border-radius: 6px; border-left: 5px solid #FF8C00; margin-bottom: 20px; }
            .export-section h3 { color: #FF8C00; margin-bottom: 10px; }
            .export-controls { display: flex; gap: 10px; flex-wrap: wrap; }
            .export-controls select, .export-btn { padding: 8px 12px; border: 1px solid #FF8C00; border-radius: 4px; }
            .export-btn { background: #4CAF50; color: white; cursor: pointer; font-weight: 600; }
            .export-btn:hover { background: #45a049; }
        </style>
    </head>
    <body>
        <div class="navbar">
            <h1>📊 Warranty Management Dashboard</h1>
        </div>
        
        <div class="container">
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Loading data...</p>
            </div>
            
            <div id="content" style="display: none;">
                <div class="tabs">
                    <button class="tab-btn active" onclick="switchTab('credit')">💰 Credit</button>
                    <button class="tab-btn" onclick="switchTab('debit')">💳 Debit</button>
                    <button class="tab-btn" onclick="switchTab('arbitration')">⚖️ Arbitration</button>
                    <button class="tab-btn" onclick="switchTab('currentMonth')">📅 Current Month</button>
                    <button class="tab-btn" onclick="switchTab('compensation')">🚚 Compensation</button>
                    <button class="tab-btn" onclick="switchTab('prApproval')">✅ PR Approval</button>
                </div>
                
                <div class="export-section">
                    <h3>📥 Export Data</h3>
                    <div class="export-controls">
                        <select id="divisionSelect"><option>Select Division</option></select>
                        <select id="typeSelect">
                            <option value="credit">Credit</option>
                            <option value="debit">Debit</option>
                            <option value="arbitration">Arbitration</option>
                            <option value="currentMonth">Current Month</option>
                            <option value="compensation">Compensation</option>
                            <option value="prApproval">PR Approval</option>
                        </select>
                        <button class="export-btn" onclick="exportData()">📥 Export</button>
                    </div>
                </div>
                
                <div id="credit" class="tab-content active"><table id="creditTable"><thead></thead><tbody></tbody></table></div>
                <div id="debit" class="tab-content"><table id="debitTable"><thead></thead><tbody></tbody></table></div>
                <div id="arbitration" class="tab-content"><table id="arbitrationTable"><thead></thead><tbody></tbody></table></div>
                <div id="currentMonth" class="tab-content"><table id="currentMonthTable"><thead></thead><tbody></tbody></table></div>
                <div id="compensation" class="tab-content"><table id="compensationTable"><thead></thead><tbody></tbody></table></div>
                <div id="prApproval" class="tab-content"><table id="prApprovalTable"><thead></thead><tbody></tbody></table></div>
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
                    document.getElementById('loading').innerHTML = '<p>⚠️ Error loading data</p>';
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
                    if (row.Division && row.Division !== 'TOTAL') divisions.add(row.Division);
                });
                const select = document.getElementById('divisionSelect');
                select.innerHTML = '<option value="All">All Divisions</option>';
                Array.from(divisions).sort().forEach(div => {
                    const opt = document.createElement('option');
                    opt.value = div;
                    opt.textContent = div;
                    select.appendChild(opt);
                });
            }
            
            async function exportData() {
                const division = document.getElementById('divisionSelect').value;
                const type = document.getElementById('typeSelect').value;
                if (!division) { alert('Select a division'); return; }
                
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

# ==================== STARTUP ====================

print("📂 Checking files...")
for name in ['Warranty Debit.xlsx', 'Pending Warranty Claim Details.xlsx', 'Transit_Claims_Merged.xlsx', 'Pr_Approval_Claims_Merged.xlsx']:
    path = get_file_path(name)
    status = "✓" if os.path.exists(path) else "❌"
    print(f"  {status} {name}")

print("\n⚙️ Loading data...")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'] = process_warranty_data()
WARRANTY_DATA['current_month_df'] = process_current_month_warranty()
WARRANTY_DATA['compensation_df'] = process_compensation_claim()
WARRANTY_DATA['pr_approval_df'] = process_pr_approval()

print(f"\n{'='*80}")
print(f"✅ DASHBOARD READY")
print(f"{'='*80}")
if IS_RENDER:
    print(f"🌐 URL: https://warranty-management-dashboard.onrender.com")
else:
    print(f"🌐 URL: http://localhost:{PORT}")
print(f"{'='*80}\n")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
