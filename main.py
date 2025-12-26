import os
import io
from pathlib import Path
from datetime import datetime

import numpy as np
import pandas as pd

import uvicorn
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# =========================================================
# ENV / PATHS
# =========================================================
IS_RENDER = os.getenv("RENDER", "").lower() == "true"
PORT = int(os.getenv("PORT", "8000"))
DATA_DIR = Path(os.getenv("DATA_DIR", "/mnt/data"))

BASE_DIR = Path(__file__).resolve().parent

# Where we search for files (Render + Local)
CANDIDATE_DIRS = [
    DATA_DIR,               # Render disk
    BASE_DIR,               # repo root
    BASE_DIR / "Data",      # repo Data folder
    BASE_DIR / "data",      # repo data folder
    Path("/opt/render/project/src"),
    Path("/opt/render/project/src/Data"),
    Path("/opt/render/project/src/data"),
]

FILES = {
    "Warranty Debit.xlsx": "Warranty Debit.xlsx",
    "Pending Warranty Claim Details.xlsx": "Pending Warranty Claim Details.xlsx",
    "Transit_Claims_Merged.xlsx": "Transit_Claims_Merged.xlsx",
    "Pr_Approval_Claims_Merged.xlsx": "Pr_Approval_Claims_Merged.xlsx",
}


def safe_listdir(p: Path):
    try:
        if p.exists() and p.is_dir():
            return sorted([x.name for x in p.iterdir()])
    except Exception:
        pass
    return []


def find_file(filename: str) -> Path | None:
    """
    Searches for a file in multiple locations.
    Works for Render and Local.
    """
    for d in CANDIDATE_DIRS:
        p = d / filename
        if p.exists() and p.is_file():
            return p
    return None


# =========================================================
# DATA STORAGE
# =========================================================
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

# =========================================================
# PROCESSING FUNCTIONS
# =========================================================
def process_pr_approval():
    p = find_file(FILES["Pr_Approval_Claims_Merged.xlsx"])
    if not p:
        print("PR Approval file not found.")
        return None, None
    try:
        df = pd.read_excel(p)
        print(f"PR Approval loaded: {p}")

        summary_columns = ["Division", "PA Request No.", "PA Date", "Request Type", "App. Claim Amt from M&M"]
        available = [c for c in summary_columns if c in df.columns]
        if not available:
            return None, df

        df_summary = df[available].copy()
        if "Division" in df_summary.columns:
            df_summary["Division"] = df_summary["Division"].astype(str).str.strip()
            df_summary = df_summary[df_summary["Division"].notna() & (df_summary["Division"] != "") & (df_summary["Division"] != "nan")]

        if "App. Claim Amt from M&M" in df_summary.columns:
            df_summary["App. Claim Amt from M&M"] = pd.to_numeric(df_summary["App. Claim Amt from M&M"], errors="coerce").fillna(0)

        if "Division" not in df_summary.columns:
            return pd.DataFrame(), df_summary

        rows = []
        for div in sorted(df_summary["Division"].unique()):
            d = df_summary[df_summary["Division"] == div]
            r = {"Division": div, "Total Requests": len(d)}
            if "App. Claim Amt from M&M" in df_summary.columns:
                r["Total Approved Amount"] = float(d["App. Claim Amt from M&M"].sum())
            rows.append(r)

        out = pd.DataFrame(rows)
        gt = {"Division": "Grand Total"}
        for c in out.columns:
            if c != "Division":
                gt[c] = out[c].sum()
        out = pd.concat([out, pd.DataFrame([gt])], ignore_index=True)
        return out, df_summary
    except Exception as e:
        print(f"Error processing PR Approval: {e}")
        return None, None


def process_compensation_claim():
    p = find_file(FILES["Transit_Claims_Merged.xlsx"])
    if not p:
        print("Compensation file not found.")
        return None, None

    try:
        df = pd.read_excel(p)
        print(f"Compensation loaded: {p}")

        required = [
            "Division", "RO Id.", "Registration No.", "RO Date", "RO Bill Date",
            "Chassis No.", "Model Group", "Claim Amount", "Request Status",
            "Claim Approved Amt.", "No. of Days"
        ]
        available = [c for c in required if c in df.columns]
        if not available:
            return None, df

        df2 = df[available].copy()
        if "Division" in df2.columns:
            df2["Division"] = df2["Division"].astype(str).str.strip()
            df2 = df2[df2["Division"].notna() & (df2["Division"] != "") & (df2["Division"] != "nan")]

        for c in ["Claim Amount", "Claim Approved Amt.", "No. of Days"]:
            if c in df2.columns:
                df2[c] = pd.to_numeric(df2[c], errors="coerce").fillna(0)

        if "Division" not in df2.columns:
            return pd.DataFrame(), df2

        rows = []
        for div in sorted(df2["Division"].unique()):
            d = df2[df2["Division"] == div]
            r = {"Division": div, "Total Claims": len(d)}
            if "Claim Amount" in df2.columns:
                r["Total Claim Amount"] = float(d["Claim Amount"].sum())
            if "Claim Approved Amt." in df2.columns:
                r["Total Approved Amount"] = float(d["Claim Approved Amt."].sum())
            if "No. of Days" in df2.columns:
                r["Avg No. of Days"] = float(d["No. of Days"].mean()) if len(d) else 0
            rows.append(r)

        out = pd.DataFrame(rows)
        gt = {"Division": "Grand Total"}
        for c in out.columns:
            if c != "Division":
                if c == "Avg No. of Days":
                    gt[c] = out[c].mean() if len(out) else 0
                else:
                    gt[c] = out[c].sum()
        out = pd.concat([out, pd.DataFrame([gt])], ignore_index=True)
        return out, df2
    except Exception as e:
        print(f"Error processing Compensation: {e}")
        return None, None


def process_current_month_warranty():
    p = find_file(FILES["Pending Warranty Claim Details.xlsx"])
    if not p:
        print("Current month file not found.")
        return None, None

    try:
        df = pd.read_excel(p, sheet_name="Pending Warranty Claim Details")
        print(f"Current Month loaded: {p}")

        required = ["Division", "Pending Claims Spares", "Pending Claims Labour"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"Missing columns in current month file: {missing}")
            return None, df

        df["Division"] = df["Division"].astype(str).str.strip()
        df = df[df["Division"].notna() & (df["Division"] != "") & (df["Division"] != "nan")]

        rows = []
        for div in sorted(df["Division"].unique()):
            d = df[df["Division"] == div]
            spares = int(d["Pending Claims Spares"].notna().sum())
            labour = int(d["Pending Claims Labour"].notna().sum())
            rows.append({
                "Division": div,
                "Pending Claims Spares Count": spares,
                "Pending Claims Labour Count": labour,
                "Total Pending Claims": spares + labour,
            })

        out = pd.DataFrame(rows)
        gt = {
            "Division": "Grand Total",
            "Pending Claims Spares Count": int(out["Pending Claims Spares Count"].sum()) if not out.empty else 0,
            "Pending Claims Labour Count": int(out["Pending Claims Labour Count"].sum()) if not out.empty else 0,
            "Total Pending Claims": int(out["Total Pending Claims"].sum()) if not out.empty else 0,
        }
        out = pd.concat([out, pd.DataFrame([gt])], ignore_index=True)
        return out, df
    except Exception as e:
        print(f"Error processing Current Month: {e}")
        return None, None


def process_warranty_data():
    p = find_file(FILES["Warranty Debit.xlsx"])
    if not p:
        print("Warranty file not found.")
        return None, None, None, None

    try:
        df = pd.read_excel(p, sheet_name="Sheet1")
        print(f"Warranty loaded: {p}")

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

        for c in ["Total Claim Amount", "Credit Note Amount", "Debit Note Amount"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
            else:
                df[c] = 0

        if "Dealer Location" not in df.columns:
            df["Dealer Location"] = ""
        if "Fiscal Month" not in df.columns:
            df["Fiscal Month"] = ""
        if "Claim arbitration ID" not in df.columns:
            df["Claim arbitration ID"] = ""

        df["Dealer_Code"] = df["Dealer Location"].map(dealer_mapping).fillna(df["Dealer Location"])
        df["Month"] = df["Fiscal Month"].astype(str).str.strip().str[:3]
        df["Claim arbitration ID"] = df["Claim arbitration ID"].astype(str).replace("nan", "").replace("", np.nan)

        dealers = sorted([x for x in df["Dealer_Code"].dropna().unique().tolist() if str(x).strip() != ""])
        months = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

        # CREDIT
        credit_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m]
            if not md.empty:
                s = md.groupby("Dealer_Code")["Credit Note Amount"].sum().reset_index()
                s.columns = ["Division", f"Credit Note {m}"]
                credit_df = credit_df.merge(s, on="Division", how="left")
            else:
                credit_df[f"Credit Note {m}"] = 0
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
            if not md.empty:
                s = md.groupby("Dealer_Code")["Debit Note Amount"].sum().reset_index()
                s.columns = ["Division", f"Debit Note {m}"]
                debit_df = debit_df.merge(s, on="Division", how="left")
            else:
                debit_df[f"Debit Note {m}"] = 0
        debit_df = debit_df.fillna(0)
        debit_cols = [f"Debit Note {m}" for m in months]
        debit_df["Total Debit"] = debit_df[debit_cols].sum(axis=1)
        gt = {"Division": "Grand Total"}
        for c in debit_df.columns[1:]:
            gt[c] = float(debit_df[c].sum())
        debit_df = pd.concat([debit_df, pd.DataFrame([gt])], ignore_index=True)

        # ARBITRATION
        def is_arb(v):
            if pd.isna(v):
                return False
            vv = str(v).strip().upper()
            return vv.startswith("ARB") and vv != "NAN"

        arbitration_df = pd.DataFrame({"Division": dealers})
        for m in months:
            md = df[df["Month"] == m].copy()
            if md.empty:
                arbitration_df[f"Claim Arbitration {m}"] = 0
                continue

            md["Is_ARB"] = md["Claim arbitration ID"].apply(is_arb)
            md["ArbAmt"] = md.apply(lambda r: r["Debit Note Amount"] if r["Is_ARB"] else 0, axis=1)
            s = md.groupby("Dealer_Code")["ArbAmt"].sum().reset_index()
            s.columns = ["Division", f"Claim Arbitration {m}"]
            arbitration_df = arbitration_df.merge(s, on="Division", how="left")

        arbitration_df = arbitration_df.fillna(0)
        arb_cols = [f"Claim Arbitration {m}" for m in months]
        total_debit = debit_df[debit_df["Division"] != "Grand Total"][["Division", "Total Debit"]].copy()
        arbitration_df = arbitration_df.merge(total_debit, on="Division", how="left")
        arbitration_df["Pending Claim Arbitration"] = arbitration_df["Total Debit"].fillna(0) - arbitration_df[arb_cols].sum(axis=1)
        arbitration_df = arbitration_df.drop(columns=["Total Debit"])

        gt = {"Division": "Grand Total"}
        for c in arbitration_df.columns[1:]:
            gt[c] = float(arbitration_df[c].sum())
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([gt])], ignore_index=True)

        return credit_df, debit_df, arbitration_df, df
    except Exception as e:
        print(f"Error processing warranty data: {e}")
        return None, None, None, None


# =========================================================
# DASHBOARD HTML (DIRECT OPEN)
# =========================================================
DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Warranty Management Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, sans-serif; background: #f5f5f5; }
        .navbar { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 18px 26px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); position: sticky; top: 0; z-index: 100; }
        .navbar h1 { font-size: 22px; font-weight: 800; }
        .container { max-width: 1400px; margin: 26px auto; padding: 0 18px; }
        .dashboard { background: white; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.10); padding: 22px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; border-bottom: 2px solid #FF8C00; }
        .tab-btn { padding: 10px 16px; border: none; background: transparent; cursor: pointer; font-weight: 800; color: #666; border-bottom: 3px solid transparent; transition: all 0.2s; }
        .tab-btn:hover, .tab-btn.active { color: #FF8C00; border-bottom-color: #FF8C00; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12px; }
        th { background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%); color: white; padding: 10px; text-align: center; font-weight: 800; white-space: nowrap; }
        td { padding: 9px 10px; border-bottom: 1px solid #eee; text-align: right; white-space: nowrap; }
        td:first-child { text-align: left; font-weight: 800; }
        tr:hover { background: #f9f9f9; }
        tr:last-child { background: #fff8f3; font-weight: 900; border-top: 2px solid #FF8C00; }
        .loading { text-align: center; padding: 40px; font-weight: 800; color: #666; }
        .spinner { border: 4px solid #ddd; border-top: 4px solid #FF8C00; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 12px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .export-section { background: #fff8f3; padding: 14px; border-radius: 10px; border-left: 5px solid #FF8C00; margin-bottom: 16px; }
        .export-section h3 { color: #FF8C00; margin-bottom: 10px; font-weight: 900; }
        .export-controls { display: flex; gap: 10px; flex-wrap: wrap; background: white; padding: 12px; border-radius: 10px; }
        .export-controls select { padding: 8px 12px; border: 2px solid #FF8C00; border-radius: 8px; font-weight: 800; }
        .export-btn { padding: 9px 22px; background: #4CAF50; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 900; }
        .export-btn:hover { background: #45a049; }
        .table-wrap { overflow-x: auto; }
    </style>
</head>
<body>
    <div class="navbar">
        <h1>Warranty Management Dashboard</h1>
    </div>

    <div class="container">
        <div class="dashboard">
            <div class="loading" id="loading">
                <div class="spinner"></div>
                Loading data
            </div>

            <div id="content" style="display:none;">
                <div class="tabs">
                    <button class="tab-btn active" onclick="switchTab('credit', event)">Credit</button>
                    <button class="tab-btn" onclick="switchTab('debit', event)">Debit</button>
                    <button class="tab-btn" onclick="switchTab('arbitration', event)">Arbitration</button>
                    <button class="tab-btn" onclick="switchTab('currentmonth', event)">Current Month</button>
                    <button class="tab-btn" onclick="switchTab('compensation', event)">Compensation</button>
                    <button class="tab-btn" onclick="switchTab('pr_approval', event)">PR Approval</button>
                </div>

                <div class="export-section">
                    <h3>Export to Excel</h3>
                    <div class="export-controls">
                        <select id="divisionFilter"></select>
                        <select id="exportType">
                            <option value="credit">Credit</option>
                            <option value="debit">Debit</option>
                            <option value="arbitration">Arbitration</option>
                            <option value="currentmonth">Current Month</option>
                            <option value="compensation">Compensation</option>
                            <option value="pr_approval">PR Approval</option>
                        </select>
                        <button onclick="exportData()" class="export-btn">Export</button>
                    </div>
                </div>

                <div id="credit" class="tab-content active"><div class="table-wrap"><table id="creditTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="debit" class="tab-content"><div class="table-wrap"><table id="debitTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="arbitration" class="tab-content"><div class="table-wrap"><table id="arbTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="currentmonth" class="tab-content"><div class="table-wrap"><table id="cmTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="compensation" class="tab-content"><div class="table-wrap"><table id="compTable"><thead></thead><tbody></tbody></table></div></div>
                <div id="pr_approval" class="tab-content"><div class="table-wrap"><table id="prTable"><thead></thead><tbody></tbody></table></div></div>
            </div>
        </div>
    </div>

    <script>
        let allData = {};

        function fmt(v){
            if(v === null || v === undefined) return "";
            if(typeof v === "number") return v.toLocaleString("en-IN", {maximumFractionDigits: 2});
            return v;
        }

        function renderTable(tableId, data){
            const table = document.getElementById(tableId);
            if(!data || !data.length){
                table.querySelector("thead").innerHTML = "";
                table.querySelector("tbody").innerHTML = "<tr><td style='text-align:left' colspan='50'>No data</td></tr>";
                return;
            }
            const headers = Object.keys(data[0]);
            table.querySelector("thead").innerHTML = "<tr>" + headers.map(h => "<th>" + h + "</th>").join("") + "</tr>";
            table.querySelector("tbody").innerHTML = data.map(row => "<tr>" +
                headers.map(h => "<td>" + fmt(row[h]) + "</td>").join("") + "</tr>").join("");
        }

        function switchTab(tab, ev){
            document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
            document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));
            document.getElementById(tab).classList.add("active");
            if(ev && ev.target) ev.target.classList.add("active");
        }

        function loadDivisions(){
            const divs = new Set();
            (allData.credit || []).forEach(r => {
                if(r.Division && r.Division !== "Grand Total") divs.add(r.Division);
            });
            const sel = document.getElementById("divisionFilter");
            sel.innerHTML = "<option value='All'>All Divisions</option>";
            Array.from(divs).sort().forEach(d => {
                const o = document.createElement("option");
                o.value = d;
                o.textContent = d;
                sel.appendChild(o);
            });
        }

        async function exportData(){
            const division = document.getElementById("divisionFilter").value;
            const type = document.getElementById("exportType").value;

            const res = await fetch("/api/export", {
                method: "POST",
                headers: {"Content-Type":"application/json"},
                body: JSON.stringify({division, type})
            });
            if(!res.ok){
                alert("Export failed");
                return;
            }
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = type + "_" + division + ".xlsx";
            document.body.appendChild(a);
            a.click();
            URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }

        async function loadData(){
            try{
                const res = await fetch("/api/data");
                allData = await res.json();

                renderTable("creditTable", allData.credit);
                renderTable("debitTable", allData.debit);
                renderTable("arbTable", allData.arbitration);
                renderTable("cmTable", allData.currentMonth);
                renderTable("compTable", allData.compensation);
                renderTable("prTable", allData.prApproval);

                loadDivisions();

                document.getElementById("loading").style.display = "none";
                document.getElementById("content").style.display = "block";
            }catch(e){
                document.getElementById("loading").innerHTML = "Error loading data";
            }
        }

        loadData();
    </script>
</body>
</html>"""

# =========================================================
# FASTAPI SETUP
# =========================================================
app = FastAPI(title="Warranty Dashboard - Direct")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"]
)

# =========================================================
# STARTUP: LOAD DATA + LOG FILE PATHS
# =========================================================
@app.on_event("startup")
def startup():
    print("=" * 90)
    print("WARRANTY DASHBOARD START")
    print(f"Environment: {'RENDER' if IS_RENDER else 'LOCAL'}")
    print(f"PORT: {PORT}")
    print(f"DATA_DIR: {DATA_DIR}")
    print("=" * 90)

    # Load data (NO LOGIN)
    WARRANTY_DATA["credit_df"], WARRANTY_DATA["debit_df"], WARRANTY_DATA["arbitration_df"], WARRANTY_DATA["source_df"] = process_warranty_data()
    WARRANTY_DATA["current_month_df"], WARRANTY_DATA["current_month_source_df"] = process_current_month_warranty()

    # IMPORTANT: this line must be WARRANTY_DATA only
    WARRANTY_DATA["compensation_df"], WARRANTY_DATA["compensation_source_df"] = process_compensation_claim()

    WARRANTY_DATA["pr_approval_df"], WARRANTY_DATA["pr_approval_source_df"] = process_pr_approval()



# Typo fix (safe)
try:
    WARRANTYY_DATA
except NameError:
    pass

# =========================================================
# ROUTES
# =========================================================
@app.get("/")
async def root():
    return HTMLResponse(DASHBOARD_HTML)


@app.get("/api/data")
async def api_data():
    def df_to_records(df):
        if df is None or df.empty:
            return []
        rec = df.to_dict("records")
        for r in rec:
            for k, v in list(r.items()):
                if pd.isna(v):
                    r[k] = 0
        return rec

    return {
        "credit": df_to_records(WARRANTY_DATA["credit_df"]),
        "debit": df_to_records(WARRANTY_DATA["debit_df"]),
        "arbitration": df_to_records(WARRANTY_DATA["arbitration_df"]),
        "currentMonth": df_to_records(WARRANTY_DATA["current_month_df"]),
        "compensation": df_to_records(WARRANTY_DATA["compensation_df"]),
        "prApproval": df_to_records(WARRANTY_DATA["pr_approval_df"]),
    }


@app.post("/api/export")
async def api_export(request: Request):
    data = await request.json()
    division = data.get("division", "All")
    export_type = data.get("type", "credit")

    if export_type == "credit":
        df = WARRANTY_DATA["credit_df"]
    elif export_type == "debit":
        df = WARRANTY_DATA["debit_df"]
    elif export_type == "arbitration":
        df = WARRANTY_DATA["arbitration_df"]
    elif export_type == "currentmonth":
        df = WARRANTY_DATA["current_month_df"]
    elif export_type == "compensation":
        df = WARRANTY_DATA["compensation_df"]
    else:
        df = WARRANTY_DATA["pr_approval_df"]

    if df is None or df.empty:
        return JSONResponse({"error": "No data"}, status_code=400)

    if division not in ("All", "Grand Total") and "Division" in df.columns:
        df_export = df[(df["Division"] == division) | (df["Division"] == "Grand Total")].copy()
    else:
        df_export = df.copy()

    wb = Workbook()
    ws = wb.active
    ws.title = export_type[:20]

    header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for col_idx, col_name in enumerate(df_export.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font

    for r_idx, row in enumerate(df_export.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    filename = f"{export_type}_{division}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return StreamingResponse(
        iter([out.getvalue()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
