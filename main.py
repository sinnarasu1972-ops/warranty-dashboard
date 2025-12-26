import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import os
import socket
from typing import Optional
import sys
from functools import lru_cache
import hashlib
import secrets
import string
from PIL import Image, ImageDraw, ImageFont
import io
import base64
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# ==================== WARRANTY DATA PROCESSING ====================

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

def process_pr_approval():
    """Process PR Approval data and return summary dataframe"""
    input_path = r"D:\Power BI New\warranty dashboard render\Pr_Approval_Claims_Merged.xlsx"
    
    try:
        df = pd.read_excel(input_path)
        print("  PR Approval data loaded successfully")
        print(f"  Available columns: {df.columns.tolist()[:10]}...")
        print(f"  Total rows in source data: {len(df)}")

        summary_columns = [
            'Division', 'PA Request No.', 'PA Date', 'Request Type', 'App. Claim Amt from M&M'
        ]
        
        available_summary_columns = [col for col in summary_columns if col in df.columns]
        missing_columns = [col for col in summary_columns if col not in df.columns]
        
        if missing_columns:
            print(f" Missing columns in PR Approval: {missing_columns}")
        
        if not available_summary_columns:
            print(f" No required columns found in PR Approval file")
            return None, None

        df_summary_display = df[available_summary_columns].copy()
        
        if 'Division' in df_summary_display.columns:
            df_summary_display['Division'] = df_summary_display['Division'].astype(str).str.strip()
            df_summary_display = df_summary_display[df_summary_display['Division'].notna() & 
                                                      (df_summary_display['Division'] != '') & 
                                                      (df_summary_display['Division'] != 'nan')]
        
        if 'App. Claim Amt from M&M' in df_summary_display.columns:
            df_summary_display['App. Claim Amt from M&M'] = pd.to_numeric(
                df_summary_display['App. Claim Amt from M&M'], errors='coerce'
            ).fillna(0)
        
        summary_data = []
        
        if 'Division' in df_summary_display.columns:
            for division in sorted(df_summary_display['Division'].unique()):
                div_data = df_summary_display[df_summary_display['Division'] == division]
                
                summary_row = {'Division': division}
                summary_row['Total Requests'] = len(div_data)
                
                if 'App. Claim Amt from M&M' in df_summary_display.columns:
                    summary_row['Total Approved Amount'] = div_data['App. Claim Amt from M&M'].sum()
                
                if 'Request Type' in df_summary_display.columns:
                    request_types = div_data['Request Type'].value_counts().to_dict()
                    for req_type, count in request_types.items():
                        if pd.notna(req_type) and str(req_type).strip() != '':
                            summary_row[f'{req_type} Count'] = count
                
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            
            grand_total = {'Division': 'Grand Total'}
            for col in summary_df.columns:
                if col != 'Division':
                    if summary_df[col].dtype in ['int64', 'float64']:
                        grand_total[col] = summary_df[col].sum()
            
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        print("\n✓ PR Approval processing completed successfully")
        if not summary_df.empty:
            print(f"  Total Requests: {len(df_summary_display)}")
            if 'App. Claim Amt from M&M' in df_summary_display.columns:
                print(f"  Total Approved Amount: {df_summary_display['App. Claim Amt from M&M'].sum():,.2f}")
        
        return summary_df, df

    except FileNotFoundError:
        print(f" PR Approval file not found: {input_path}")
        return None, None
    except Exception as e:
        import traceback
        print(f" Error processing PR Approval data: {e}")
        traceback.print_exc()
        return None, None

def process_compensation_claim():
    """Process compensation claim data and return summary dataframe"""
    input_path = r"D:\Power BI New\Warranty Debit\Transit_Claims_Merged.xlsx"
    
    try:
        df = pd.read_excel(input_path)
        print("✓ Compensation Claim data loaded successfully")
        print(f"  Available columns: {df.columns.tolist()[:10]}...")
        print(f"  Total rows in source data: {len(df)}")

        required_columns = [
            'Division', 'RO Id.', 'Registration No.', 'RO Date', 'RO Bill Date',
            'Chassis No.', 'Model Group', 'Claim Amount', 'Request Status',
            'Claim Approved Amt.', 'No. of Days'
        ]
        
        available_columns = [col for col in required_columns if col in df.columns]
        
        if not available_columns:
            print(f" No required columns found in Compensation Claim file")
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
                
                summary_row = {'Division': division}
                summary_row['Total Claims'] = len(div_data)
                
                if 'Claim Amount' in df_filtered.columns:
                    summary_row['Total Claim Amount'] = div_data['Claim Amount'].sum()
                
                if 'Claim Approved Amt.' in df_filtered.columns:
                    summary_row['Total Approved Amount'] = div_data['Claim Approved Amt.'].sum()
                
                if 'No. of Days' in df_filtered.columns:
                    summary_row['Avg No. of Days'] = div_data['No. of Days'].mean()
                
                summary_data.append(summary_row)
            
            summary_df = pd.DataFrame(summary_data)
            
            grand_total = {'Division': 'Grand Total'}
            
            if 'Total Claims' in summary_df.columns:
                grand_total['Total Claims'] = summary_df['Total Claims'].sum()
            
            if 'Total Claim Amount' in summary_df.columns:
                grand_total['Total Claim Amount'] = summary_df['Total Claim Amount'].sum()
            
            if 'Total Approved Amount' in summary_df.columns:
                grand_total['Total Approved Amount'] = summary_df['Total Approved Amount'].sum()
            
            if 'Avg No. of Days' in summary_df.columns:
                grand_total['Avg No. of Days'] = summary_df['Avg No. of Days'].mean()
            
            summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
        else:
            summary_df = pd.DataFrame()

        print("\n✓ Compensation Claim processing completed successfully")
        if not summary_df.empty:
            print(f"  Total Claims: {len(df_filtered)}")
            if 'Claim Amount' in df_filtered.columns:
                print(f"  Total Claim Amount: {df_filtered['Claim Amount'].sum():,.2f}")
        
        return summary_df, df_filtered

    except FileNotFoundError:
        print(f" Compensation Claim file not found: {input_path}")
        return None, None
    except Exception as e:
        import traceback
        print(f" Error processing compensation claim data: {e}")
        traceback.print_exc()
        return None, None

def process_current_month_warranty():
    """Process current month warranty data and return summary dataframe"""
    input_path = r"D:\Power BI New\Warranty Debit\Pending Warranty Claim Details.xlsx"
    
    try:
        df = pd.read_excel(input_path, sheet_name='Pending Warranty Claim Details')
        print("✓ Current Month Warranty data loaded successfully")
        print(f"  Available columns: {df.columns.tolist()[:10]}...")
        print(f"  Total rows in source data: {len(df)}")

        required_columns = ['Division', 'Pending Claims Spares', 'Pending Claims Labour']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f" Missing columns in Current Month Warranty: {missing_columns}")
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

        print("\n✓ Current Month Warranty processing completed successfully")
        print(f"  Total Pending Claims Spares: {grand_total['Pending Claims Spares Count']}")
        print(f"  Total Pending Claims Labour: {grand_total['Pending Claims Labour Count']}")
        
        return summary_df, df

    except FileNotFoundError:
        print(f" Current Month Warranty file not found: {input_path}")
        return None, None
    except Exception as e:
        import traceback
        print(f" Error processing current month warranty data: {e}")
        traceback.print_exc()
        return None, None

def process_warranty_data():
    """Process warranty data and return credit, debit, and arbitration dataframes"""
    input_path = r"D:\Power BI New\Warranty Debit\Warranty Debit.xlsx"
    
    try:
        df = pd.read_excel(input_path, sheet_name='Sheet1')
        print("✓ Warranty data loaded successfully")
        print(f"  Available columns: {df.columns.tolist()[:5]}...")
        print(f"  Total rows in source data: {len(df)}")

        dealer_mapping = {
            'AMRAVATI': 'AMT',
            'CHAUFULA_SZZ': 'CHA',
            'CHIKHALI': 'CHI',
            'KOLHAPUR_WS': 'KOL',
            'NAGPUR_KAMPTHEE ROAD': 'HO',
            'NAGPUR_WARDHAMAN NGR': 'CITY',
            'SHIKRAPUR_SZS': 'SHI',
            'WAGHOLI': 'WAG',
            'YAVATMAL': 'YAT',
            'NAGPUR_WARDHAMAN NGR_CQ': 'CQ'
        }

        numeric_columns = ['Total Claim Amount', 'Credit Note Amount', 'Debit Note Amount']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        print(f"\n  Summary:")
        print(f"    Total Credit Note: {df['Credit Note Amount'].sum():,.2f}")
        print(f"    Total Debit Note: {df['Debit Note Amount'].sum():,.2f}")

        df['Dealer_Code'] = df['Dealer Location'].map(dealer_mapping).fillna(df['Dealer Location'])
        df['Month'] = df['Fiscal Month'].astype(str).str.strip().str[:3]
        df['Claim arbitration ID'] = df['Claim arbitration ID'].astype(str).replace('nan', '').replace('', np.nan)

        dealers = sorted(df['Dealer_Code'].unique())
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov']

        # 1. CREDIT NOTE TABLE
        credit_df = pd.DataFrame({'Division': dealers})
        print("\n  Processing Credit Note Amounts...")
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Credit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Credit Note {month}']
                credit_df = credit_df.merge(summary, on='Division', how='left')
                print(f"    {month}: {month_data['Credit Note Amount'].sum():,.2f}")
            else:
                credit_df[f'Credit Note {month}'] = 0
        
        credit_df = credit_df.fillna(0)
        credit_columns = [f'Credit Note {month}' for month in months]
        credit_df['Total Credit'] = credit_df[credit_columns].sum(axis=1)
        
        grand_total_credit = {'Division': 'Grand Total'}
        for col in credit_df.columns[1:]:
            grand_total_credit[col] = credit_df[col].sum()
        credit_df = pd.concat([credit_df, pd.DataFrame([grand_total_credit])], ignore_index=True)

        # 2. DEBIT NOTE TABLE
        debit_df = pd.DataFrame({'Division': dealers})
        print("\n  Processing Debit Note Amounts...")
        for month in months:
            month_data = df[df['Month'] == month]
            if not month_data.empty:
                summary = month_data.groupby('Dealer_Code')['Debit Note Amount'].sum().reset_index()
                summary.columns = ['Division', f'Debit Note {month}']
                debit_df = debit_df.merge(summary, on='Division', how='left')
                print(f"    {month}: {month_data['Debit Note Amount'].sum():,.2f}")
            else:
                debit_df[f'Debit Note {month}'] = 0
        
        debit_df = debit_df.fillna(0)
        debit_columns = [f'Debit Note {month}' for month in months]
        debit_df['Total Debit'] = debit_df[debit_columns].sum(axis=1)
        
        grand_total_debit = {'Division': 'Grand Total'}
        for col in debit_df.columns[1:]:
            grand_total_debit[col] = debit_df[col].sum()
        debit_df = pd.concat([debit_df, pd.DataFrame([grand_total_debit])], ignore_index=True)

        # 3. CLAIM ARBITRATION TABLE
        arbitration_df = pd.DataFrame({'Division': dealers})
        print("\n  Processing Claim Arbitration...")
        
        def is_arbitration(value):
            if pd.isna(value): return False
            value = str(value).strip().upper()
            return value.startswith('ARB') and value != 'NAN'

        for month in months:
            month_data = df[df['Month'] == month].copy()
            month_data['Is_ARB'] = month_data['Claim arbitration ID'].apply(is_arbitration)
            month_data['Arbitration_Amount'] = month_data.apply(
                lambda row: row['Debit Note Amount'] if row['Is_ARB'] else 0,
                axis=1
            )
            arb_summary = month_data.groupby('Dealer_Code')['Arbitration_Amount'].sum().reset_index()
            arb_summary.columns = ['Division', f'Claim Arbitration {month}']
            arbitration_df = arbitration_df.merge(arb_summary, on='Division', how='left')
            print(f"    {month}: {month_data['Arbitration_Amount'].sum():,.2f}")
        
        arbitration_df = arbitration_df.fillna(0)
        
        arbitration_cols = [f'Claim Arbitration {m}' for m in months]
        total_debit_by_dealer = debit_df[debit_df['Division'] != 'Grand Total'][['Division', 'Total Debit']].copy()
        arbitration_df = arbitration_df.merge(total_debit_by_dealer, on='Division', how='left')
        
        arbitration_df['Pending Claim Arbitration'] = (
            arbitration_df['Total Debit'] - arbitration_df[arbitration_cols].sum(axis=1)
        )
        
        arbitration_df = arbitration_df.drop('Total Debit', axis=1)
        
        grand_total_arb = {'Division': 'Grand Total'}
        for col in arbitration_df.columns[1:]:
            grand_total_arb[col] = arbitration_df[col].sum()
        arbitration_df = pd.concat([arbitration_df, pd.DataFrame([grand_total_arb])], ignore_index=True)

        print("\n✓ Warranty data processing completed successfully")
        return credit_df, debit_df, arbitration_df, df

    except FileNotFoundError:
        print(f" Warranty file not found: {input_path}")
        return None, None, None, None
    except Exception as e:
        import traceback
        print(f" Error processing warranty data: {e}")
        traceback.print_exc()
        return None, None, None, None

# ==================== IMAGE HANDLING ====================

def get_mahindra_images():
    """Load Mahindra vehicle images from the folder"""
    image_folder = r"D:\Power BI New\Warranty Debit\Image"
    images = []
    branding_images = []
    vehicle_images = []
    
    if os.path.exists(image_folder):
        try:
            for file in os.listdir(image_folder):
                if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    image_path = os.path.join(image_folder, file)
                    try:
                        with open(image_path, 'rb') as img_file:
                            img_data = base64.b64encode(img_file.read()).decode()
                            img_dict = {
                                'name': file,
                                'data': img_data,
                                'path': image_path
                            }
                            
                            file_lower = file.lower()
                            if 'mahindra' in file_lower or 'logo' in file_lower or 'hero' in file_lower:
                                branding_images.append(img_dict)
                                print(f"  ✓ Loaded Branding: {file}")
                            else:
                                vehicle_images.append(img_dict)
                                print(f"  ✓ Loaded Vehicle: {file}")
                    except Exception as e:
                        print(f"   Could not load {file}: {e}")
        except Exception as e:
            print(f" Error reading image folder: {e}")
    else:
        print(f" Image folder not found: {image_folder}")
    
    images = branding_images + vehicle_images
    return images

print("Loading Mahindra vehicle images...")
MAHINDRA_IMAGES = get_mahindra_images()
print(f" Loaded {len(MAHINDRA_IMAGES)} vehicle images\n")

# ==================== DASHBOARD HTML (NO LOGIN) ====================

DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Unnati Warranty Management Dashboard</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f5f5 0%, #e0e0e0 100%);
            padding: 0;
            margin: 0;
            min-height: 100vh;
        }
        
        .navbar {
            background: linear-gradient(135deg, #FF8C00 0%, #FF6B35 100%);
            color: white;
            padding: 15px 0;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
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
            font-size: 24px;
            font-weight: 700;
        }
        
        .container {
            max-width: 1400px;
            margin: 30px auto;
            padding: 0 20px;
        }
        
        h1 {
            color: #333;
            margin-bottom: 30px;
            text-align: center;
            font-weight: 700;
        }
        
        .dashboard-content {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            padding: 30px;
        }
        
        .nav-tabs {
            border-bottom: 2px solid #FF8C00;
            margin-bottom: 30px;
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        .nav-tabs .nav-link {
            color: #666;
            font-weight: 600;
            border: none;
            border-bottom: 3px solid transparent;
            padding: 12px 20px;
            cursor: pointer;
            transition: all 0.3s ease;
            background: none;
        }
        
        .nav-tabs .nav-link:hover {
            color: #FF8C00;
            border-bottom-color: #FF8C00;
        }
        
        .nav-tabs .nav-link.active {
            color: #FF8C00;
            border-bottom-color: #FF8C00;
        }
        
        .tab-content {
            display: none;
        }
        
        .tab-content.active {
            display: block;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 12px;
            overflow-x: auto;
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
        
        .data-table tbody tr:hover {
            background: #f9f9f9;
        }
        
        .data-table tbody tr:last-child {
            background: #fff8f3;
            font-weight: 700;
            border-top: 2px solid #FF8C00;
            border-bottom: 2px solid #FF8C00;
        }
        
        .data-table tbody tr:last-child td {
            color: #FF8C00;
        }
        
        .loading-spinner {
            display: none;
            text-align: center;
            padding: 40px;
        }
        
        .spinner {
            border: 4px solid rgba(255, 140, 0, 0.2);
            border-top: 4px solid #FF8C00;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .table-title {
            font-size: 16px;
            font-weight: 700;
            color: #FF8C00;
            margin-bottom: 15px;
        }
        
        .table-wrapper {
            overflow-x: auto;
        }
        
        .export-section {
            margin: 30px 0;
            padding: 20px;
            background: linear-gradient(135deg, #fff8f3 0%, #ffe8d6 100%);
            border-radius: 8px;
            border-left: 5px solid #FF8C00;
            box-shadow: 0 2px 8px rgba(255, 140, 0, 0.1);
        }
        
        .export-section h3 {
            color: #FF8C00;
            margin-bottom: 15px;
            font-size: 16px;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .export-controls {
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
            background: white;
            padding: 15px;
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
            transition: all 0.3s ease;
            min-width: 150px;
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
            transition: all 0.3s ease;
        }
        
        .export-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(76, 175, 80, 0.3);
        }
        
        .export-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }
    </style>
</head>
<body>
    <nav class="navbar navbar-dark">
        <div class="container-fluid">
            <span class="navbar-brand">📊 Unnati Motors Warranty Management Dashboard</span>
        </div>
    </nav>
    
    <div class="container">
        <div class="dashboard-content">
            <div class="loading-spinner" id="loadingSpinner">
                <div class="spinner"></div>
                <p style="margin-top: 15px; color: #666;">Loading warranty data...</p>
            </div>
            
            <div id="warrantyTabs" style="display: none;">
                <!-- Tab Navigation -->
                <div class="nav-tabs">
                    <button class="nav-link active" onclick="switchTab('credit')">💰 Warranty Credit</button>
                    <button class="nav-link" onclick="switchTab('debit')">💳 Warranty Debit</button>
                    <button class="nav-link" onclick="switchTab('arbitration')">⚖️ Claim Arbitration</button>
                    <button class="nav-link" onclick="switchTab('currentmonth')">📅 Current Month Warranty</button>
                    <button class="nav-link" onclick="switchTab('compensation')">🚚 Compensation Claim</button>
                    <button class="nav-link" onclick="switchTab('pr_approval')">✅ PR Approval</button>
                </div>

                <!-- EXPORT SECTION -->
                <div class="export-section">
                    <h3>📥 Export to Excel</h3>
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
                        
                        <button onclick="exportToExcel()" class="export-btn" id="exportBtn">
                            📥 Export to Excel
                        </button>
                    </div>
                </div>
                
                <!-- Credit Note Tab -->
                <div id="credit" class="tab-content active">
                    <div class="table-title">Warranty Credit Note by Division & Month</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="creditTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Debit Note Tab -->
                <div id="debit" class="tab-content">
                    <div class="table-title">Warranty Debit Note by Division & Month</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="debitTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Claim Arbitration Tab -->
                <div id="arbitration" class="tab-content">
                    <div class="table-title">Warranty Claim Arbitration by Division</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="arbitrationTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Current Month Warranty Tab -->
                <div id="currentmonth" class="tab-content">
                    <div class="table-title">Current Month Warranty - Pending Claims Summary</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="currentMonthTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Compensation Claim Tab -->
                <div id="compensation" class="tab-content">
                    <div class="table-title">Compensation Claim - Transit Claims Summary</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="compensationTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                
                <!-- PR Approval Tab -->
                <div id="pr_approval" class="tab-content">
                    <div class="table-title">PR Approval - Claims Summary</div>
                    <div class="table-wrapper">
                        <table class="data-table" id="prApprovalTable">
                            <thead></thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        let warrantyData = {};
        
        async function loadDashboard() {
            const spinner = document.getElementById('loadingSpinner');
            const tabs = document.getElementById('warrantyTabs');
            
            spinner.style.display = 'block';
            tabs.style.display = 'none';
            
            try {
                const response = await fetch('/api/warranty-data', {
                    method: 'GET',
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
                
                if (!response.ok) {
                    throw new Error('Failed to load warranty data');
                }
                
                warrantyData = await response.json();
                
                displayCreditTable(warrantyData.credit);
                displayDebitTable(warrantyData.debit);
                displayArbitrationTable(warrantyData.arbitration);
                displayCurrentMonthTable(warrantyData.currentMonth);
                displayCompensationTable(warrantyData.compensation);
                displayPrApprovalTable(warrantyData.prApproval);
                
                loadDivisions();
                
                spinner.style.display = 'none';
                tabs.style.display = 'block';
            } catch (error) {
                console.error('Error loading dashboard:', error);
                spinner.innerHTML = '<p style="color: red; padding: 20px;">Error loading warranty data</p>';
            }
        }
        
        function displayCreditTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('creditTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 0}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function displayDebitTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('debitTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 0}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function displayArbitrationTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('arbitrationTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 0}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function displayCurrentMonthTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('currentMonthTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 0}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function displayCompensationTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('compensationTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
                return '<tr>' + headers.map((h) => {
                    const value = typeof row[h] === 'number' ? row[h].toLocaleString('en-IN', {maximumFractionDigits: 2}) : row[h];
                    return '<td>' + value + '</td>';
                }).join('') + '</tr>';
            }).join('');
        }
        
        function displayPrApprovalTable(data) {
            if (!data || data.length === 0) return;
            const table = document.getElementById('prApprovalTable');
            const headers = Object.keys(data[0]);
            const headerRow = table.querySelector('thead');
            headerRow.innerHTML = headers.map(h => '<th>' + h + '</th>').join('');
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = data.map((row) => {
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
            const currentType = document.getElementById('exportType').value;
            let dataSource = warrantyData.credit;
            
            if (currentType === 'debit') dataSource = warrantyData.debit;
            if (currentType === 'arbitration') dataSource = warrantyData.arbitration;
            if (currentType === 'currentmonth') dataSource = warrantyData.currentMonth;
            if (currentType === 'compensation') dataSource = warrantyData.compensation;
            if (currentType === 'pr_approval') dataSource = warrantyData.prApproval;
            
            if (dataSource && dataSource.length > 0) {
                dataSource.forEach(row => {
                    if (row.Division && row.Division !== 'Grand Total') {
                        divisions.add(row.Division);
                    }
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

        document.getElementById('exportType')?.addEventListener('change', loadDivisions);

        async function exportToExcel() {
            const division = document.getElementById('divisionFilter').value;
            const type = document.getElementById('exportType').value;
            const exportBtn = document.getElementById('exportBtn');
            
            if (!division) {
                alert('Please select a division');
                return;
            }
            
            exportBtn.disabled = true;
            exportBtn.textContent = '⏳ Exporting...';
            
            try {
                const response = await fetch('/api/export-to-excel', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({division: division, type: type})
                });
                
                if (!response.ok) throw new Error('Export failed');
                
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${type}_${division}_${new Date().toISOString().split('T')[0]}.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                alert('Export completed successfully!');
            } catch (error) {
                alert('Export failed: ' + error.message);
            } finally {
                exportBtn.disabled = false;
                exportBtn.textContent = '📥 Export to Excel';
            }
        }
        
        window.onload = function() {
            loadDashboard();
        };
    </script>
</body>
</html>
"""

# ==================== FASTAPI SETUP ====================

app = FastAPI()

# Add CORS middleware for Render
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==================== API ENDPOINTS ====================

@app.post("/api/export-to-excel")
async def export_to_excel(request: Request):
    """Export selected division data to Excel"""
    try:
        body = await request.json()
        selected_division = body.get('division', 'All')
        export_type = body.get('type', 'credit')
        
        if export_type not in ['credit', 'debit', 'arbitration', 'currentmonth', 'compensation', 'pr_approval']:
            raise HTTPException(status_code=400, detail="Invalid export type")
        
        if export_type == 'currentmonth':
            return await export_current_month_warranty(selected_division)
        if export_type == 'compensation':
            return await export_compensation_claim(selected_division)
        if export_type == 'pr_approval':
            return await export_pr_approval(selected_division)
        
        if export_type == 'credit':
            df = WARRANTY_DATA['credit_df']
        elif export_type == 'debit':
            df = WARRANTY_DATA['debit_df']
        else:
            df = WARRANTY_DATA['arbitration_df']
        
        if df is None or df.empty:
            raise HTTPException(status_code=500, detail="No data available for export")
        
        if selected_division != 'All' and selected_division != 'Grand Total':
            df_export = df[df['Division'] == selected_division].copy()
            grand_total_row = df[df['Division'] == 'Grand Total']
            if not grand_total_row.empty:
                df_export = pd.concat([df_export, grand_total_row], ignore_index=True)
        else:
            df_export = df.copy()
        
        wb = Workbook()
        header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        ws1 = wb.active
        ws1.title = export_type.capitalize()
        
        for col_idx, column in enumerate(df_export.columns, 1):
            cell = ws1.cell(row=1, column=col_idx, value=column)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = ws1.cell(row=row_idx, column=col_idx)
                if isinstance(value, (int, float)):
                    cell.value = value
                    cell.number_format = '#,##0.00'
                else:
                    cell.value = str(value)
                cell.border = border
        
        for col_idx, column in enumerate(df_export.columns, 1):
            max_length = min(max(df_export[column].astype(str).map(len).max(), len(str(column))) + 2, 30)
            ws1.column_dimensions[chr(64 + col_idx)].width = max_length
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{selected_division}_{export_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        print(f"Export error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

async def export_current_month_warranty(selected_division: str):
    """Export Current Month Warranty data"""
    try:
        summary_df = WARRANTY_DATA['current_month_df']
        if summary_df is None or summary_df.empty:
            raise HTTPException(status_code=500, detail="No data available")
        
        if selected_division != 'All':
            df_export = summary_df[summary_df['Division'] == selected_division].copy()
        else:
            df_export = summary_df.copy()
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Summary"
        
        header_fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        
        for col_idx, column in enumerate(df_export.columns, 1):
            cell = ws1.cell(row=1, column=col_idx, value=column)
            cell.fill = header_fill
            cell.font = header_font
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                ws1.cell(row=row_idx, column=col_idx, value=value)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"{selected_division}_CurrentMonthWarranty_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

async def export_compensation_claim(selected_division: str):
    """Export Compensation Claim data"""
    try:
        summary_df = WARRANTY_DATA['compensation_df']
        if summary_df is None or summary_df.empty:
            raise HTTPException(status_code=500, detail="No data available")
        
        if selected_division != 'All':
            df_export = summary_df[summary_df['Division'] == selected_division].copy()
        else:
            df_export = summary_df.copy()
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Summary"
        
        for col_idx, column in enumerate(df_export.columns, 1):
            ws1.cell(row=1, column=col_idx, value=column)
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                ws1.cell(row=row_idx, column=col_idx, value=value)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"{selected_division}_CompensationClaim_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

async def export_pr_approval(selected_division: str):
    """Export PR Approval data"""
    try:
        summary_df = WARRANTY_DATA['pr_approval_df']
        if summary_df is None or summary_df.empty:
            raise HTTPException(status_code=500, detail="No data available")
        
        if selected_division != 'All':
            df_export = summary_df[summary_df['Division'] == selected_division].copy()
        else:
            df_export = summary_df.copy()
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Summary"
        
        for col_idx, column in enumerate(df_export.columns, 1):
            ws1.cell(row=1, column=col_idx, value=column)
        
        for row_idx, row in enumerate(df_export.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                ws1.cell(row=row_idx, column=col_idx, value=value)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"{selected_division}_PrApproval_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/warranty-data")
async def get_warranty_data():
    """Get warranty data"""
    try:
        if WARRANTY_DATA['credit_df'] is None:
            return {
                "credit": [],
                "debit": [],
                "arbitration": [],
                "currentMonth": [],
                "compensation": [],
                "prApproval": []
            }
        
        credit_records = WARRANTY_DATA['credit_df'].to_dict('records')
        debit_records = WARRANTY_DATA['debit_df'].to_dict('records')
        arbitration_records = WARRANTY_DATA['arbitration_df'].to_dict('records')
        
        current_month_records = []
        if WARRANTY_DATA['current_month_df'] is not None:
            current_month_records = WARRANTY_DATA['current_month_df'].to_dict('records')
        
        compensation_records = []
        if WARRANTY_DATA['compensation_df'] is not None:
            compensation_records = WARRANTY_DATA['compensation_df'].to_dict('records')
        
        pr_approval_records = []
        if WARRANTY_DATA['pr_approval_df'] is not None:
            pr_approval_records = WARRANTY_DATA['pr_approval_df'].to_dict('records')
        
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
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    """Serve dashboard"""
    return HTMLResponse(content=DASHBOARD_HTML)

@app.get("/dashboard")
async def dashboard():
    """Serve dashboard"""
    return HTMLResponse(content=DASHBOARD_HTML)

# ==================== STARTUP ====================

print("\n" + "=" * 100)
print("STARTING WARRANTY MANAGEMENT SYSTEM")
print("=" * 100)

print("\nProcessing warranty data...")
WARRANTY_DATA['credit_df'], WARRANTY_DATA['debit_df'], WARRANTY_DATA['arbitration_df'], WARRANTY_DATA['source_df'] = process_warranty_data()

print("\nProcessing current month warranty data...")
WARRANTY_DATA['current_month_df'], WARRANTY_DATA['current_month_source_df'] = process_current_month_warranty()

print("\nProcessing compensation claim data...")
WARRANTY_DATA['compensation_df'], WARRANTY_DATA['compensation_source_df'] = process_compensation_claim()

print("\nProcessing PR Approval data...")
WARRANTY_DATA['pr_approval_df'], WARRANTY_DATA['pr_approval_source_df'] = process_pr_approval()

if __name__ == "__main__":
    # Get port from environment variable (for Render)
    port = int(os.getenv('PORT', 8001))
    
    print("\n" + "=" * 100)
    print(f"✅ SERVER READY - Warranty Dashboard")
    print("=" * 100)
    print(f"🌐 PORT: {port}")
    print(f"🌐 URL: http://localhost:{port}/")
    print("=" * 100 + "\n")
    
    uvicorn.run(app, host="0.0.0.0", port=port)
