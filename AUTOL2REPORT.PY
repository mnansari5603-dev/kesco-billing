# -*- coding: utf-8 -*-
import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta

def run_automation():
    # --- 1. CONFIGURATION & FILE PATHS ---
    report1_path = r"c:\Users\Admin\Downloads\WC Consumer Installation (24).csv"
    report2_path = r"c:\Users\Admin\Downloads\L2 APROVED TILL DATE.xlsx"
    
    # 🟢 AUTOMATIC YESTERDAY DATE FORMULA:
    # Agar aaj 20-05-2026 hai, toh yeh automatic '19-05-2026' nikalega
    target_date_str = (datetime.now() - timedelta(days=1)).strftime('%d-%m-%Y')
    
    print("="*60)
    print(f"Automation Shuru Ho Raha Hai: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Targeting Yesterday's Data: {target_date_str}")
    print("="*60)
    
    if not os.path.exists(report1_path) or not os.path.exists(report2_path):
        print("Error: Files check karein, path galat hai ya file nahi mili.")
        return

    # Check write permission for the output workbook before we start processing
    try:
        with open(report2_path, 'a', encoding='utf-8'):
            pass
    except PermissionError:
        print(f"Error: '{report2_path}' currently open hai. Pehle Excel file band karo, phir script dobara chalayein.")
        return
    except Exception as e:
        print(f"Error: '{report2_path}' par write access check fail hua: {e}")
        return

    # --- 2. READ REPORT 1 (CSV FILE) ---
    print(f"Step 1: '{report1_path}' se data read kiya ja raha hai...")
    try:
        df_r1 = pd.read_csv(report1_path, low_memory=False)
    except Exception as e:
        print(f"Error: Report 1 read karne me dikkat aayi: {e}")
        return

    # Sirf yahi 5 columns nikalne hain
    required_cols = [
        'Consumer Connection Type', 
        'Consumer Number', 
        'Scan New Meter Serial No', 
        'L2 Approval Status', 
        'L2 Approval Date'
    ]

    for col in required_cols:
        if col not in df_r1.columns:
            print(f"Error: Required column '{col}' CSV data me nahi mila!")
            return

    # --- 3. FILTER DATA FOR 'APPROVED' & YESTERDAY'S DATE ---
    print(f"Step 2: '{target_date_str}' aur Status 'APPROVED' ka data filter ho raha hai...")
    
    status_match = df_r1['L2 Approval Status'].astype(str).str.strip().str.upper() == 'APPROVED'
    date_match = df_r1['L2 Approval Date'].astype(str).str.strip().str.contains(target_date_str, na=False)
    
    df_filtered = df_r1[status_match & date_match][required_cols].copy()
    total_rows = len(df_filtered)
    print(f"-> Total {total_rows} filtered rows mili hain.")

    # --- 4. WRITE TO NEW SHEET (NAMED AS YESTERDAY'S DATE) ---
    print(f"Step 3: '{report2_path}' me '{target_date_str}' naam ki sheet add ho rahi hai...")
    try:
        with pd.ExcelWriter(report2_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_filtered.to_excel(writer, sheet_name=target_date_str, index=False)
        print(f"-> Sheet '{target_date_str}' successfully add/replace ho gayi.")
    except PermissionError:
        print(f"Error: '{report2_path}' par write permission nahi hai. Pehle Excel file band karo, phir script dobara chalayein.")
        return
    except Exception as e:
        print(f"Error: Report 2 me sheet write karne me dikkat aayi: {e}")
        return

    # --- 5. CALCULATE COUNTS ---
    print(f"Step 4: ESC Counts aur WC Total calculate kiya ja raha hai...")
    conn_series = df_filtered['Consumer Connection Type'].astype(str).str.strip().str.upper()
    
    esc_count = int((conn_series == 'ESC').sum())
    wc_count = total_rows  # Total filtered rows hi aapka WC count hai
    mco_count = 0          # Requirement ke mutabik hamesha 0 rahega
    
    print(f"-> Calculated: ESC = {esc_count}, WC = {wc_count}, MCO = {mco_count}")

    # --- 6. UPDATE SUMMARY SHEET ---
    print(f"Step 5: '{report2_path}' ki 'summary' sheet update ho rahi hai...")
    try:
        wb = load_workbook(report2_path)
        
        if 'summary' not in wb.sheetnames:
            summary_sheet = wb.create_sheet('summary', 0)
            summary_sheet.append(['Date', 'ESC Count', 'WC Count', 'MCO Count'])
            print("-> 'summary' sheet created.")
        else:
            summary_sheet = wb['summary']
            summary_sheet['A1'] = 'Date'
            summary_sheet['B1'] = 'ESC Count'
            summary_sheet['C1'] = 'WC Count'
            summary_sheet['D1'] = 'MCO Count'
            print("-> 'summary' sheet found and header verified.")
            
        next_row = summary_sheet.max_row + 1
        
        # Summary rows values insertion
        summary_sheet[f'A{next_row}'] = target_date_str
        summary_sheet[f'B{next_row}'] = esc_count
        summary_sheet[f'C{next_row}'] = wc_count
        summary_sheet[f'D{next_row}'] = mco_count
        print(f"-> Writing summary row {next_row}: Date={target_date_str}, ESC={esc_count}, WC={wc_count}, MCO={mco_count}")
        
        # Formatting (Styling)
        from openpyxl.styles import Font, Alignment
        font_style = Font(name='Arial', size=11)
        align_center = Alignment(horizontal='center')
        
        for col in ['A', 'B', 'C', 'D']:
            summary_sheet[f'{col}{next_row}'].font = font_style
            summary_sheet[f'{col}{next_row}'].alignment = align_center
            if col in ['B', 'C', 'D']:
                summary_sheet[f'{col}{next_row}'].number_format = '#,##0'
        
        wb.save(report2_path)
        wb.close()
        print("-> Summary sheet successfully update ho gayi!")
        
    except Exception as e:
        print(f"Error: Summary sheet update karne me dikkat aayi: {e}")
        return

    print("="*60)
    print("SUCCESS: AUTOMATION PROCESS COMPLETE HO GAYA!")
    print("="*60)

if __name__ == "__main__":
    run_automation()