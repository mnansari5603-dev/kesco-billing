import pandas as pd
import os

def filter_wc_report():
    input_file = r"c:/Users/Admin/Downloads/WC Consumer Installation (10).csv"
    output_file = r"c:/Users/Admin/Downloads/L2_Approved_Filtered_Report.xlsx"
    
    print("\n--- 🚀 WC REPORT ULTIMATE AUTOMATION SHURU ---")
    
    if not os.path.exists(input_file):
        input_file = r"c:/Users/Admin/Downloads/WC Consumer Installation (10).xlsx"
    if not os.path.exists(input_file):
        print(f"❌ ERROR: File nahi mili!")
        return

    print(f"🔄 Step 1: File load ho rahi hai...")
    try:
        if input_file.endswith('.csv'):
            df = pd.read_csv(input_file, low_memory=False)
        else:
            df = pd.read_excel(input_file, engine='openpyxl')
            
        print(f"✅ Original File Load Ho Gayi! Total Rows: {len(df)}")
        df.columns = [str(col).strip() for col in df.columns]
        
        required_cols = [
            'Consumer Connection Type', 
            'Scan New Meter Serial No', 
            'L2 Approval Status', 
            'L2 Approval Date'
        ]

        print("🔍 Step 2: Dates aur Status filter ho raha hai...")
        
        # Data ko clean karein
        df['L2_Status_Clean'] = df['L2 Approval Status'].astype(str).str.strip().str.upper()
        
        # 💡 SOLUTION: Hum '16-05-2026' YA '15-05-2026' dono dhoondhenge (Jo bhi mil jaye)
        # Agar aapko sirf 16 ki jagah 15 ka data chahiye toh dono nikal aayega
        filtered_df = df[
            (df['L2 Approval Date'].astype(str).str.contains('16-05-2026|15-05-2026|2026-05-16|2026-05-15', na=False)) & 
            (df['L2_Status_Clean'].str.contains('APPROV', na=False))
        ]
        
        final_df = filtered_df[required_cols].copy()
        total_rows = len(final_df)
        print(f"📊 Filtering Complete! Total {total_rows} rows mili hain.")
        
        if total_rows == 0:
            print("\n⚠️ SYSTEM ALERT: Dono dates par 0 rows mili!")
            print("📋 Sheet me available kuch unique dates ye hain (Taaki aapko pata chal sake kaunsi date ka data hai):")
            # Top 10 unique entries dikhayega terminal par
            print(df['L2 Approval Date'].dropna().unique()[:10])
            return
            
        print("💾 Step 3: Excel file save ho rahi hai...")
        final_df.to_excel(output_file, index=False)
        print(f"\n🎉 SUCCESSFUL!!! Aapki filtered file downloads me taiyaar hai:\n👉 {output_file}\n")
        
    except Exception as e:
        print(f"❌ KUCH GADBAD HUI: {e}")

if __name__ == "__main__":
    filter_wc_report()
    