import pandas as pd
import os
import re

def append_daily_report_to_master():
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    daily_file_csv = os.path.join(downloads_path, "WC Consumer Installation (10).csv")
    daily_file_xlsx = os.path.join(downloads_path, "WC Consumer Installation (10).xlsx")
    master_file = os.path.join(downloads_path, "L2 APROVED TILL DATE.xlsx")
    
    print("\n--- 🚀 MASTER WC SHEET AUTOMATION SHURU ---")
    
    if os.path.exists(daily_file_csv):
        daily_file = daily_file_csv
    elif os.path.exists(daily_file_xlsx):
        daily_file = daily_file_xlsx
    else:
        print(f"❌ ERROR: Daily file nahi mili!")
        return
        
    if not os.path.exists(master_file):
        print(f"❌ ERROR: Master file nahi mili!")
        return

    print("\n👉 Kis date ka data nikal kar Master file me New Sheet jodni hai?")
    user_input = input("📅 Date Type Karein aur ENTER dabayein: ").strip()
    
    # 💡 SMART FIX: Agar aap 'L2 Approval Date 18-05-2026' bhi likh dein, 
    # toh yeh usme se sirf '18-05-2026' nikal lega automatically
    date_match = re.search(r'\d{2}-\d{2}-\d{4}', user_input)
    if date_match:
        target_date = date_match.group(0)
    else:
        target_date = user_input # Backup agar normal enter kiya ho

    sheet_name = target_date

    print(f"\n🔄 Step 1: Daily data file load ho rahi hai...")
    try:
        if daily_file.endswith('.csv'):
            df_daily = pd.read_csv(daily_file, low_memory=False)
        else:
            df_daily = pd.read_excel(daily_file, engine='openpyxl')
            
        print(f"✅ Daily File Load Ho Gayi! Total Rows: {len(df_daily)}")
        df_daily.columns = [str(col).strip() for col in df_daily.columns]
        
        required_cols = [
            'Consumer Connection Type', 
            'Scan New Meter Serial No', 
            'L2 Approval Status', 
            'L2 Approval Date'
        ]

        print(f"🔍 Step 2: Date ({target_date}) ke hisab se data filter ho raha hai...")
        df_daily['L2_Status_Clean'] = df_daily['L2 Approval Status'].astype(str).str.strip().str.upper()
        
        filtered_df = df_daily[
            (df_daily['L2 Approval Date'].astype(str).str.contains(target_date, na=False)) & 
            (df_daily['L2_Status_Clean'].str.contains('APPROV', na=False))
        ]
        
        final_df = filtered_df[required_cols].copy()
        total_rows = len(final_df)
        print(f"📊 Filtering Complete! Total {total_rows} rows mili hain.")
        
        if total_rows == 0:
            print(f"⚠️ WARNING: Date '{target_date}' ka koi data nahi mila. Ek baar shuruat ki dates check karein:")
            print(df_daily['L2 Approval Date'].dropna().unique()[:5])
            return
            
        print(f"💾 Step 3: Master File me New Sheet '{sheet_name}' add ho rahi hai...")
        with pd.ExcelWriter(master_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
        print(f"\n🎉 SUCCESSFUL!!! Master file update ho gayi hai.")
        print(f"👉 Nayi Sheet ka naam: '{sheet_name}' (Total Rows Added: {total_rows})\n")
        
    except Exception as e:
        print(f"❌ KUCH GADBAD HUI: {e}")

if __name__ == "__main__":
    append_daily_report_to_master()