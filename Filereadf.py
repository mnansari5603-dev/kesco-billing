import os
import glob
import pandas as pd

# Safe Path configuration
current_dir = os.path.dirname(os.path.abspath(__file__))
downloads_path = os.path.dirname(current_dir) # Downloads folder
folder_path = os.path.join(downloads_path, "total")

print("--- Final Inside-File Structured Matching Start ---")

# 1. 'total' folder ki saari files ke ANDAR se meter numbers scan karna
if not os.path.exists(folder_path):
    print(f"Error: '{folder_path}' folder nahi mila!")
    exit()

print("1. 'total' folder ki saari files ke andar se structured data read ho raha hai...")
meter_to_file_map = {}

all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

for file_path in all_files:
    filename = os.path.basename(file_path)
    try:
        # File ka delimiter auto-detect karne ke liye sep=None aur engine='python' lagaya hai
        # Kyunki KE01... files me data separator comma, tab, ya pipe ho sakta hai
        sub_df = pd.read_csv(file_path, sep=None, engine='python', low_memory=False)
        
        # Column names ke aage-piche se space hatana aur uppercase karna
        sub_df.columns = sub_df.columns.str.strip().str.upper()
        
        # Hamare target columns check karna
        target_col = None
        if "SERIAL_NBR" in sub_df.columns:
            target_col = "SERIAL_NBR"
        elif "BADGE_NBR" in sub_df.columns:
            target_col = "BADGE_NBR"
            
        if target_col:
            # Uss file ke saare meter numbers nikal kar clean karna
            meters_in_file = sub_df[target_col].dropna().astype(str).str.strip().str.upper()
            
            for m in meters_in_file:
                # Agar float format me ho (jaise 12345.0) toh .0 hatane ke liye
                clean_m = m.split('.')[0].strip()
                if clean_m and clean_m != 'NAN':
                    # Map me save karein
                    meter_to_file_map[clean_m] = filename
                    
                    # Prefix variations handle karna (GE ke saath aur bina GE ke)
                    if not clean_m.startswith("GE") and clean_m.isdigit():
                        meter_to_file_map["GE" + clean_m] = filename
                    elif clean_m.startswith("GE"):
                        meter_to_file_map[clean_m.replace("GE", "")] = filename
    except Exception as e:
        # Agar koi bina structured header wali file ho toh skip karein
        continue

print(f"Successfully scanned! Total unique meters mapped from folder: {len(meter_to_file_map)}")

if len(meter_to_file_map) == 0:
    print("Error: Kisi bhi file ke andar 'SERIAL_NBR' ya 'BADGE_NBR' column nahi mila ya load nahi hua!")
    exit()

# 2. Main CSV file ko auto-detect karna
csv_files = glob.glob(os.path.join(downloads_path, "*.csv"))
csv_path = None

for file in csv_files:
    try:
        temp_df = pd.read_csv(file, nrows=1)
        temp_df.columns = temp_df.columns.str.strip()
        if "New Meter Serial No" in temp_df.columns:
            csv_path = file
            break
    except:
        continue

if csv_path is None:
    print("Error: Downloads folder me 'New Meter Serial No' column wali koi CSV file nahi mili!")
    exit()

print(f"Sahi file mil gayi hai: {os.path.basename(csv_path)}")

# 3. Main CSV file load karna
print("2. Main CSV file load ho rahi hai (1.5 Lakh+ Rows)...")
df = pd.read_csv(csv_path)
df.columns = df.columns.str.strip()

# Main file ke column ko clean text banana
df["New Meter Serial No"] = df["New Meter Serial No"].astype(str).str.strip().str.upper()

# 4. Matching Function
def find_file(meter_no):
    if pd.isna(meter_no) or meter_no == "NAN":
        return ""
    clean_meter = meter_no.split('.')[0].strip()
    return meter_to_file_map.get(clean_meter, "Not Found")

print("3. Matching process chal rahi hai...")
target_col = [col for col in df.columns if "FILE NAME" in col.upper()]
col_to_write = target_col[0] if target_col else "FILE NAME REQUIRED"

df[col_to_write] = df["New Meter Serial No"].apply(find_file)

# 5. Result Excel me save karna
output_path = os.path.join(downloads_path, "FILE_NAME_FINAL_RESULT.xlsx")
print("4. Final Data Excel sheet me save ho raha hai...")
df.to_excel(output_path, index=False)

print("\n--- Ekdum Sahi Kaam Poora Hua! ---")
print(f"Aapki final file yahan save ho gayi hai:\n{output_path}")