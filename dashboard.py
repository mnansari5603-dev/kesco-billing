import pandas as pd

summary_file = r"C:\Users\Admin\Downloads\KESCO\FINAL SUMMARY.xlsx.xlsx"

# Summary Sheet Read
df = pd.read_excel(
    summary_file,
    sheet_name="Summary",
    header=None
)

# Grand Billed Total Row Find
row = df[df[0].astype(str).str.contains(
    "Grand Billed Total",
    case=False,
    na=False
)]

if len(row) == 0:
    print("Grand Billed Total not found")
    exit()

r = row.iloc[0]

dashboard = pd.DataFrame([{
    "Total_MI": r[1],
    "L1_Done": r[2],
    "L1_Pending": r[3],
    "L2_Done": r[4],
    "L2_Pending": r[5],
    "MCO_Done": r[6],
    "MCO_Pending": r[7],
    "Sync_Done": r[9],
    "Billing_Sent": r[13]
}])

dashboard.to_excel(
    r"C:\KESCO\Dashboard_Data.xlsx",
    index=False
)

print("Dashboard Updated Successfully")