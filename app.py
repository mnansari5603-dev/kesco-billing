import pandas as pd
import os
import time
import webbrowser

print("⏳ 1. Excel File dhoond rahe hain...")

USER_PROFILE = os.environ.get("USERPROFILE", "C:\\Users\\Admin")
path_normal = os.path.join(USER_PROFILE, "Desktop", "BILLING REPORTS", "04-06-2026", "FINAL SUMMARY.xlsx")
path_onedrive = os.path.join(USER_PROFILE, "OneDrive", "Desktop", "BILLING REPORTS", "04-06-2026", "FINAL SUMMARY.xlsx")

if os.path.exists(path_onedrive):
    EXCEL_PATH = path_onedrive
elif os.path.exists(path_normal):
    EXCEL_PATH = path_normal
else:
    EXCEL_PATH = None

if EXCEL_PATH is None:
    print("❌ Error: Desktop par file nahi mili!")
else:
    print(f"✅ File mil gayi: {EXCEL_PATH}")
    print("🚀 2. Calamine Engine se 'Summary' sheet load ho rahi hai...")
    
    try:
        # Excel ko load kiya
        df = pd.read_excel(EXCEL_PATH, sheet_name="Summary", engine="calamine")
        
        # --- CLEANING THE 'UNNAMED' HEADERS ---
        # Pehli row ko asli headers banane ke liye (kyunki Excel me top cell merged tha)
        as_on_date_title = "Kesco Billing Summary As on 05-06-2026"
        
        # Agar pehli row me actual column names hain (jaise Total MI, L1 DONE etc.)
        if len(df) > 0:
            actual_headers = df.iloc[0].tolist()
            # Pehli column ka naam sahi rakhna
            actual_headers[0] = "Billing Category / Status"
            df.columns = actual_headers
            df = df.iloc[1:]  # Pehli row ko data se hata diya kyunki woh ab header hai
            
        df = df.fillna('')
        total_rows = len(df)
        
        # Generatng Beautiful HTML Rows
        table_rows_html = ""
        for _, row in df.iterrows():
            table_rows_html += "<tr>"
            for val in row:
                # Agar row me data 'TOD' ya main header jaisa hai toh use bold highlighting dena
                if val in ['TOD', 'TOTAL', 'Total']:
                    table_rows_html += f"<td class='fw-bold bg-light'>{val}</td>"
                else:
                    table_rows_html += f"<td>{val}</td>"
            table_rows_html += "</tr>"

        # Dynamically creating columns for HTML table header
        header_cols_html = "".join([f"<th>{str(col).strip()}</th>" for col in df.columns])

        print("⏳ 3. Premium HTML Dashboard design ho raha hai...")
        
        # Ultra Premium Corporate UI Design
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Kesco Billing Live Dashboard</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
            <style>
                body {{ background-color: #f3f4f6; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 25px; }}
                .dashboard-header {{ background: linear-gradient(135deg, #0f2027, #203a43, #2c5364); color: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); margin-bottom: 25px; text-align: center; }}
                .main-title {{ font-size: 28px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin: 0; }}
                .sub-title {{ font-size: 14px; opacity: 0.8; margin-top: 5px; font-family: monospace; }}
                .kpi-card {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #0d6efd; transition: transform 0.2s; }}
                .kpi-card:hover {{ transform: translateY(-2px); }}
                .table-container {{ background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); overflow-x: auto; margin-top: 20px; }}
                .custom-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }}
                .custom-table th {{ background-color: #1e293b !important; color: #ffffff !important; font-weight: 600; padding: 14px 12px; text-align: left; position: sticky; top: 0; border: 1px solid #334155; text-transform: uppercase; font-size: 12px; letter-spacing: 0.5px; }}
                .custom-table td {{ padding: 12px; border-bottom: 1px solid #e2e8f0; color: #334155; min-width: 110px; }}
                .custom-table tbody tr:hover {{ background-color: #f8fafc; }}
                .badge-source {{ background-color: rgba(255,255,255,0.15); padding: 5px 10px; border-radius: 6px; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="container-fluid">
                <div class="dashboard-header">
                    <h1 class="main-title">🏢 {as_on_date_title}</h1>
                    <div class="sub-title">System Output: Live Connected Report View</div>
                </div>
                
                <div class="row g-3">
                    <div class="col-md-4">
                        <div class="kpi-card">
                            <span class="text-muted text-uppercase fw-bold" style="font-size: 11px; letter-spacing: 0.5px;">Total Tracked Rows</span>
                            <h2 class="text-primary fw-bold m-0 mt-1">{total_rows:,}</h2>
                        </div>
                    </div>
                    <div class="col-md-8">
                        <div class="kpi-card" style="border-left-color: #64748b;">
                            <span class="text-muted text-uppercase fw-bold" style="font-size: 11px; letter-spacing: 0.5px;">Data File Source Path</span>
                            <div class="mt-1"><code class="text-dark bg-light p-1 px-2 rounded" style="font-size: 13px;">{EXCEL_PATH}</code></div>
                        </div>
                    </div>
                </div>
                
                <div class="table-container">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h4 class="text-secondary fw-bold m-0" style="font-size: 18px;">📋 Summary Records Grid</h4>
                        <span class="badge bg-success px-3 py-2 rounded-pill">Status: Active</span>
                    </div>
                    <div style="max-height: 550px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px;">
                        <table class="custom-table">
                            <thead>
                                tr>{header_cols_html}</tr>
                            </thead>
                            <tbody>
                                {table_rows_html}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """
        
        output_html_path = os.path.join(USER_PROFILE, "Desktop", "BILLING REPORTS", "04-06-2026", "DASHBOARD.html")
        with open(output_html_path, "w", encoding="utf-8") as f:
            f.write(html_content)
            
        print("\n🎯 DESIGN UPDATED! Ek dum brand new premium dashboard tayyar hai.")
        webbrowser.open(f"file:///{output_html_path}")

    except Exception as e:
        print(f"❌ Error: {e}")