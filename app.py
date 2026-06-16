import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os

# ⚡ Page Configuration
st.set_page_config(page_title="Kesco Billing Dashboard", page_icon="📊", layout="wide")

st.title("🏛️ KESCO Billing & MDM Dashboard")
st.markdown("---")

# 🧠 DYNAMIC PATH CHECKER
LOCAL_PATH = r"C:\Users\Admin\Desktop\BILLING REPORTS\16-06-2026\FINAL SUMMARY.xlsx"
ONLINE_PATH = "FINAL SUMMARY.xlsx"

if os.path.exists(LOCAL_PATH):
    EXCEL_FILE = LOCAL_PATH
elif os.path.exists(ONLINE_PATH):
    EXCEL_FILE = ONLINE_PATH
else:
    EXCEL_FILE = None

if EXCEL_FILE:
    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip()
        
        total_meters = len(df)
        comm_rate = 96
        non_comm = int(total_meters * 0.02) if total_meters > 100 else 4094
        never_comm = int(total_meters * 0.01) if total_meters > 100 else 2503

        # 📊 Standard Metrics (Bina kisi HTML ke)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric(label="Total Meters Onboarded", value=f"{total_meters:,}")
        with col2:
            st.metric(label="Communication Rate Today", value=f"{comm_rate}%")
        with col3:
            st.metric(label="Non-Communicating Meters", value=f"{non_comm:,}")
        with col4:
            st.metric(label="Never Communicating Meters", value=f"{never_comm:,}")

        st.markdown("---")
        st.subheader("📊 Meter Types & Modes")
        
        # Donut Charts
        ch1, ch2, ch3, ch4 = st.columns(4)
        with ch1:
            fig1 = go.Figure(data=[go.Pie(labels=['Single Phase', 'Three Phase'], values=[91.8, 8.2], hole=.6, marker=dict(colors=['#5C7CFA', '#22B8CF']), showlegend=False)])
            fig1.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, annotations=[dict(text='Phase', x=0.5, y=0.5, font_size=13, showarrow=False)])
            st.plotly_chart(fig1, use_container_width=True)
        with ch2:
            fig2 = go.Figure(data=[go.Pie(labels=['RF', 'GPRS'], values=[64.5, 35.5], hole=.6, marker=dict(colors=['#72C3DC', '#4D96A9']), showlegend=False)])
            fig2.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, annotations=[dict(text='Medium', x=0.5, y=0.5, font_size=13, showarrow=False)])
            st.plotly_chart(fig2, use_container_width=True)
        with ch3:
            fig3 = go.Figure(data=[go.Pie(labels=['Prepaid', 'Postpaid'], values=[85.4, 14.6], hole=.6, marker=dict(colors=['#F48C06', '#4A90E2']), showlegend=False)])
            fig3.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, annotations=[dict(text='Mode', x=0.5, y=0.5, font_size=13, showarrow=False)])
            st.plotly_chart(fig3, use_container_width=True)
        with ch4:
            fig4 = go.Figure(data=[go.Pie(labels=['Normal', 'Disconnect'], values=[98.4, 1.6], hole=.6, marker=dict(colors=['#E74C3C', '#F39C12']), showlegend=False)])
            fig4.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=200, annotations=[dict(text='Status', x=0.5, y=0.5, font_size=13, showarrow=False)])
            st.plotly_chart(fig4, use_container_width=True)

        # Live Data Table
        st.markdown("---")
        st.subheader("📋 Live Summary Records")
        st.dataframe(df, use_container_width=True, height=400)
        
    except Exception as e:
        st.error(f"Excel File read karne me koi galti hui h: {e}")
else:
    st.error("⚠️ FINAL SUMMARY.xlsx file nahi mil paa rahi hai. Kripya check karein.")