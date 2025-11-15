
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path

st.set_page_config(page_title="Inventory Simulator", layout="wide")

# ===========================
# 1. Load Data
# ===========================
file_path = Path(r"C:\Users\ashalaby\OneDrive - Halwani Bros\Planning - Sources\Materials.xlsb")
sheet_name = "Data"

@st.cache_data
def load_data():
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="pyxlsb")
    df["ItemNumber"] = df["ItemNumber"].astype(str)
    return df

df = load_data()

# ===========================
# 2. Sidebar Filters
# ===========================
st.sidebar.header("Filters")

item_search = st.sidebar.text_input("ابحث باسم أو كود الصنف")
factories = sorted(df["Factory"].dropna().astype(str).unique().tolist())
storages = sorted(df["Storagetype"].dropna().astype(str).unique().tolist())
families = sorted(df["Family"].dropna().astype(str).unique().tolist())


selected_factory = st.sidebar.multiselect("اختر Factory", factories)
selected_storage = st.sidebar.multiselect("اختر Storagetype", [])
selected_family = st.sidebar.multiselect("اختر Family", [])

if selected_factory:
    df = df[df["Factory"].isin(selected_factory)]
    storages = sorted(df["Storagetype"].dropna().unique().tolist())
    selected_storage = st.sidebar.multiselect("اختر Storagetype", storages)

if selected_storage:
    df = df[df["Storagetype"].isin(selected_storage)]
    families = sorted(df["Family"].dropna().unique().tolist())
    selected_family = st.sidebar.multiselect("اختر Family", families)

if selected_family:
    df = df[df["Family"].isin(selected_family)]

if item_search.strip():
    df = df[(df["ItemName"].str.contains(item_search, case=False, na=False)) | 
            (df["ItemNumber"].str.contains(item_search, case=False, na=False))]

# ===========================
# 3. Calculations
# ===========================
df["CurAS_value"] = df["CurAS"] * df["Cost"]
df["CurAP_value"] = df["CurAP"] * df["Cost"]
df["ClosingStock"] = df["StartInv"] + (df["CurST"] - df["CurAS"]) - (df["CurAPP"] - df["CurAP"])
df["ClosingStock_value"] = df["ClosingStock"] * df["Cost"]

# ===========================
# 4. Views
# ===========================
tab_inventory, tab_fg, tab_bom = st.tabs(["Inventory", "FG", "BOM"])

with tab_inventory:
    st.subheader("📊 Inventory Dashboard")

    summary = df.groupby(["Factory", "Storagetype", "Family"], as_index=False).agg({
        "StartInv": "sum",
        "CurAS": "sum",
        "JanAPP": "sum",
        "ClosingStock": "sum",
        "Cost": "mean"
    })

    summary["StartInv_value"] = summary["StartInv"] * summary["Cost"]
    summary["CurAS_value"] = summary["CurAS"] * summary["Cost"]
    summary["CurAP_value"] = summary["JanAPP"] * summary["Cost"]
    summary["Closing_value"] = summary["ClosingStock"] * summary["Cost"]

    st.dataframe(summary.style.format(thousands=","))

    # ===========================
    # 5. Plotly Chart (Quantity + Value)
    # ===========================
    st.subheader("📈 Inventory Timeline")

    volume_or_value = st.radio("اختر ما تريد عرضه:", ["Volume", "Value"], horizontal=True)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.15,
                        subplot_titles=("Inventory Quantities", "Inventory Values"))

    for item in df["ItemNumber"].unique():
        temp = df[df["ItemNumber"] == item]
        if volume_or_value == "Volume":
            fig.add_trace(go.Bar(x=["OH", "CurAS", "CurAP", "Closing"],
                                 y=[temp["StartInv"].sum(), temp["CurAS"].sum(), temp["JanAPP"].sum(), temp["ClosingStock"].sum()],
                                 name=f"{item} Qty"), row=1, col=1)
        else:
            fig.add_trace(go.Bar(x=["OH", "CurAS", "CurAP", "Closing"],
                                 y=[temp["StartInv_value"].sum(), temp["CurAS_value"].sum(), temp["CurAP_value"].sum(), temp["ClosingStock_value"].sum()],
                                 name=f"{item} Value"), row=2, col=1)

    fig.update_layout(height=700, showlegend=True, barmode="group")
    st.plotly_chart(fig, use_container_width=True)

with tab_fg:
    st.info("FG view content coming soon.")

with tab_bom:
    st.info("BOM view content coming soon.")
