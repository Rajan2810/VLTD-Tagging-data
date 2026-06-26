import streamlit as st
import pandas as pd
import openpyxl
import json
import os
from datetime import datetime
import pytz

# ================= CONFIG =================

SAVE_FOLDER = r"D:\OneDrive - 太思科技股份有限公司\Desktop\rajan\python\VLTD"

os.makedirs(SAVE_FOLDER, exist_ok=True)

DATA_FILE = os.path.join(SAVE_FOLDER, "tagging_requests.json")

EXCEL_FILE = os.path.join(SAVE_FOLDER, "VLTD Tagging data.xlsx")

IST = pytz.timezone("Asia/Kolkata")


# ================= LOAD DATA =================

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


# ================= SAVE DATA =================

def save_data(data):

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VLTD Tagging"

    headers = [
        "ID",
        "VIN",
        "State",
        "Dealer Code",
        "Request Date",
        "Vahan Status",
        "Vahan tagged by",
        "Forwarded to Lumax",
        "Forwarded Time",
        "Remarks",
        "Tagging Status",
        "Statebackend tagged by",
        "Closure Date"
    ]

    ws.append(headers)

    for r in data:
        ws.append([
            r.get("id"),
            r.get("vin"),
            r.get("state"),
            r.get("dealer_code"),
            r.get("request_date"),
            r.get("vahan_status"),
            r.get("vahan_tagged_by"),
            "Yes" if r.get("forwarded_to_lumax") else "No",
            r.get("forwarded_time"),
            r.get("remarks"),
            r.get("tagging_status"),
            r.get("backend_tagged_by"),
            r.get("closure_date")
        ])

    wb.save(EXCEL_FILE)


# ================= UI =================

st.set_page_config(page_title="VLTD Tagging", layout="wide")

st.title("VLTD Tagging System")

data = load_data()

menu = st.sidebar.selectbox(
    "Menu",
    [
        "Dashboard",
        "Add Request",
        "Bulk Upload",
        "Vahan Status",
        "Backend Status",
        "Download Data"
    ]
)


# ================= DASHBOARD =================

if menu == "Dashboard":

    st.subheader("📊 Dashboard")

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data available")
        st.stop()

    df["request_date"] = pd.to_datetime(df["request_date"])

    start_date = st.date_input("Start Date")
    end_date = st.date_input("End Date")

    filtered_df = df[
        (df["request_date"].dt.date >= start_date) &
        (df["request_date"].dt.date <= end_date)
    ]

    total_requests = len(filtered_df)

    total_vahan_done = len(
        filtered_df[filtered_df["vahan_status"] == "Done"]
    )

    total_backend_done = len(
        filtered_df[filtered_df["tagging_status"] == "Completed"]
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Requests", total_requests)
    col2.metric("Vahan Done", total_vahan_done)
    col3.metric("Backend Completed", total_backend_done)

    chart_data = pd.DataFrame({
        "Category": [
            "Total Requests",
            "Vahan Done",
            "Backend Completed"
        ],
        "Count": [
            total_requests,
            total_vahan_done,
            total_backend_done
        ]
    })

    st.bar_chart(chart_data.set_index("Category"))

    st.dataframe(filtered_df)


# ================= ADD REQUEST =================

elif menu == "Add Request":

    vin = st.text_input("VIN")
    state = st.text_input("State")
    dealer = st.text_input("Dealer Code")

    if st.button("Add"):

        data.append({
            "id": len(data) + 1,
            "vin": vin,
            "state": state,
            "dealer_code": dealer,
            "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
            "vahan_status": "Pending",
            "vahan_tagged_by": None,
            "forwarded_to_lumax": False,
            "forwarded_time": None,
            "remarks": "",
            "tagging_status": None,
            "backend_tagged_by": None,
            "closure_date": None
        })

        save_data(data)
        st.success("Request Added")


# ================= BULK UPLOAD =================

elif menu == "Bulk Upload":

    file = st.file_uploader("Upload Excel/CSV", type=["xlsx", "csv"])

    if file:

        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        st.dataframe(df)

        if st.button("Confirm Upload"):

            for _, row in df.iterrows():

                data.append({
                    "id": len(data) + 1,
                    "vin": row["VIN"],
                    "state": row["State"],
                    "dealer_code": row["Dealer Code"],
                    "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                    "vahan_status": "Pending",
                    "vahan_tagged_by": None,
                    "forwarded_to_lumax": False,
                    "forwarded_time": None,
                    "remarks": "",
                    "tagging_status": None,
                    "backend_tagged_by": None,
                    "closure_date": None
                })

            save_data(data)
            st.success("Bulk Upload Completed")


# ================= VAHAN STATUS =================

elif menu == "Vahan Status":

    st.subheader("Vahan Status")

    pending = [r for r in data if not r.get("forwarded_to_lumax")]

    st.dataframe(pd.DataFrame(pending))

    req = st.number_input("Request ID", min_value=1)

    status = st.selectbox("Vahan Status", ["Pending", "Done"])

    remarks = st.text_input("Remarks")

    tagged_by = st.selectbox(
        "Vahan Tagged By",
        ["Rajan", "Vishal", "Lumax Team"]
    )

    if st.button("Update Vahan"):

        for r in data:
            if r["id"] == req:
                r["vahan_status"] = status
                r["remarks"] = remarks
                r["vahan_tagged_by"] = tagged_by

        save_data(data)
        st.success("Vahan Updated")

    if st.button("Forward To Lumax"):

        ok = False

        for r in data:
            if r["id"] == req and r["vahan_status"] == "Done":
                r["forwarded_to_lumax"] = True
                r["forwarded_time"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                ok = True

        save_data(data)

        if ok:
            st.success("Forwarded to Lumax")
        else:
            st.error("Set status to Done first")


# ================= BACKEND =================

elif menu == "Backend Status":

    st.subheader("Backend Status")

    backend = [r for r in data if r.get("forwarded_to_lumax")]

    st.dataframe(pd.DataFrame(backend))

    req = st.number_input("Request ID", min_value=1)

    status = st.selectbox("Tagging Status", ["Pending", "Completed"])

    remarks = st.text_input("Remarks")

    tagged_by = st.selectbox(
        "Backend Tagged By",
        ["Rajan", "Vishal", "Lumax Team"]
    )

    if st.button("Save Backend"):

        for r in data:
            if r["id"] == req:

                r["tagging_status"] = status
                r["remarks"] = remarks
                r["backend_tagged_by"] = tagged_by

                if status == "Completed":
                    r["closure_date"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

        save_data(data)
        st.success("Backend Updated")


# ================= DOWNLOAD =================

elif menu == "Download Data":

    save_data(data)

    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "⬇ Download Excel",
            f,
            file_name="VLTD Tagging data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success(f"Saved at: {EXCEL_FILE}")
