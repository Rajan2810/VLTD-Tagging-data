import streamlit as st
import pandas as pd
import openpyxl
import json
import os
from datetime import datetime
import pytz

from office365.graph_client import GraphClient
from office365.runtime.auth.client_credential import ClientCredential


# ================= CONFIG =================

SAVE_FOLDER = "data"
os.makedirs(SAVE_FOLDER, exist_ok=True)

DATA_FILE = os.path.join(SAVE_FOLDER, "tagging_requests.json")
EXCEL_FILE = os.path.join(SAVE_FOLDER, "VLTD_tagging.xlsx")

IST = pytz.timezone("Asia/Kolkata")


# ================= LOAD DATA =================

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


# ================= ONEDRIVE UPLOAD =================

def upload_to_onedrive(local_file):
    try:
        credentials = ClientCredential(
            st.secrets["CLIENT_ID"],
            st.secrets["CLIENT_SECRET"]
        )

        client = GraphClient(
            st.secrets["TENANT_ID"],
            credentials
        )

        with open(local_file, "rb") as f:
            content = f.read()

        (
            client.me.drive.root
            .get_by_path(st.secrets["ONEDRIVE_FILE"])
            .upload(content)
            .execute_query()
        )

        return True

    except Exception as e:
        st.error(f"Upload failed: {e}")
        return False


# ================= SAVE DATA =================

def save_data(data):

    # ---------- Save JSON ----------
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    # ---------- Save Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VLTD Tagging"

    headers = [
        "ID", "VIN", "State", "Dealer Code", "Request Date",
        "Vahan Status", "Vahan tagged by", "Forwarded to Lumax",
        "Forwarded Time", "Remarks", "Tagging Status",
        "Statebackend tagged by", "Closure Date"
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

    st.write("Excel saved:", EXCEL_FILE)

    # ---------- Upload ----------
    result = upload_to_onedrive(EXCEL_FILE)
    st.write("Upload status:", result)


# ================= UI =================

st.set_page_config(page_title="VLTD Tagging", layout="wide")
st.title("VLTD Tagging System")

data = load_data()

menu = st.sidebar.selectbox(
    "Menu",
    ["Dashboard", "Add Request", "Bulk Upload",
     "Vahan Status", "Backend Status", "Download Data"]
)


# ================= DASHBOARD =================

if menu == "Dashboard":

    st.subheader("Dashboard")

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data available")
        st.stop()

    df["request_date"] = pd.to_datetime(df["request_date"], errors="coerce")

    st.dataframe(df, use_container_width=True)

    csv = df.to_csv(index=False)
    st.download_button(
        "Download CSV",
        csv,
        file_name="dashboard.csv",
        mime="text/csv"
    )


# ================= ADD REQUEST =================

elif menu == "Add Request":

    st.subheader("Add Request")

    vin = st.text_input("VIN")
    state = st.text_input("State")
    dealer = st.text_input("Dealer Code")

    if st.button("Add"):

        if not vin or not state or not dealer:
            st.error("Fill all fields")
        else:

            data.append({
                "id": len(data) + 1,
                "vin": vin,
                "state": state,
                "dealer_code": dealer,
                "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
                "vahan_status": "Pending",
                "vahan_tagged_by": "",
                "forwarded_to_lumax": False,
                "forwarded_time": "",
                "remarks": "",
                "tagging_status": "",
                "backend_tagged_by": "",
                "closure_date": ""
            })

            save_data(data)
            st.success("Request Added")
            st.rerun()


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
                    "vahan_tagged_by": "",
                    "forwarded_to_lumax": False,
                    "forwarded_time": "",
                    "remarks": "",
                    "tagging_status": "",
                    "backend_tagged_by": "",
                    "closure_date": ""
                })

            save_data(data)
            st.success("Bulk Upload Done")


# ================= VAHAN STATUS =================

elif menu == "Vahan Status":

    st.subheader("Vahan Status")

    req = st.number_input("Request ID", min_value=1)
    status = st.selectbox("Status", ["Pending", "Done"])

    if st.button("Update"):

        for r in data:
            if r["id"] == req:
                r["vahan_status"] = status

        save_data(data)
        st.success("Updated")


# ================= BACKEND =================

elif menu == "Backend Status":

    st.subheader("Backend Status")

    req = st.number_input("Request ID", min_value=1)
    status = st.selectbox("Tagging Status", ["Pending", "Completed"])

    if st.button("Save"):

        for r in data:
            if r["id"] == req:
                r["tagging_status"] = status

                if status == "Completed":
                    r["closure_date"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

        save_data(data)
        st.success("Updated")


# ================= DOWNLOAD =================

elif menu == "Download Data":

    save_data(data)

    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "Download Excel",
            f,
            file_name="VLTD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
