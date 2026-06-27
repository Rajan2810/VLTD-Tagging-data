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

DATA_FILE = os.path.join(SAVE_FOLDER, "vltd_data.json")
EXCEL_FILE = os.path.join(SAVE_FOLDER, "VLTD_tagging.xlsx")

IST = pytz.timezone("Asia/Kolkata")


# ================= LOAD DATA =================

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


# ================= SAVE DATA =================

def save_data(data):

    # ---------- JSON SAVE ----------
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

    # ---------- EXCEL SAVE ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VLTD Tagging"

    headers = [
        "ID", "VIN", "State", "Dealer Code", "Request Date",
        "Vahan Status", "Vahan Tagged By",
        "Forwarded To Lumax", "Forwarded Time",
        "Remarks", "Tagging Status",
        "Backend Tagged By", "Closure Date"
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

    upload_to_onedrive(EXCEL_FILE)


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

        target_path = st.secrets["ONEDRIVE_FILE"]

        with open(local_file, "rb") as f:
            content = f.read()

        (
            client.me.drive.root
            .get_by_path(target_path)
            .upload_file(content)   # ✅ FIXED METHOD
            .execute_query()
        )

        return True

    except Exception as e:
        st.error(f"Upload failed: {e}")
        return False


# ================= APP UI =================

st.set_page_config(page_title="VLTD Tagging", layout="wide")
st.title("🚗 VLTD Tagging System")

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

    df["request_date"] = pd.to_datetime(df["request_date"], errors="coerce")

    st.dataframe(df, use_container_width=True)

    st.download_button(
        "⬇ Download CSV",
        df.to_csv(index=False),
        file_name="dashboard.csv",
        mime="text/csv"
    )


# ================= ADD REQUEST =================

elif menu == "Add Request":

    st.subheader("➕ Add Request")

    vin = st.text_input("VIN")
    state = st.text_input("State")
    dealer = st.text_input("Dealer Code")

    if st.button("Submit"):

        if not vin or not state or not dealer:
            st.error("Please fill all fields")

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
            st.success("Request added")
            st.rerun()


# ================= BULK UPLOAD =================

elif menu == "Bulk Upload":

    file = st.file_uploader("Upload Excel/CSV", type=["xlsx", "csv"])

    if file:

        if file.name.endswith("csv"):
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
            st.success("Bulk upload completed")


# ================= VAHAN STATUS =================

elif menu == "Vahan Status":

    st.subheader("🚘 Vahan Status")

    df = pd.DataFrame(data)
    st.dataframe(df)

    req = st.number_input("Request ID", min_value=1, step=1)

    status = st.selectbox("Vahan Status", ["Pending", "Done"])
    remarks = st.text_input("Remarks")

    tagged_by = st.selectbox(
        "Tagged By",
        ["Rajan", "Vishal", "Lumax Team"]
    )

    if st.button("Update Vahan Status"):

        updated = False

        for r in data:
            if r["id"] == req:
                r["vahan_status"] = status
                r["remarks"] = remarks
                r["vahan_tagged_by"] = tagged_by
                updated = True

        save_data(data)

        if updated:
            st.success("Updated successfully")
        else:
            st.error("Invalid Request ID")

    st.markdown("---")

    st.subheader("📤 Forward to Lumax")

    if st.button("Forward To Lumax"):

        ok = False

        for r in data:
            if r["id"] == req:

                if r["vahan_status"] == "Done":
                    r["forwarded_to_lumax"] = True
                    r["forwarded_time"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                    ok = True

        save_data(data)

        if ok:
            st.success("Forwarded successfully")
        else:
            st.error("Set Vahan Status = Done first")


# ================= BACKEND =================

elif menu == "Backend Status":

    st.subheader("🛠 Backend Status")

    df = pd.DataFrame([r for r in data if r.get("forwarded_to_lumax")])
    st.dataframe(df)

    req = st.number_input("Request ID", min_value=1, step=1)

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
        st.success("Backend updated")


# ================= DOWNLOAD =================

elif menu == "Download Data":

    save_data(data)

    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "⬇ Download Excel",
            f,
            file_name="VLTD_tagging.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
