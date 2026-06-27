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


# ================= LOAD =================

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

        target_path = st.secrets["ONEDRIVE_FILE"].replace("\x00", "")

        with open(local_file, "rb") as f:
            content = f.read()

        (
            client.me.drive.root
            .get_by_path(target_path)
            .upload_file(content)
            .execute_query()
        )

        return True

    except Exception as e:
        st.error(f"Upload failed: {e}")
        return False


# ================= SAVE =================

def save_data(data):

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

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


# ================= APP =================

st.set_page_config(page_title="VLTD Tagging", layout="wide")
st.title("🚗 VLTD Tagging System")

data = load_data()

menu = st.sidebar.selectbox(
    "Menu",
    ["Dashboard", "Add Request", "Bulk Upload", "Vahan Status", "Backend Status", "Download"]
)


# ================= DASHBOARD =================

if menu == "Dashboard":

    st.subheader("📊 Dashboard")

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data found")
        st.stop()

    st.dataframe(df, use_container_width=True)


# ================= ADD REQUEST =================

elif menu == "Add Request":

    st.subheader("➕ Add Request")

    vin = st.text_input("VIN")
    state = st.text_input("State")
    dealer = st.text_input("Dealer Code")

    if st.button("Submit"):

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
        st.success("Added successfully")
        st.rerun()


# ================= VAHAN STATUS (WITH CHECKBOX SELECTION) =================

elif menu == "Vahan Status":

    st.subheader("🚘 Vahan Status - Select Records")

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data available")
        st.stop()

    selected_ids = []

    st.write("### Select Records:")

    for i, row in df.iterrows():
        checked = st.checkbox(
            f"ID {row['id']} | VIN: {row['vin']} | Status: {row['vahan_status']}",
            key=f"chk_{row['id']}"
        )

        if checked:
            selected_ids.append(row["id"])

    st.markdown("---")

    status = st.selectbox("Vahan Status", ["Pending", "Done"])
    tagged_by = st.selectbox("Tagged By", ["Rajan", "Vishal", "Lumax Team"])
    remarks = st.text_input("Remarks")

    if st.button("Update Selected"):

        if not selected_ids:
            st.error("No records selected")
        else:
            for r in data:
                if r["id"] in selected_ids:
                    r["vahan_status"] = status
                    r["vahan_tagged_by"] = tagged_by
                    r["remarks"] = remarks

            save_data(data)
            st.success("Selected records updated")


    st.markdown("---")

    if st.button("Forward Selected to Lumax"):

        if not selected_ids:
            st.error("No records selected")

        else:
            success = False

            for r in data:
                if r["id"] in selected_ids:
                    if r["vahan_status"] == "Done":
                        r["forwarded_to_lumax"] = True
                        r["forwarded_time"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
                        success = True

            save_data(data)

            if success:
                st.success("Selected records forwarded")
            else:
                st.error("Only 'Done' records can be forwarded")


# ================= BACKEND =================

elif menu == "Backend Status":

    st.subheader("🛠 Backend Status")

    df = pd.DataFrame(data)
    st.dataframe(df)

    selected_ids = []

    for i, row in df.iterrows():
        if st.checkbox(f"Select ID {row['id']}", key=f"b_{row['id']}"):
            selected_ids.append(row["id"])

    status = st.selectbox("Status", ["Pending", "Completed"])
    tagged_by = st.selectbox("Tagged By", ["Rajan", "Vishal", "Lumax Team"])
    remarks = st.text_input("Remarks")

    if st.button("Update Backend"):

        for r in data:
            if r["id"] in selected_ids:
                r["tagging_status"] = status
                r["backend_tagged_by"] = tagged_by
                r["remarks"] = remarks

                if status == "Completed":
                    r["closure_date"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

        save_data(data)
        st.success("Backend updated")


# ================= DOWNLOAD =================

elif menu == "Download":

    save_data(data)

    with open(EXCEL_FILE, "rb") as f:
        st.download_button(
            "Download Excel",
            f,
            file_name="VLTD_tagging.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
