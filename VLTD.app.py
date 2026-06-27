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

SAVE_FOLDER = r"D:\OneDrive - 太思科技股份有限公司"

DATA_FILE = os.path.join(
    SAVE_FOLDER,
    "tagging_requests.json"
)

EXCEL_FILE = os.path.join(
    SAVE_FOLDER,
    "VLTD tagging Data.xlsx"
)

IST = pytz.timezone(
    "Asia/Kolkata"
)


# ================= LOAD =================

def load_data():

    if os.path.exists(DATA_FILE):

        with open(
            DATA_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            return json.load(f)

    return []


# ================= ADD HERE =================
# PUT upload_to_onedrive() HERE

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

        with open(local_file, "rb") as file:

            client.me.drive.root.get_by_path(
                st.secrets["ONEDRIVE_FILE"]
            ).upload(
                file.read()
            ).execute_query()

        return True

    except Exception as e:

        st.error(
            f"Upload failed: {e}"
        )

        return False


# ================= SAVE =================

def save_data(data):

    with open(
        DATA_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            indent=4
        )

    wb = openpyxl.Workbook()

    ws = wb.active

    ws.title = "VLTD Tagging"

    wb.save(EXCEL_FILE)

    # ADD THIS
    upload_to_onedrive(
        EXCEL_FILE
    )
# ================= SAVE DATA =================

def save_data(data):

    # Save JSON
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            indent=4,
            ensure_ascii=False
        )

    # Save Excel
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

    st.write("Excel saved:", EXCEL_FILE)

result = upload_to_onedrive(EXCEL_FILE)

st.write("Upload status:", result)

    # Upload to OneDrive
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
            client
            .me
            .drive
            .root
            .get_by_path(
                st.secrets["ONEDRIVE_FILE"]
            )
            .upload(content)
            .execute_query()
        )

        return True

    except Exception as e:

        st.error(
            f"Upload failed: {e}"
        )

        return False
# ================= UI =================

st.set_page_config(
    page_title="VLTD Tagging",
    layout="wide"
)

st.title(
    "VLTD Tagging System"
)

data = load_data()

# CREATE MENU FIRST
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

    st.subheader("📊 Status Dashboard")

    df = pd.DataFrame(data)

    if df.empty:
        st.warning("No data available")
        st.stop()

    # Convert date
    df["request_date"] = pd.to_datetime(
        df["request_date"],
        errors="coerce"
    )

    # ---------- DATE FILTER ----------

    st.subheader("📅 Filter")

    col1, col2 = st.columns(2)

    with col1:
        start_date = st.date_input(
            "Start Date",
            value=df["request_date"].min().date()
        )

    with col2:
        end_date = st.date_input(
            "End Date",
            value=df["request_date"].max().date()
        )

    filtered_df = df[

        (df["request_date"].dt.date >= start_date)

        &

        (df["request_date"].dt.date <= end_date)

    ]

    # ---------- KPI ----------

    total_requests = len(
        filtered_df
    )

    total_vahan = len(

        filtered_df[
            filtered_df[
                "vahan_status"
            ] == "Done"
        ]

    )

    total_backend = len(

        filtered_df[
            filtered_df[
                "tagging_status"
            ] == "Completed"
        ]

    )

    total_forward = len(

        filtered_df[
            filtered_df[
                "forwarded_to_lumax"
            ] == True
        ]

    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📦 Total Requests",
        total_requests
    )

    c2.metric(
        "🚘 Vahan Tagged",
        total_vahan
    )

    c3.metric(
        "🏁 Backend Tagged",
        total_backend
    )

    c4.metric(
        "📤 Forwarded",
        total_forward
    )

    # ---------- SUMMARY ----------

    st.subheader(
        "📈 Status Summary"
    )

    chart = pd.DataFrame({

        "Category": [

            "Requests",

            "Vahan",

            "Backend",

            "Forwarded"

        ],

        "Count": [

            total_requests,

            total_vahan,

            total_backend,

            total_forward

        ]

    })

    st.bar_chart(

        chart.set_index(
            "Category"
        )

    )

    # ---------- STATE ----------

    st.subheader(
        "📍 State Wise Requests"
    )

    state_chart = (

        filtered_df

        .groupby(
            "state"
        )

        .size()

        .reset_index(
            name="Total"
        )

    )

    st.bar_chart(

        state_chart.set_index(
            "state"
        )

    )

    # ---------- TAGGED BY ----------

    st.subheader(
        "👤 Vahan Tagged By"
    )

    tagged = (

        filtered_df[
            "vahan_tagged_by"
        ]

        .fillna(
            "Pending"
        )

        .value_counts()

    )

    st.bar_chart(
        tagged
    )

    # ---------- RECORDS ----------

    st.subheader(
        "📋 Records"
    )

    st.dataframe(
        filtered_df,
        use_container_width=True
    )

    # ---------- DOWNLOAD ----------

    csv = filtered_df.to_csv(
        index=False
    )

    st.download_button(

        "⬇ Download Dashboard Data",

        csv,

        file_name="Dashboard_Report.csv",

        mime="text/csv"

    )
    # ---------------- DOWNLOAD FILTERED ----------------

    csv = filtered_df.to_csv(
        index=False
    )

    st.download_button(
    label="⬇ Download Dashboard Data",
    data=csv,
    file_name="Dashboard_Report.csv",
    mime="text/csv",
    key="dashboard_download"
)

    

    # ---------------- DATA TABLE ----------------

    st.subheader(
        "📋 Filtered Records"
    )

    st.dataframe(
        filtered_df,
        use_container_width=True
    )

    # ---------------- DOWNLOAD FILTERED ----------------

    csv = filtered_df.to_csv(
        index=False
    )
# ================= ADD REQUEST =================

elif menu == "Add Request":

    st.subheader("Add New Request")

    vin = st.text_input("VIN")

    state = st.text_input("State")

    dealer = st.text_input("Dealer Code")

    if st.button(
        "Add Request",
        key="add_request_btn"
    ):

        if vin == "" or state == "" or dealer == "":

            st.error(
                "Please fill all fields"
            )

        else:

            data.append({

                "id": len(data) + 1,

                "vin": vin,

                "state": state,

                "dealer_code": dealer,

                "request_date": datetime.now(
                    IST
                ).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

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

            st.success(
                "Request Added Successfully"
            )

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
