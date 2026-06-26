import streamlit as st
import pandas as pd
import openpyxl
import json
import os
from datetime import datetime
import pytz

# ---------------- Config ----------------
DATA_FILE = "tagging_requests.json"
EXCEL_FILE = "tagging_requests.xlsx"
IST = pytz.timezone("Asia/Kolkata")


# ---------------- Functions ----------------
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return []


def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=4)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tagging Requests"

    headers = [
        "ID",
        "VIN",
        "State",
        "Dealer Code",
        "Request Date",
        "Vahan Status",
        "Forwarded",
        "Forwarded Time",
        "Remarks",
        "Tagging Status",
        "Closure Date"
    ]

    ws.append(headers)

    for r in data:
        ws.append([
            r["id"],
            r["vin"],
            r["state"],
            r["dealer_code"],
            r["request_date"],
            r["vahan_status"],
            r["forwarded_to_lumax"],
            r["forwarded_time"],
            r["remarks"],
            r["tagging_status"],
            r["closure_date"]
        ])

    wb.save(EXCEL_FILE)


# ---------------- UI ----------------
st.set_page_config(page_title="VLTD Tagging", layout="wide")

st.title("VLTD Tagging Management")

menu = st.sidebar.selectbox(
    "Menu",
    [
        "Add Request",
        "Bulk Upload",
        "Vahan Status",
        "Backend Status",
        "Download"
    ]
)

data = load_data()


# ---------------- Add Request ----------------
if menu == "Add Request":

    st.subheader("Add Request")

    with st.form("add_form"):

        vin = st.text_input("VIN")
        state = st.text_input("State")
        dealer = st.text_input("Dealer Code")

        submit = st.form_submit_button("Submit")

        if submit:

            data.append({
                "id": len(data)+1,
                "vin": vin,
                "state": state,
                "dealer_code": dealer,
                "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M"),
                "vahan_status": "Pending",
                "forwarded_to_lumax": False,
                "forwarded_time": None,
                "remarks": "",
                "tagging_status": None,
                "closure_date": None
            })

            save_data(data)

            st.success("Request Added")


# ---------------- Bulk Upload ----------------
elif menu == "Bulk Upload":

    file = st.file_uploader(
        "Upload Excel",
        type=["xlsx", "csv"]
    )

    if file:

        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        st.dataframe(df)

        if st.button("Confirm Upload"):

            for _, row in df.iterrows():

                data.append({
                    "id": len(data)+1,
                    "vin": row["VIN"],
                    "state": row["State"],
                    "dealer_code": row["Dealer Code"],
                    "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M"),
                    "vahan_status": "Pending",
                    "forwarded_to_lumax": False,
                    "forwarded_time": None,
                    "remarks": "",
                    "tagging_status": None,
                    "closure_date": None
                })

            save_data(data)

            st.success("Upload Completed")


# ---------------- Vahan Status ----------------
elif menu == "Vahan Status":

    st.subheader("Vahan Status")

    visible = [
        r for r in data
        if not r.get("forwarded_to_lumax", False)
    ]

    st.dataframe(pd.DataFrame(visible))

    req_id = st.number_input(
        "Request ID",
        min_value=1,
        step=1,
        key="vahan"
    )

    status = st.selectbox(
        "Vahan Status",
        ["Pending", "Done"]
    )

    remarks = st.text_input("Remarks")

    tagged_by = st.selectbox(
        "Tagged By",
        [
            "Rajan",
            "Vishal",
            "Lumax Team"
        ],
        key="vahan_tag"
    )

    if st.button("Update Status"):

        for r in data:

            if r["id"] == req_id:

                r["vahan_status"] = status
                r["remarks"] = remarks
                r["tagged_by"] = tagged_by

        save_data(data)

        st.success("Vahan Status Updated")

    if st.button("Forward To Lumax"):

        success = False

        for r in data:

            if (
                r["id"] == req_id
                and r["vahan_status"] == "Done"
            ):

                r["forwarded_to_lumax"] = True

                r["forwarded_time"] = (
                    datetime.now(IST)
                    .strftime("%Y-%m-%d %H:%M:%S")
                )

                success = True

        save_data(data)

        if success:
            st.success("Forwarded To Lumax")
        else:
            st.error("Status must be Done")


# ---------------- Backend ----------------
elif menu == "Backend Status":

    st.subheader("Backend Status")

    backend = [

        r for r in data

        if r.get("forwarded_to_lumax")
    ]

    st.dataframe(
        pd.DataFrame(backend)
    )

    req_id = st.number_input(
        "Request ID ",
        min_value=1,
        step=1,
        key="backend"
    )

    status = st.selectbox(
        "Tagging Status",
        [
            "Pending",
            "Completed"
        ]
    )

    remarks = st.text_input(
        "Remarks"
    )

    tagged_by = st.selectbox(
        "Tagged By",
        [
            "Rajan",
            "Vishal",
            "Lumax Team"
        ],
        key="backend_tag"
    )

    if st.button(
        "Save Backend"
    ):

        for r in data:

            if r["id"] == req_id:

                r["tagging_status"] = status
                r["remarks"] = remarks
                r["tagged_by"] = tagged_by

                if status == "Completed":

                    r["closure_date"] = (
                        datetime.now(IST)
                        .strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    )

        save_data(data)

        st.success(
            "Backend Updated"
        )

# ---------------- Download ----------------
elif menu == "Download":

    if os.path.exists(EXCEL_FILE):

        with open(EXCEL_FILE, "rb") as f:

            st.download_button(
                "Download Excel",
                f,
                file_name="tagging_requests.xlsx"
            )
