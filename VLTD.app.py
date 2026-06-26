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

DATA_FILE = os.path.join(
    SAVE_FOLDER,
    "tagging_requests.json"
)

EXCEL_FILE = os.path.join(
    SAVE_FOLDER,
    "VLTD Tagging data.xlsx"
)

IST = pytz.timezone("Asia/Kolkata")


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
            indent=4,
            ensure_ascii=False
        )

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

            "Yes"
            if r.get(
                "forwarded_to_lumax"
            )
            else "No",

            r.get("forwarded_time"),

            r.get("remarks"),

            r.get("tagging_status"),

            r.get("backend_tagged_by"),

            r.get("closure_date")

        ])

    wb.save(EXCEL_FILE)


# ================= APP =================

st.set_page_config(
    page_title="VLTD Tagging",
    layout="wide"
)

st.title("VLTD Tagging Management")


data = load_data()


menu = st.sidebar.selectbox(

    "Menu",

    [

        "Add Request",

        "Bulk Upload",

        "Vahan Status",

        "Backend Status",

        "Download Data"

    ]

)


# ================= ADD =================

if menu == "Add Request":

    st.subheader("Add Request")

    with st.form("add"):

        vin = st.text_input("VIN")

        state = st.text_input("State")

        dealer = st.text_input(
            "Dealer Code"
        )

        submit = st.form_submit_button(
            "Add"
        )

        if submit:

            data.append({

                "id":
                len(data)+1,

                "vin":
                vin,

                "state":
                state,

                "dealer_code":
                dealer,

                "request_date":

                datetime.now(
                    IST
                ).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),

                "vahan_status":
                "Pending",

                "vahan_tagged_by":
                None,

                "forwarded_to_lumax":
                False,

                "forwarded_time":
                None,

                "remarks":
                "",

                "tagging_status":
                None,

                "backend_tagged_by":
                None,

                "closure_date":
                None

            })

            save_data(data)

            st.success(
                "Request Added"
            )


# ================= BULK =================

elif menu == "Bulk Upload":

    file = st.file_uploader(

        "Upload",

        type=[
            "xlsx",
            "csv"
        ]

    )

    if file:

        if file.name.endswith(
            ".csv"
        ):

            df = pd.read_csv(
                file
            )

        else:

            df = pd.read_excel(
                file
            )

        st.dataframe(df)

        if st.button(
            "Confirm Upload"
        ):

            for _, row in df.iterrows():

                data.append({

                    "id":
                    len(data)+1,

                    "vin":
                    row["VIN"],

                    "state":
                    row["State"],

                    "dealer_code":
                    row[
                        "Dealer Code"
                    ],

                    "request_date":

                    datetime.now(
                        IST
                    ).strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                    "vahan_status":
                    "Pending",

                    "vahan_tagged_by":
                    None,

                    "forwarded_to_lumax":
                    False,

                    "forwarded_time":
                    None,

                    "remarks":
                    "",

                    "tagging_status":
                    None,

                    "backend_tagged_by":
                    None,

                    "closure_date":
                    None

                })

            save_data(data)

            st.success(
                "Upload Complete"
            )


# ================= VAHAN =================

elif menu == "Vahan Status":

    visible = [

        x

        for x in data

        if not x.get(
            "forwarded_to_lumax"
        )

    ]

    st.dataframe(
        pd.DataFrame(
            visible
        )
    )

    req = st.number_input(
        "Request ID",
        min_value=1
    )

    status = st.selectbox(

        "Vahan Status",

        [

            "Pending",

            "Done"

        ]

    )

    remarks = st.text_input(
        "Remarks"
    )

    tagged = st.selectbox(

        "Vahan Tagged By",

        [

            "Rajan",

            "Vishal",

            "Lumax Team"

        ]

    )

    if st.button(
        "Update Vahan"
    ):

        for r in data:

            if r["id"] == req:

                r[
                    "vahan_status"
                ] = status

                r[
                    "remarks"
                ] = remarks

                r[
                    "vahan_tagged_by"
                ] = tagged

        save_data(data)

        st.success(
            "Updated"
        )

    if st.button(
        "Forward To Lumax"
    ):

        done = False

        for r in data:

            if (

                r["id"] == req

                and

                r[
                    "vahan_status"
                ]
                ==
                "Done"

            ):

                r[
                    "forwarded_to_lumax"
                ] = True

                r[
                    "forwarded_time"
                ] = datetime.now(
                    IST
                ).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                done = True

        save_data(data)

        if done:

            st.success(
                "Forwarded"
            )

        else:

            st.error(
                "Status must be Done"
            )


# ================= BACKEND =================

elif menu == "Backend Status":

    backend = [

        x

        for x in data

        if x.get(
            "forwarded_to_lumax"
        )

    ]

    st.dataframe(
        pd.DataFrame(
            backend
        )
    )

    req = st.number_input(
        "Request ID ",
        min_value=1
    )

    status = st.selectbox(

        "Tagging Status",

        [

            "Pending",

            "Completed"

        ]

    )

    remarks = st.text_input(
        "Remarks "
    )

    tagged = st.selectbox(

        "Backend Tagged By",

        [

            "Rajan",

            "Vishal",

            "Lumax Team"

        ]

    )

    if st.button(
        "Save Backend"
    ):

        for r in data:

            if r["id"] == req:

                r[
                    "tagging_status"
                ] = status

                r[
                    "remarks"
                ] = remarks

                r[
                    "backend_tagged_by"
                ] = tagged

                if (

                    status
                    ==
                    "Completed"

                ):

                    r[
                        "closure_date"
                    ] = datetime.now(
                        IST
                    ).strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )

        save_data(data)

        st.success(
            "Backend Updated"
        )


# ================= DOWNLOAD =================

elif menu == "Download Data":

    save_data(data)

    with open(
        EXCEL_FILE,
        "rb"
    ) as f:

        st.download_button(

            "⬇ Download Excel",

            data=f,

            file_name=
            "VLTD Tagging data.xlsx",

            mime=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

    st.success(
        EXCEL_FILE
    )
