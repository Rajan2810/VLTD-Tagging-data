import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import pytz

DATA_FILE="tagging_requests.json"
EXCEL_FILE="tagging_requests.xlsx"

IST=pytz.timezone("Asia/Kolkata")


def load_data():

    if os.path.exists(DATA_FILE):

        with open(
            DATA_FILE,
            "r"
        ) as f:

            return json.load(f)

    return []


def save_data(data):

    with open(
        DATA_FILE,
        "w"
    ) as f:

        json.dump(
            data,
            f,
            indent=4
        )

    pd.DataFrame(
        data
    ).to_excel(
        EXCEL_FILE,
        index=False
    )


st.title(
"VLTD Tagging Data Tracking"
)

vin=st.text_input(
"Enter VIN"
)

if st.button(
"Search"
):

    data=load_data()

    row=None

    for r in data:

        if r["vin"]==vin:

            row=r

            break

    if not row:

        st.error(
        "VIN NOT MAPPED"
        )

    else:

        if row["vahan_status"]=="Pending":

            output=(
                "LIT1|"
                "LIT1LIA022600053884|"
                "860982089793432|"
                "8991102506560327863|"
                "03/2026|214"
            )

            st.code(output)

            state=st.selectbox(
                "Select State",
                [
                    "Haryana",
                    "Rajasthan",
                    "UP",
                    "Delhi"
                ]
            )

            if st.button(
                "Submit"
            ):

                row["state"]=state

                row[
                    "request_date"
                ]=(
                    datetime.now(
                        IST
                    )
                    .strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

                save_data(
                    data
                )

                st.success(
                    "Submitted"
                )

        else:

            st.success(
                "Already Completed"
            )

if os.path.exists(
EXCEL_FILE
):

    with open(
        EXCEL_FILE,
        "rb"
    ) as f:

        st.download_button(
            "Download Excel",
            f,
            EXCEL_FILE
        )
