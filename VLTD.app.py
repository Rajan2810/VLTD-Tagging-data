import streamlit as st
import pandas as pd
import os

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

st.set_page_config(
    page_title="VLTD Dashboard",
    layout="wide"
)

st.title("VLTD Dashboard")


def load_data():

    columns = [
        "Request Date",
        "VIN",
        "State",
        "Dealer Code",
        "Vahan Status",
        "State Backend Status"
    ]

    if os.path.exists(FILE):

        try:
            df = pd.read_excel(FILE)

        except:
            df = pd.DataFrame(
                columns=columns
            )

    else:

        df = pd.DataFrame(
            columns=columns
        )

        df.to_excel(
            FILE,
            index=False
        )

    for col in columns:

        if col not in df.columns:
            df[col] = ""

    return df


df = load_data()

col1, col2, col3 = st.columns(3)

col1.metric(
    "Total Request",
    len(df)
)

col2.metric(
    "Vahan Completed",
    (
        df["Vahan Status"]
        .eq("Complete")
        .sum()
    )
)

col3.metric(
    "State Backend Completed",
    (
        df["State Backend Status"]
        .eq("Completed")
        .sum()
    )
)

st.divider()

st.subheader("Pages")

c1, c2, c3 = st.columns(3)

with c1:
    if st.button("Add Request"):
        st.switch_page(
            "pages/1_Add_Request.py"
        )

with c2:
    if st.button("Vahan Status"):
        st.switch_page(
            "pages/2_Vahan_Status.py"
        )

with c3:
    if st.button("State Backend"):
        st.switch_page(
            "pages/3_State_Backend_Status.py"
        )

if os.path.exists(FILE):

    with open(FILE, "rb") as f:

        st.download_button(
            "Download Excel",
            f,
            "VLTD_Tagging_Data.xlsx"
        )
import streamlit as st
import pandas as pd
import os
from datetime import datetime

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

st.title("➕ Add Request")

states = [
"Delhi","Haryana","UP","Punjab","Rajasthan","Bihar","MP",
"Maharashtra","Tamil Nadu","Karnataka","Kerala","Gujarat"
]

vin = st.text_input("Enter VIN")
state = st.selectbox("Select State", states)
dealer = st.text_input("Enter Dealer Code")

def load_data():
    if os.path.exists(FILE):
        return pd.read_excel(FILE)
    return pd.DataFrame()

if st.button("Submit"):

    df = load_data()

    new_row = {
        "Request Date": datetime.now(),
        "VIN": vin,
        "State": state,
        "Dealer Code": dealer,
        "Vahan Status": "Pending",
        "Vahan Tagged By": "",
        "Vahan Remarks": "",
        "State Backend Status": "Pending",
        "State Tagged By": "",
        "State Remarks": "",
        "Forward To Lumax": "No"
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(FILE, index=False)

    st.success("Request Added Successfully")

if st.button("Clear"):
    st.rerun()
