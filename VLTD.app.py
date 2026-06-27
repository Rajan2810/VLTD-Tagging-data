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
