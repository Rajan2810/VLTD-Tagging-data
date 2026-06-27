import streamlit as st
import pandas as pd
import os
import plotly.express as px

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

st.set_page_config(page_title="VLTD Dashboard", layout="wide")
st.title("📊 VLTD Dashboard")

def load_data():
    if os.path.exists(FILE):
        return pd.read_excel(FILE)
    return pd.DataFrame()

df = load_data()

# DATE FILTER
if not df.empty:
    df["Request Date"] = pd.to_datetime(df["Request Date"], errors="coerce")

    start = st.sidebar.date_input("From", df["Request Date"].min())
    end = st.sidebar.date_input("To", df["Request Date"].max())

    df = df[(df["Request Date"].dt.date >= start) &
            (df["Request Date"].dt.date <= end)]

# METRICS
c1, c2, c3 = st.columns(3)

c1.metric("Total Request", len(df))
c2.metric("Vahan Completed", df["Vahan Status"].eq("Complete").sum())
c3.metric("State Backend Completed", df["State Backend Status"].eq("Completed").sum())

# STATE WISE CHART
st.subheader("📍 State Wise Tagging")

if not df.empty:
    chart = df.groupby("State").agg({
        "VIN": "count",
        "Vahan Status": lambda x: (x == "Complete").sum(),
        "State Backend Status": lambda x: (x == "Completed").sum()
    }).reset_index()

    fig = px.bar(chart, x="State", y=["Vahan Status", "State Backend Status"],
                 barmode="group")

    st.plotly_chart(fig, use_container_width=True)

# TAGGED BY CHART
st.subheader("👤 Tagged By Person")

if not df.empty:
    fig2 = px.bar(df.groupby("Vahan Tagged By").size().reset_index(name="Count"),
                  x="Vahan Tagged By", y="Count")
    st.plotly_chart(fig2, use_container_width=True)

# DOWNLOAD
st.download_button("⬇ Download Excel", data=open(FILE, "rb"), file_name="VLTD.xlsx")

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

import streamlit as st
import pandas as pd
from datetime import datetime
import os

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

st.title("🚗 Vahan Status")

def load_data():
    return pd.read_excel(FILE)

df = load_data()

pending_df = df[df["Vahan Status"] == "Pending"]

selected = st.multiselect("Select VIN", pending_df["VIN"].tolist())

status = st.selectbox("Vahan Status", ["Pending", "Complete"])

tagged_by = st.selectbox("Tagged By", ["Rahan", "Vishal", "Lumax Team"])

remarks = ""

if status == "Pending":
    remarks = st.text_area("Remarks")

if st.button("Update"):

    df.loc[df["VIN"].isin(selected), "Vahan Status"] = status
    df.loc[df["VIN"].isin(selected), "Vahan Tagged By"] = tagged_by

    if status == "Pending":
        df.loc[df["VIN"].isin(selected), "Vahan Remarks"] = remarks
    else:
        df.loc[df["VIN"].isin(selected), "Forward To Lumax"] = "Yes"
        df.loc[df["VIN"].isin(selected), "Lumax Forward Time"] = str(datetime.now())

    df.to_excel(FILE, index=False)

    st.success("Updated Successfully")

st.dataframe(pending_df)


import streamlit as st
import pandas as pd
from datetime import datetime

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

st.title("🏢 State Backend Status")

df = pd.read_excel(FILE)

df = df[df["Forward To Lumax"] == "Yes"]

pending = df[df["State Backend Status"] == "Pending"]

selected = st.multiselect("Select VIN", pending["VIN"].tolist())

tagged_by = st.selectbox("Tagged By", ["Rahan", "Vishal", "Lumax Team"])

status = st.selectbox("Status", ["Pending", "Completed"])

remarks = ""

if status == "Pending":
    remarks = st.text_area("Remarks")

if st.button("Update"):

    df_all = pd.read_excel(FILE)

    df_all.loc[df_all["VIN"].isin(selected), "State Backend Status"] = status
    df_all.loc[df_all["VIN"].isin(selected), "State Tagged By"] = tagged_by

    if status == "Pending":
        df_all.loc[df_all["VIN"].isin(selected), "State Remarks"] = remarks
    else:
        df_all.loc[df_all["VIN"].isin(selected), "State Update Time"] = str(datetime.now())

    df_all.to_excel(FILE, index=False)

    st.success("Updated Successfully")

st.dataframe(pending)
