import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import plotly.express as px

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="VLTD System", layout="wide")

# =========================
# NAVIGATION
# =========================
page = st.sidebar.selectbox(
    "Select Page",
    ["Dashboard", "Add Request", "Vahan Status", "State Backend Status"]
)

# =========================
# GOOGLE SHEET CONFIG
# =========================
SHEET_ID = st.secrets["1BNu9Do5Hz6DIqyTsRPsPOieBMy8bbXz2_bG8r7z_VTI"]

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
def get_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )
    return gspread.authorize(creds)

# =========================
# COLUMNS
# =========================
COLUMNS = [
    "Request Date", "VIN", "State", "Dealer Code",
    "Vahan Status", "Vahan Tagged By", "Vahan Remarks", "Vahan Update Time",
    "Forward To Lumax", "Lumax Forward Time",
    "State Backend Status", "State Tagged By", "State Remarks", "State Update Time"
]

# =========================
# LOAD DATA
# =========================
def load_data():
    try:
        client = get_client()
        sheet = client.open_by_key(SHEET_ID).sheet1
        data = sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame(columns=COLUMNS)

df = load_data()

for c in COLUMNS:
    if c not in df.columns:
        df[c] = ""

df = df.fillna("")

# =========================
# SAVE DATA
# =========================
def save_data(df):
    try:
        client = get_client()
        sheet = client.open_by_key(SHEET_ID).sheet1
        sheet.clear()

        data = [df.columns.tolist()] + df.astype(str).values.tolist()
        sheet.update(data)

    except Exception as e:
        st.error(f"Save Error: {e}")

# =========================
# STATES
# =========================
ALL_STATES = [
"Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh",
"Goa","Gujarat","Haryana","Himachal Pradesh","Jharkhand","Karnataka",
"Kerala","Madhya Pradesh","Maharashtra","Manipur","Meghalaya",
"Mizoram","Nagaland","Odisha","Punjab","Rajasthan","Sikkim",
"Tamil Nadu","Telangana","Tripura","Uttar Pradesh","Uttarakhand",
"West Bengal","Delhi","Jammu and Kashmir","Ladakh","Puducherry"
]

# =========================
# DASHBOARD
# =========================
if page == "Dashboard":

    st.title("📊 VLTD Dashboard")

    temp = df.copy()

    if len(temp):
        temp["Request Date"] = pd.to_datetime(temp["Request Date"], errors="coerce")

        col1, col2 = st.columns(2)
        with col1:
            from_date = st.date_input("From Date")
        with col2:
            to_date = st.date_input("To Date")

        temp = temp[
            (temp["Request Date"].dt.date >= from_date) &
            (temp["Request Date"].dt.date <= to_date)
        ]

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Requests", len(temp))
    c2.metric("Vahan Completed", (temp["Vahan Status"] == "Complete").sum())
    c3.metric("State Completed", (temp["State Backend Status"] == "Completed").sum())

    st.divider()

    st.subheader("State Wise Chart")

    if len(temp):
        chart = temp.groupby("State").agg({
            "VIN": "count",
            "Vahan Status": lambda x: (x == "Complete").sum(),
            "State Backend Status": lambda x: (x == "Completed").sum()
        }).reset_index()

        fig = px.bar(chart, x="State",
                     y=["Vahan Status", "State Backend Status"],
                     barmode="group")

        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(temp, use_container_width=True)

    st.download_button(
        "⬇ Download CSV",
        temp.to_csv(index=False).encode("utf-8"),
        "VLTD_Data.csv",
        "text/csv"
    )

# =========================
# ADD REQUEST
# =========================
elif page == "Add Request":

    st.title("➕ Add Request")

    vin = st.text_input("Enter VIN")
    state = st.selectbox("Select State", ALL_STATES)
    dealer = st.text_input("Enter Dealer Code")

    if st.button("Submit"):

        if vin:

            new_row = {
                "Request Date": str(datetime.now()),
                "VIN": vin,
                "State": state,
                "Dealer Code": dealer,
                "Vahan Status": "Pending",
                "Vahan Tagged By": "",
                "Vahan Remarks": "",
                "Vahan Update Time": "",
                "Forward To Lumax": "No",
                "Lumax Forward Time": "",
                "State Backend Status": "Pending",
                "State Tagged By": "",
                "State Remarks": "",
                "State Update Time": ""
            }

            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df)

            st.success("Request Added")

        else:
            st.error("VIN required")

    st.dataframe(df.tail(20), use_container_width=True)

# =========================
# VAHAN STATUS
# =========================
elif page == "Vahan Status":

    st.title("🚗 Vahan Status")

    pending = df[df["Forward To Lumax"].astype(str) != "Yes"].copy()

    st.dataframe(pending, use_container_width=True)

    selected_vin = st.multiselect("Select VIN", pending["VIN"].tolist())

    status = st.selectbox("Status", ["Pending", "Complete"])
    tag = st.selectbox("Tagged By", ["Rajan", "Vishal", "Lumax Team"])

    remarks = st.text_area("Remarks") if status == "Pending" else ""

    if st.button("Update"):

        mask = df["VIN"].isin(selected_vin)

        df.loc[mask, "Vahan Status"] = status
        df.loc[mask, "Vahan Tagged By"] = tag
        df.loc[mask, "Vahan Remarks"] = remarks
        df.loc[mask, "Vahan Update Time"] = str(datetime.now())

        save_data(df)
        st.success("Updated")

    if st.button("Forward To Lumax"):

        mask = df["VIN"].isin(selected_vin)

        df.loc[mask, "Forward To Lumax"] = "Yes"
        df.loc[mask, "Lumax Forward Time"] = str(datetime.now())
        df.loc[mask, "Vahan Status"] = "Complete"

        save_data(df)
        st.success("Forwarded")

# =========================
# STATE BACKEND
# =========================
elif page == "State Backend Status":

    st.title("🏢 State Backend")

    pending = df[
        (df["Forward To Lumax"].astype(str) == "Yes") &
        (df["State Backend Status"].astype(str) == "Pending")
    ].copy()

    st.dataframe(pending, use_container_width=True)

    selected_vin = st.multiselect("Select VIN", pending["VIN"].tolist())

    status = st.selectbox("Status", ["Pending", "Completed"])
    tag = st.selectbox("Tagged By", ["Rajan", "Vishal", "Lumax Team"])

    remarks = st.text_area("Remarks") if status == "Pending" else ""

    if st.button("Update State"):

        mask = df["VIN"].isin(selected_vin)

        df.loc[mask, "State Backend Status"] = status
        df.loc[mask, "State Tagged By"] = tag
        df.loc[mask, "State Remarks"] = remarks
        df.loc[mask, "State Update Time"] = str(datetime.now())

        save_data(df)
        st.success("Updated")
