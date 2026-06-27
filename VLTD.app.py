import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px


FILE = "VLTD_Tagging_Data.xlsx"


COLUMNS = [

"Request Date",
"VIN",
"State",
"Dealer Code",

"Vahan Status",
"Vahan Tagged By",
"Vahan Remarks",
"Vahan Update Time",

"Forward To Lumax",
"Lumax Forward Time",

"State Backend Status",
"State Tagged By",
"State Remarks",
"State Update Time"
]


def load_data():

    if os.path.exists(FILE):

        df = pd.read_excel(FILE)

    else:

        df = pd.DataFrame(
            columns=COLUMNS
        )

        df.to_excel(
            FILE,
            index=False
        )

    for c in COLUMNS:

        if c not in df.columns:
            df[c] = ""

    return df


def save_data(df):

    df.to_excel(
        FILE,
        index=False,
        engine="openpyxl"
    )


st.set_page_config(
    page_title="VLTD Dashboard",
    layout="wide"
)

df = load_data()


page = st.sidebar.radio(
    "Menu",
    [
        "Dashboard",
        "Add Request",
        "Vahan Status",
        "State Backend Status"
    ]
)


# ===================
# DASHBOARD
# ===================

if page == "Dashboard":

    st.title("VLTD Dashboard")

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Total Request",
        len(df)
    )

    c2.metric(
        "Vahan Complete",
        (
            df["Vahan Status"]
            == "Complete"
        ).sum()
    )

    c3.metric(
        "State Complete",
        (
            df["State Backend Status"]
            == "Completed"
        ).sum()
    )


    if len(df):

        chart = (
            df.groupby(
                "State"
            )
            .agg({
                "Vahan Status":
                lambda x:
                (
                    x=="Complete"
                ).sum(),

                "State Backend Status":
                lambda x:
                (
                    x=="Completed"
                ).sum()
            })

            .reset_index()
        )

        fig = px.bar(
            chart,
            x="State",
            y=[
                "Vahan Status",
                "State Backend Status"
            ],
            barmode="group"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )


    st.download_button(
        "Download Excel",
        open(FILE,"rb"),
        "VLTD.xlsx"
    )


# ===================
# ADD REQUEST
# ===================

elif page == "Add Request":

    st.title("Add Request")

    states = [

"Delhi",
"Haryana",
"UP",
"Punjab",
"Rajasthan",
"Maharashtra",
"Gujarat",
"Karnataka",
"Tamil Nadu",
"Kerala"

]

    vin = st.text_input(
        "Enter VIN"
    )

    state = st.selectbox(
        "State",
        states
    )

    dealer = st.text_input(
        "Dealer Code"
    )


    if st.button("Submit"):

        row = {

            "Request Date":
            datetime.now(),

            "VIN":
            vin,

            "State":
            state,

            "Dealer Code":
            dealer,

            "Vahan Status":
            "Pending",

            "Forward To Lumax":
            "No",

            "State Backend Status":
            "Pending"
        }

        df = pd.concat(
            [
                df,
                pd.DataFrame(
                    [row]
                )
            ]
        )

        save_data(df)

        st.success(
            "Added"
        )


# ===================
# VAHAN
# ===================

elif page == "Vahan Status":

    st.title("Vahan Status")

    pending = df[
        df[
            "Vahan Status"
        ]
        ==
        "Pending"
    ]


    selected = st.multiselect(
        "Select VIN",
        pending["VIN"]
    )


    status = st.selectbox(
        "Status",
        [
            "Pending",
            "Complete"
        ]
    )


    tag = st.selectbox(
        "Tagged By",
        [
            "Rahan",
            "Vishal",
            "Lumax Team"
        ]
    )


    remarks = ""

    if status=="Pending":

        remarks = st.text_area(
            "Remarks"
        )


    if st.button(
        "Update"
    ):

        mask = (
            df[
            "VIN"
            ]
            .isin(
            selected
            )
        )


        df.loc[
            mask,
            "Vahan Status"
        ] = status


        df.loc[
            mask,
            "Vahan Tagged By"
        ] = tag


        df.loc[
            mask,
            "Vahan Update Time"
        ] = str(
            datetime.now()
        )


        if status=="Pending":

            df.loc[
                mask,
                "Vahan Remarks"
            ] = remarks

        else:

            df.loc[
                mask,
                "Forward To Lumax"
            ] = "Yes"


            df.loc[
                mask,
                "Lumax Forward Time"
            ] = str(
                datetime.now()
            )

        save_data(df)

        st.success(
            "Updated"
        )


# ===================
# STATE
# ===================

else:

    st.title(
        "State Backend"
    )

    pending = df[

    (
        df[
        "Forward To Lumax"
        ]
        ==
        "Yes"
    )

    &

    (
        df[
        "State Backend Status"
        ]
        ==
        "Pending"
    )

    ]


    selected = st.multiselect(
        "VIN",
        pending["VIN"]
    )


    status = st.selectbox(
        "Status",
        [
            "Pending",
            "Completed"
        ]
    )


    tag = st.selectbox(
        "Tagged By",
        [
            "Rahan",
            "Vishal",
            "Lumax Team"
        ]
    )


    remarks = ""

    if status=="Pending":

        remarks = st.text_area(
            "Remarks"
        )


    if st.button(
        "Update"
    ):

        mask = (
            df[
            "VIN"
            ]
            .isin(
            selected
            )
        )

        df.loc[
            mask,
            "State Backend Status"
        ] = status


        df.loc[
            mask,
            "State Tagged By"
        ] = tag


        df.loc[
            mask,
            "State Remarks"
        ] = remarks


        df.loc[
            mask,
            "State Update Time"
        ] = str(
            datetime.now()
        )

        save_data(df)

        st.success(
            "Updated"
        )
