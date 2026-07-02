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


ALL_STATES = [

"Andhra Pradesh",
"Arunachal Pradesh",
"Assam",
"Bihar",
"Chhattisgarh",
"Delhi",
"Goa",
"Gujarat",
"Haryana",
"Himachal Pradesh",
"Jharkhand",
"Karnataka",
"Kerala",
"Madhya Pradesh",
"Maharashtra",
"Manipur",
"Meghalaya",
"Mizoram",
"Nagaland",
"Odisha",
"Punjab",
"Rajasthan",
"Sikkim",
"Tamil Nadu",
"Telangana",
"Tripura",
"Uttar Pradesh",
"Uttarakhand",
"West Bengal",

"Andaman and Nicobar",
"Chandigarh",
"Dadra and Nagar Haveli",
"Daman and Diu",
"Jammu and Kashmir",
"Ladakh",
"Lakshadweep",
"Puducherry"

]


def load_data():

    if os.path.exists(FILE):

        df = pd.read_excel(
            FILE,
            dtype=str
        )

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

    return df.fillna("")



def save_data(df):

    df.to_excel(
        FILE,
        index=False,
        engine="openpyxl"
    )



st.set_page_config(
page_title="VLTD",
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

# ==================================
# DASHBOARD
# ==================================

if page == "Dashboard":

    st.title("📊 VLTD Dashboard")

    temp = df.copy()

    if len(temp):

        temp["Request Date"] = pd.to_datetime(
            temp["Request Date"],
            errors="coerce"
        )

        cdate1, cdate2 = st.columns(2)

        with cdate1:

            from_date = st.date_input(
                "From Date"
            )

        with cdate2:

            to_date = st.date_input(
                "To Date"
            )

        temp = temp[
            (
                temp[
                    "Request Date"
                ]
                .dt.date
                >= from_date
            )
            &
            (
                temp[
                    "Request Date"
                ]
                .dt.date
                <= to_date
            )
        ]


    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Total Request",
        len(temp)
    )

    c2.metric(
        "Total Vahan Tagging",
        (
            temp[
                "Vahan Status"
            ]
            ==
            "Complete"
        ).sum()
    )

    c3.metric(
        "Total State Backend Tagging",
        (
            temp[
                "State Backend Status"
            ]
            ==
            "Completed"
        ).sum()
    )


    st.divider()


    st.subheader(
        "State Wise Tagging"
    )

    if len(temp):

        chart = (

            temp.groupby(
                "State"
            )

            .agg(

            {

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

            }

            )

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


    st.subheader(
        "Tagged By"
    )

    if len(temp):

        tag = (

            temp

            .groupby(
                "Vahan Tagged By"
            )

            .size()

            .reset_index(
                name="Count"
            )

        )

        fig2 = px.bar(

            tag,

            x="Vahan Tagged By",

            y="Count"

        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )


    st.subheader(
        "All Requests"
    )

    st.dataframe(
        temp,
        use_container_width=True
    )


    if os.path.exists(FILE):

        st.download_button(

            "⬇ Download Excel",

            open(
                FILE,
                "rb"
            ),

            "VLTD_Tagging_Data.xlsx"

        )

# ==================================
# ADD REQUEST
# ==================================

elif page == "Add Request":

    st.title("➕ Add Request")

    vin = st.text_input(
        "Enter VIN"
    )

    state = st.selectbox(
        "Select State / UT",
        ALL_STATES
    )

    dealer = st.text_input(
        "Enter Dealer Code"
    )

    c1, c2 = st.columns(2)

    with c1:

        if st.button(
            "Submit"
        ):

            vin = vin.strip()
            dealer = dealer.strip()

            if not vin:

                st.error(
                    "VIN required"
                )

            else:

                duplicate = (

                    df[
                        "VIN"
                    ]
                    .astype(str)

                    .str.upper()

                    ==
                    vin.upper()

                ).any()

                if duplicate:

                    st.warning(
                        "VIN already exists"
                    )

                else:

                    row = {

                    "Request Date":
                    str(
                    datetime.now()
                    ),

                    "VIN":
                    vin,

                    "State":
                    state,

                    "Dealer Code":
                    dealer,

                    "Vahan Status":
                    "Pending",

                    "Vahan Tagged By":
                    "",

                    "Vahan Remarks":
                    "",

                    "Vahan Update Time":
                    "",

                    "Forward To Lumax":
                    "No",

                    "Lumax Forward Time":
                    "",

                    "State Backend Status":
                    "Pending",

                    "State Tagged By":
                    "",

                    "State Remarks":
                    "",

                    "State Update Time":
                    ""

                    }

                    df = pd.concat(

                        [

                        df,

                        pd.DataFrame(
                            [row]
                        )

                        ],

                        ignore_index=True

                    )

                    save_data(
                        df
                    )

                    st.success(
                        "Request Added"
                    )


    with c2:

        if st.button(
            "Clear"
        ):

            st.rerun()


    st.divider()

    st.subheader(
        "Recent Requests"
    )

    show = df[

        [

        "Request Date",

        "VIN",

        "State",

        "Dealer Code",

        "Vahan Status"

        ]

    ]

    st.dataframe(

        show.tail(20),

        use_container_width=True

    )

# ==================================
# VAHAN STATUS
# ==================================

# ==================================
# VAHAN STATUS
# ==================================

elif page == "Vahan Status":

    st.title("🚗 Vahan Status")

    pending = df[

        df[
            "Forward To Lumax"
        ]

        .astype(str)

        .str.strip()

        !=

        "Yes"

    ].copy()


    st.subheader(
        "Vahan Requests"
    )


    if len(pending) == 0:

        st.info(
            "No records available"
        )

    else:

        pending = pending.reset_index()

        pending["Select"] = False


        selected = st.data_editor(

            pending[

                [

                "Select",

                "Request Date",

                "VIN",

                "State",

                "Dealer Code",

                "Vahan Status",

                "Vahan Tagged By"

                ]

            ],

            hide_index=True,

            use_container_width=True

        )


        selected_rows = (

            selected[

                selected[
                    "Select"
                ]

            ]

            .index

            .tolist()

        )


        selected_vin = (

            pending

            .loc[
                selected_rows,
                "VIN"
            ]

            .tolist()

        )


        st.write(
            "Selected VIN:",
            len(
                selected_vin
            )
        )


        status = st.selectbox(

            "Vahan Status",

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


        if status == "Pending":

            remarks = st.text_area(
                "Remarks"
            )


        c1, c2 = st.columns(2)


        with c1:

            if st.button(
                "Update Status"
            ):

                if not selected_vin:

                    st.warning(
                        "Select VIN"
                    )

                else:

                    mask = (

                        df[
                            "VIN"
                        ]

                        .isin(
                            selected_vin
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
                        "Vahan Remarks"
                    ] = remarks


                    df.loc[
                        mask,
                        "Vahan Update Time"
                    ] = str(
                        datetime.now()
                    )


                    save_data(
                        df
                    )

                    st.success(
                        "Status Updated"
                    )

                    st.rerun()


        with c2:

            if st.button(
                "Forward To Lumax"
            ):

                if not selected_vin:

                    st.warning(
                        "Select VIN"
                    )

                else:

                    mask = (

                        df[
                            "VIN"
                        ]

                        .isin(
                            selected_vin
                        )

                    )


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


                    df.loc[
                        mask,
                        "Vahan Status"
                    ] = "Complete"


                    save_data(
                        df
                    )

                    st.success(
                        "Forwarded To Lumax"
                    )

                    st.rerun()


        st.divider()


        st.subheader(
            "Current Vahan List"
        )


        st.dataframe(

            pending[

                [

                "Request Date",

                "VIN",

                "State",

                "Dealer Code",

                "Vahan Status",

                "Vahan Tagged By",

                "Vahan Remarks"

                ]

            ],

            use_container_width=True

        )
        # ==================================
# STATE BACKEND STATUS
# ==================================

elif page == "State Backend Status":

    st.title("🏢 State Backend Status")

    pending = df[

        (
            df[
                "Forward To Lumax"
            ]
            .astype(str)
            .str.strip()
            ==
            "Yes"
        )

        &

        (
            df[
                "State Backend Status"
            ]
            .astype(str)
            .str.strip()
            ==
            "Pending"
        )

    ].copy()


    st.subheader(
        "Pending State Backend Requests"
    )


    if len(pending) == 0:

        st.info(
            "No pending VIN"
        )


    else:

        pending = pending.reset_index()

        pending["Select"] = False


        selected_table = st.data_editor(

            pending[

                [

                "Select",

                "Request Date",

                "VIN",

                "State",

                "Dealer Code",

                "Vahan Tagged By",

                "Forward To Lumax"

                ]

            ],

            hide_index=True,

            use_container_width=True

        )


        selected_rows = (

            selected_table[
                selected_table[
                    "Select"
                ]
            ]

            .index

            .tolist()

        )


        selected_vin = (

            pending

            .loc[
                selected_rows,
                "VIN"
            ]

            .tolist()

        )


        st.write(
            "Selected:",
            len(
                selected_vin
            )
        )


        tag = st.selectbox(

            "Tagged By",

            [

            "Rahan",

            "Vishal",

            "Lumax Team"

            ]

        )


        status = st.selectbox(

            "State Status",

            [

            "Pending",

            "Completed"

            ]

        )


        remarks = ""

        if status == "Pending":

            remarks = st.text_area(
                "Remarks"
            )


        if st.button(
            "Update State Status"
        ):

            if not selected_vin:

                st.warning(
                    "Select VIN"
                )

            else:

                mask = (

                    df[
                        "VIN"
                    ]

                    .isin(
                        selected_vin
                    )

                )


                df.loc[
                    mask,
                    "State Backend Status"
                ] = status


                df.loc[
                    mask,
                    "State Tagged By"
                ] = str(
                    tag
                )


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


                save_data(
                    df
                )

                st.success(
                    "Updated"
                )

                st.rerun()


        st.divider()


        st.subheader(
            "Pending List"
        )


        st.dataframe(

            pending[

                [

                "Request Date",

                "VIN",

                "State",

                "Dealer Code",

                "Vahan Tagged By",

                "State Backend Status"

                ]

            ],

            use_container_width=True

        )
