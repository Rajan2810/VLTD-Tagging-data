import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import pytz

DATA_FILE="tagging_requests.json"
EXCEL_FILE="tagging_requests.xlsx"

IST=pytz.timezone("Asia/Kolkata")

# ================= DATABASE =================

DB = pymysql.connect(
    host="esimproddb.taisys.in",
    user="iconnect_user",
    password="kG7TwbkkSGZd86mX",
    database="taisys_connect",
    cursorclass=pymysql.cursors.DictCursor
)


# ================= FILE =================

def load_data():

    if os.path.exists(DATA_FILE):

        with open(DATA_FILE, "r") as f:
            return json.load(f)

    return []


def save_data(data):

    with open(DATA_FILE, "w") as f:

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


# ================= SQL SEARCH =================

def search_vin(vin):

    sql = """

WITH ranked_messages AS (

SELECT
m.*,

ROW_NUMBER() OVER(
PARTITION BY m.esim_id
ORDER BY id DESC
) rn

FROM esim_lifecycle m
),

ranked_lifecycle AS (

SELECT *

FROM ranked_messages

WHERE rn=1

)

SELECT

ar.request_id,

DATE_FORMAT(
ar.request_date,
'%d/%m/%Y'
) request_date,

dl.contact_person dealer_name,

ar.state,

d.esn unique_device_code,

d.imei,

e.primary_iccid iccid,

DATE_FORMAT(
d.created_on,
'%m/%Y'
) manuf_month,

v.vin,

v.engine_number,

RIGHT(
v.engine_number,
5
) engine_last_5

FROM esim e

LEFT JOIN ranked_lifecycle el
ON el.esim_id=e.id

LEFT JOIN device d
ON d.esim_id=e.id

LEFT JOIN dealer dl
ON dl.id=d.dealer_id

LEFT JOIN vehicle v
ON v.device_id=d.id

LEFT JOIN activation_vin av
ON av.vin=v.vin

LEFT JOIN activation_request ar
ON ar.id=av.activation_request_id

WHERE
el.rn=1
AND
v.vin=%s

LIMIT 1

"""

    cur = DB.cursor()

    cur.execute(
        sql,
        (vin,)
    )

    return cur.fetchone()


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
