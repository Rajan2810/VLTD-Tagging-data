import streamlit as st
import pandas as pd
import pymysql
import pytz
import json
import os
from datetime import datetime


# ================= CONFIG =================

DATA_FILE = "tagging_requests.json"
EXCEL_FILE = "tagging_requests.xlsx"

IST = pytz.timezone("Asia/Kolkata")


# ================= DATABASE =================

def get_connection():

    try:

        conn = pymysql.connect(
            host="esimproddb.taisys.in",
            user="iconnect_user",
            password="kG7TwbkkSGZd86mX",   # Replace
            database="taisys_connect",
            port=3306,
            connect_timeout=20,
            cursorclass=pymysql.cursors.DictCursor
        )

        return conn

    except Exception as e:

        st.error(f"DB Error: {e}")

        return None


# ================= SEARCH VIN =================

def search_vin(vin):

    conn = get_connection()

    if conn is None:
        return None

    try:

        cur = conn.cursor()

        sql = """

SELECT

v.vin,

d.esn AS unique_device_code,

d.imei,

e.primary_iccid AS iccid,

DATE_FORMAT(
d.created_on,
'%m/%Y'
) manuf_month,

ar.state,

dl.contact_person

FROM vehicle v

LEFT JOIN device d
ON d.id=v.device_id

LEFT JOIN esim e
ON e.id=d.esim_id

LEFT JOIN dealer dl
ON dl.id=d.dealer_id

LEFT JOIN activation_vin av
ON av.vin=v.vin

LEFT JOIN activation_request ar
ON ar.id=av.activation_request_id

WHERE
TRIM(v.vin)=TRIM(%s)

LIMIT 1

"""

        cur.execute(
            sql,
            (vin,)
        )

        result = cur.fetchone()

        return result

    except Exception as e:

        st.error(
            f"Search Error: {e}"
        )

        return None

    finally:

        conn.close()


# ================= FILE =================

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


# ================= UI =================

st.set_page_config(
    page_title="VLTD Tracking",
    layout="wide"
)

st.title(
    "VLTD Tagging Data Tracking"
)

vin = st.text_input(
    "Enter VIN"
)


if st.button(
    "Search"
):

    if not vin:

        st.warning(
            "Enter VIN"
        )

    else:

        record = search_vin(vin)

        if record is None:

            st.error(
                "VIN NOT MAPPED"
            )

        else:

            st.success(
                "VIN FOUND"
            )

            output = (

                "LIT1|"

                f"{record['unique_device_code']}|"

                f"{record['imei']}|"

                f"{record['iccid']}|"

                f"{record['manuf_month']}|"

                "214"

            )

            st.subheader(
                "Output"
            )

            st.code(
                output
            )

            c1, c2 = st.columns(2)

            with c1:

                st.write(
                    "Dealer",
                    record.get(
                        "contact_person",
                        "-"
                    )
                )

            with c2:

                st.write(
                    "State",
                    record.get(
                        "state",
                        "-"
                    )
                )

            state = st.selectbox(

                "Select State",

                [

                    "Haryana",

                    "Rajasthan",

                    "UP",

                    "Delhi",

                    "Punjab"

                ]

            )

            if st.button(
                "Submit"
            ):

                data = load_data()

                data.append({

                    "VIN":
                    record["vin"],

                    "Unique Device":

                    record[
                        "unique_device_code"
                    ],

                    "IMEI":

                    record[
                        "imei"
                    ],

                    "ICCID":

                    record[
                        "iccid"
                    ],

                    "MFG":

                    record[
                        "manuf_month"
                    ],

                    "Selected State":

                    state,

                    "Request Date":

                    datetime.now(
                        IST
                    ).strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),

                    "Vahan Status":

                    "Pending"

                })

                save_data(
                    data
                )

                st.success(
                    "Request Saved"
                )

                st.dataframe(
                    pd.DataFrame(
                        data
                    )
                )


# ================= DOWNLOAD =================

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

            file_name=
            EXCEL_FILE
        )
