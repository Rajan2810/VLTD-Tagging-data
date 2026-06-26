import streamlit as st
from flask import request, redirect
from flask import url_for
from flask import send_file, flash

from datetime import datetime

import json
import os
import uuid

import pandas as pd
import openpyxl
import pytz
import pymysql


DATA_FILE="tagging_requests.json"
EXCEL_FILE="tagging_requests.xlsx"

IST=pytz.timezone("Asia/Kolkata")

# ================= DATABASE =================
import pymysql


# ================= DATABASE =================

def get_connection():

    try:

        conn = pymysql.connect(
            host="esimproddb.taisys.in",
            user="iconnect_user",
            password="YOUR_PASSWORD",
            database="taisys_connect",
            port=3306,
            connect_timeout=10,
            cursorclass=pymysql.cursors.DictCursor
        )

        return conn

    except Exception as e:

        print("DB Connection Error:", e)

        return None


# ================= SQL SEARCH =================

def search_vin(vin):

    conn = get_connection()

    if not conn:
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
            ) AS manuf_month,

            ar.state,

            dl.contact_person

        FROM vehicle v

        LEFT JOIN device d
            ON v.device_id = d.id

        LEFT JOIN esim e
            ON d.esim_id = e.id

        LEFT JOIN dealer dl
            ON dl.id = d.dealer_id

        LEFT JOIN activation_vin av
            ON av.vin = v.vin

        LEFT JOIN activation_request ar
            ON ar.id = av.activation_request_id

        WHERE TRIM(v.vin)=TRIM(%s)

        LIMIT 1

        """

        cur.execute(
            sql,
            (vin,)
        )

        result = cur.fetchone()

        print("VIN SEARCH RESULT:", result)

        return result

    except Exception as e:

        print("Search Error:", e)

        return None

    finally:

        conn.close()

    return result

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
