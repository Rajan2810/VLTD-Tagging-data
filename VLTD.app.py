import streamlit as st
from flask import Flask, render_template, request, redirect, url_for
from flask import send_file, flash

import pymysql
import pandas as pd
import openpyxl
import pytz

from datetime import datetime
import os
import json

app = Flask(__name__)
app.secret_key = "secret"

DATA_FILE = "tracking.json"
EXCEL_FILE = "tracking.xlsx"

IST = pytz.timezone("Asia/Kolkata")


# ================= DATABASE =================

DB = pymysql.connect(
    host="YOUR_HOST",
    user="YOUR_USER",
    password="YOUR_PASSWORD",
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


# ================= PAGE 1 =================

@app.route(
"/",
methods=[
"GET",
"POST"
]
)

def home():

    output = None

    error = None

    record = None

    if request.method == "POST":

        vin = request.form["vin"]

        record = search_vin(vin)

        if not record:

            error = "VIN NOT MAPPED"

        else:

            output = (
                f"LIT1|"
                f"{record['unique_device_code']}|"
                f"{record['imei']}|"
                f"{record['iccid']}|"
                f"{record['manuf_month']}|"
                f"214"
            )

        return render_template(
            "home.html",
            output=output,
            record=record,
            error=error
        )

    return render_template(
        "home.html"
    )


@app.route(
"/submit",
methods=["POST"]
)

def submit():

    data = load_data()

    vin = request.form["vin"]

    state = request.form["state"]

    sql = search_vin(vin)

    obj = {

        "id": len(data) + 1,

        "vin": vin,

        "state": state,

        "unique_device":

        sql[
            "unique_device_code"
        ],

        "imei":

        sql[
            "imei"
        ],

        "iccid":

        sql[
            "iccid"
        ],

        "mfg":

        sql[
            "manuf_month"
        ],

        "request_date":

        datetime.now(
            IST
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        ),

        "tagged_by": "",

        "vahan_status": "Pending",

        "remarks": "",

        "forwarded": False,

        "forward_time": "",

        "backend_status": "",

        "closure_date": ""

    }

    data.append(obj)

    save_data(data)

    return redirect(
        url_for(
            "vahan"
        )
    )


# ================= PAGE 2 =================

@app.route(
"/vahan",
methods=[
"GET",
"POST"
]
)

def vahan():

    data = load_data()

    if request.method == "POST":

        rid = int(
            request.form["id"]
        )

        for r in data:

            if r["id"] == rid:

                r[
                    "tagged_by"
                ] = request.form[
                    "tagged_by"
                ]

                r[
                    "vahan_status"
                ] = request.form[
                    "status"
                ]

                if (
                    r[
                        "vahan_status"
                    ]
                    ==
                    "Pending"
                ):

                    r[
                        "remarks"
                    ] = request.form[
                        "remarks"
                    ]

                else:

                    r[
                        "vahan_complete"
                    ] = (
                        datetime.now(
                            IST
                        ).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    )

        save_data(data)

        flash(
            "Updated"
        )

        return redirect(
            "/vahan"
        )

    return render_template(
        "vahan.html",
        data=data
    )


# ================= FORWARD =================

@app.route(
"/forward/<int:id>"
)

def forward(id):

    data = load_data()

    for r in data:

        if (

            r["id"]

            ==

            id

            and

            r[
                "vahan_status"
            ]

            ==

            "Completed"

        ):

            r[
                "forwarded"
            ] = True

            r[
                "forward_time"
            ] = (

                datetime.now(
                    IST
                )

                .strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

            )

    save_data(data)

    return redirect(
        "/backend"
    )


# ================= PAGE 3 =================

@app.route(
"/backend",
methods=[
"GET",
"POST"
]
)

def backend():

    data = load_data()

    if request.method == "POST":

        rid = int(
            request.form["id"]
        )

        for r in data:

            if r["id"] == rid:

                r[
                    "backend_status"
                ] = request.form[
                    "status"
                ]

                r[
                    "tagged_by"
                ] = request.form[
                    "tagged_by"
                ]

                if (

                    r[
                        "backend_status"
                    ]

                    ==

                    "Pending"

                ):

                    r[
                        "remarks"
                    ] = request.form[
                        "remarks"
                    ]

                else:

                    r[
                        "closure_date"
                    ] = (

                        datetime.now(
                            IST
                        )

                        .strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )

                    )

        save_data(data)

        flash(
            "Completed"
        )

        return redirect(
            "/backend"
        )

    data = [

        x

        for x

        in data

        if x[
            "forwarded"
        ]

    ]

    return render_template(
        "backend.html",
        data=data
    )


# ================= DOWNLOAD =================

@app.route(
"/download"
)

def download():

    return send_file(
        EXCEL_FILE,
        as_attachment=True
    )


# ================= RUN =================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=8080,
        debug=True
    )
