import pandas as pd
import os

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
