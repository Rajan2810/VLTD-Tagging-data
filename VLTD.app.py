# utils.py

import pandas as pd
import os

FILE = r"D:\OneDrive - 太思科技股份有限公司\VLTD tagging Data.xlsx"

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

```
if os.path.exists(FILE):

    df = pd.read_excel(FILE)

else:

    df = pd.DataFrame(columns=COLUMNS)

    df.to_excel(FILE,index=False)

for c in COLUMNS:

    if c not in df.columns:

        df[c]=""

return df
```

def save_data(df):

```
df.to_excel(
    FILE,
    index=False,
    engine="openpyxl"
)
```

# =====================================

# app.py

# =====================================

import streamlit as st
import plotly.express as px
from utils import load_data, FILE

st.set_page_config(
page_title="VLTD Dashboard",
layout="wide"
)

df=load_data()

st.title("VLTD Dashboard")

col1,col2,col3=st.columns(3)

col1.metric(
"Total Request",
len(df)
)

col2.metric(
"Vahan Complete",
(
df["Vahan Status"]
.eq("Complete")
.sum()
)
)

col3.metric(
"State Complete",
(
df["State Backend Status"]
.eq("Completed")
.sum()
)
)

st.divider()

st.subheader("Open Pages")

c1,c2,c3=st.columns(3)

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

if len(df)>0:

```
chart=(
    df.groupby(
    "State"
    )
    .agg(
    {
    "Vahan Status":
    lambda x:
    (x=="Complete").sum(),

    "State Backend Status":
    lambda x:
    (x=="Completed").sum()
    }
    )
    .reset_index()
)

fig=px.bar(
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
```

st.download_button(
"Download Excel",
open(FILE,"rb"),
"VLTD_Data.xlsx"
)

# =====================================

# pages/1_Add_Request.py

# =====================================

import streamlit as st
import pandas as pd
from datetime import datetime
from utils import load_data,save_data

states=[

"Andhra Pradesh",
"Arunachal Pradesh",
"Assam",
"Bihar",
"Chhattisgarh",
"Delhi",
"Gujarat",
"Haryana",
"Himachal Pradesh",
"Jharkhand",
"Karnataka",
"Kerala",
"Madhya Pradesh",
"Maharashtra",
"Odisha",
"Punjab",
"Rajasthan",
"Tamil Nadu",
"Telangana",
"UP",
"Uttarakhand",
"West Bengal"

]

st.title("Add Request")

vin=st.text_input(
"Enter VIN"
)

state=st.selectbox(
"State",
states
)

dealer=st.text_input(
"Dealer Code"
)

if st.button(
"Submit"
):

```
df=load_data()

row={

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

"State Backend Status":
"Pending",

"Forward To Lumax":
"No"
}

df=pd.concat(
[
df,
pd.DataFrame([row])
]
)

save_data(df)

st.success(
"Request Added"
)
```

if st.button(
"Dashboard"
):

```
st.switch_page(
"app.py"
)
```

# =====================================

# pages/2_Vahan_Status.py

# =====================================

import streamlit as st
from datetime import datetime
from utils import load_data,save_data

st.title("Vahan Status")

df=load_data()

df=df[
df[
"Vahan Status"
]=="Pending"
]

selected=st.multiselect(
"Select VIN",
df["VIN"]
)

status=st.selectbox(
"Status",
[
"Pending",
"Complete"
]
)

tag=st.selectbox(
"Tagged By",
[
"Rahan",
"Vishal",
"Lumax Team"
]
)

remarks=""

if status=="Pending":

```
remarks=st.text_area(
"Remarks"
)
```

if st.button(
"Update"
):

```
full=load_data()

mask=(
full["VIN"]
.isin(selected)
)

full.loc[
mask,
"Vahan Status"
]=status

full.loc[
mask,
"Vahan Tagged By"
]=tag

full.loc[
mask,
"Vahan Update Time"
]=str(
datetime.now()
)

if status=="Pending":

    full.loc[
    mask,
    "Vahan Remarks"
    ]=remarks

else:

    full.loc[
    mask,
    "Forward To Lumax"
    ]="Yes"

    full.loc[
    mask,
    "Lumax Forward Time"
    ]=str(
    datetime.now()
    )

save_data(full)

st.success(
"Updated"
)
```

# =====================================

# pages/3_State_Backend_Status.py

# =====================================

import streamlit as st
from datetime import datetime
from utils import load_data,save_data

st.title(
"State Backend"
)

df=load_data()

df=df[
(
df[
"Forward To Lumax"
]=="Yes"
)
&
(
df[
"State Backend Status"
]=="Pending"
)
]

selected=st.multiselect(
"Select VIN",
df["VIN"]
)

tag=st.selectbox(
"Tagged By",
[
"Rahan",
"Vishal",
"Lumax Team"
]
)

status=st.selectbox(
"Status",
[
"Pending",
"Completed"
]
)

remarks=""

if status=="Pending":

```
remarks=st.text_area(
"Remarks"
)
```

if st.button(
"Update"
):

```
full=load_data()

mask=(
full["VIN"]
.isin(selected)
)

full.loc[
mask,
"State Backend Status"
]=status

full.loc[
mask,
"State Tagged By"
]=tag

full.loc[
mask,
"State Remarks"
]=remarks

full.loc[
mask,
"State Update Time"
]=str(
datetime.now()
)

save_data(full)

st.success(
"Updated"
)
```
