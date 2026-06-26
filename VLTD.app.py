import streamlit as st
from datetime import datetime
import json, os, uuid
import pandas as pd
import openpyxl
import pytz

app = Flask(__name__)
app.secret_key = "secret"  # for flash messages

DATA_FILE = "tagging_requests.json"
EXCEL_FILE = "tagging_requests.xlsx"
TMP_FOLDER = os.path.abspath("tmp")  # absolute path for temp files
SAMPLE_FILE = os.path.abspath("static/sample_bulk_upload.xlsx")
IST = pytz.timezone('Asia/Kolkata')

os.makedirs(TMP_FOLDER, exist_ok=True)


# ---------------- Utility Functions ----------------
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return []


def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=4, default=str)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tagging Requests"

    headers = [
        "ID", "VIN", "State", "Dealer Code", "Request Date",
        "Vahan Status", "Forwarded to Lumax", "Forwarded Time",
        "Remarks", "Tagging Status", "Closure Date"
    ]
    ws.append(headers)

    for r in data:
        ws.append([
            r.get("id"),
            r.get("vin"),
            r.get("state"),
            r.get("dealer_code"),
            r.get("request_date"),
            r.get("vahan_status"),
            "Yes" if r.get("forwarded_to_lumax") else "No",
            r.get("forwarded_time"),
            r.get("remarks"),
            r.get("tagging_status"),
            r.get("closure_date"),
        ])

    wb.save(EXCEL_FILE)


# ---------------- Page 1: Add Request ----------------
@app.route('/')
def index():
    return render_template('add_request.html')


@app.route('/add', methods=['POST'])
def add():
    requests = load_data()
    new_request = {
        "id": len(requests) + 1,
        "vin": request.form['vin'],
        "state": request.form['state'],
        "dealer_code": request.form['dealer_code'],
        "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
        "vahan_status": "Pending",
        "forwarded_to_lumax": False,
        "forwarded_time": None,
        "tagging_status": None,
        "closure_date": None,
        "remarks": None
    }
    requests.append(new_request)
    save_data(requests)
    flash("Request added successfully.", "success")
    return redirect(url_for('index'))


# ---------------- Bulk Upload ----------------
@app.route("/bulk_upload", methods=["POST"])
def bulk_upload():
    file = request.files.get("file")
    if not file or file.filename == "":
        flash("No file selected.", "danger")
        return redirect(url_for("index"))

    tmp_filename = os.path.join(TMP_FOLDER, f"{uuid.uuid4().hex}_{file.filename}")
    file.save(tmp_filename)

    # Read into pandas for preview
    if tmp_filename.endswith(".csv"):
        df = pd.read_csv(tmp_filename)
    else:
        df = pd.read_excel(tmp_filename)

    required_cols = ["VIN", "State", "Dealer Code"]
    if not all(col in df.columns for col in required_cols):
        flash(f"File must contain columns: {', '.join(required_cols)}", "danger")
        os.remove(tmp_filename)
        return redirect(url_for("index"))

    preview_html = df.to_html(classes="table table-bordered table-sm table-striped", index=False)
    return render_template("preview.html", table=preview_html, tmp_filename=tmp_filename)


@app.route("/confirm_upload", methods=["POST"])
def confirm_upload():
    tmp_filename = request.form.get("tmp_filename")
    if not tmp_filename or not os.path.exists(tmp_filename):
        flash("Temporary file not found. Please upload again.", "danger")
        return redirect(url_for("index"))

    if tmp_filename.endswith(".csv"):
        df = pd.read_csv(tmp_filename)
    else:
        df = pd.read_excel(tmp_filename)

    requests = load_data()
    for _, row in df.iterrows():
        new_request = {
            "id": len(requests) + 1,
            "vin": row["VIN"],
            "state": row["State"],
            "dealer_code": row["Dealer Code"],
            "request_date": datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S"),
            "vahan_status": "Pending",
            "forwarded_to_lumax": False,
            "forwarded_time": None,
            "tagging_status": None,
            "closure_date": None,
            "remarks": None
        }
        requests.append(new_request)

    save_data(requests)
    os.remove(tmp_filename)
    flash("Bulk upload successful.", "success")
    return redirect(url_for("index"))


# ---------------- Page 2: Vahan Status ----------------
@app.route('/vahan', methods=['GET', 'POST'])
def vahan_status():
    data = load_data()

    if request.method == "POST":
        req_id = int(request.form.get("id"))
        for r in data:
            if r["id"] == req_id:
                r["vahan_status"] = request.form.get("vahan_status")
                r["remarks"] = request.form.get("remarks")
        save_data(data)
        flash("Vahan status updated.", "success")
        return redirect(url_for("vahan_status"))

    visible_requests = [
        r for r in data
        if r["vahan_status"] in ["Pending", "Done"] and not r.get("forwarded_to_lumax", False)
    ]
    return render_template("vahan_status.html", requests=visible_requests)


# ---------------- Forward to Lumax ----------------
@app.route('/forward/<int:request_id>')
def forward_to_lumax(request_id):
    data = load_data()
    for r in data:
        if r["id"] == request_id and r["vahan_status"] == "Done":
            r["forwarded_to_lumax"] = True
            r["forwarded_time"] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
            flash("Forwarded to Lumax.", "success")
    save_data(data)
    return redirect(url_for("vahan_status"))


# ---------------- Page 3: Backend / Tagging Status ----------------
@app.route('/backend', methods=['GET', 'POST'])
def backend_status():
    data = load_data()

    if request.method == "POST":
        req_id = int(request.form.get("id"))
        for r in data:
            if r["id"] == req_id:
                r['tagging_status'] = request.form.get('tagging_status')
                if r['tagging_status'] == "Pending":
                    r['remarks'] = request.form.get('remarks')
                elif r['tagging_status'] == "Completed":
                    r['closure_date'] = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
        save_data(data)
        flash("Backend status updated.", "success")
        return redirect(url_for("backend_status"))

    visible_requests = [
        r for r in data
        if r.get('forwarded_to_lumax') and (r.get('tagging_status') == "Pending" or r.get('tagging_status') is None)
    ]
    return render_template('backend_status.html', requests=visible_requests)


# ---------------- Download Excel ----------------
@app.route('/download')
def download():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    flash("Excel file not found.", "danger")
    return redirect(url_for("index"))


# ---------------- Sample Format Download ----------------
@app.route("/sample_format")
def sample_format():
    if os.path.exists(SAMPLE_FILE):
        return send_file(SAMPLE_FILE, as_attachment=True)
    flash("Sample file not found.", "danger")
    return redirect(url_for("index"))


# ---------------- Run App ----------------
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
    
