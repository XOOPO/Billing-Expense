import os
from flask import Flask, render_template, request, redirect, jsonify, send_file
import pandas as pd
import requests
from werkzeug.utils import secure_filename
from datetime import datetime

# ============================
# CONFIG
# ============================
API_URL = "https://script.google.com/macros/s/AKfycbwZAlVS7-UgNi8q3pwm9JaF84cuCH_ZQcfOA6_gw4tarjSbMmTuNr6zx6SpDJ8oGQIq/exec"
SHEET_CSV = "https://docs.google.com/spreadsheets/d/1eFRZufnqtjcINgrohblzwWutn4SC8H42bSPQ02YlCac/export?format=csv&gid=1703506144"

UPLOAD_FOLDER = os.path.join("static", "receipts")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB


# ============================
# UTILITIES
# ============================
def send_to_sheet(payload):
    try:
        requests.post(API_URL, json=payload, timeout=8)
    except Exception as e:
        print("Send to sheet failed:", e)


def read_sheet():
    try:
        df = pd.read_csv(SHEET_CSV)
    except Exception as e:
        print("Sheet read error:", e)
        return pd.DataFrame(columns=[
            "date","tool","used_by","department","amount",
            "company","status","cycle","renewal","receipt","desc"
        ])

    # NORMALIZE COLUMN NAMES AUTOMATICALLY
    col_map = {}
    for c in df.columns:
        lc = c.lower()
        if "date" in lc and ("purchase" in lc or "date" == lc): col_map[c] = "date"
        elif "tool" in lc or "service" in lc: col_map[c] = "tool"
        elif "use" in lc: col_map[c] = "used_by"
        elif "department" in lc: col_map[c] = "department"
        elif "amount" in lc: col_map[c] = "amount"
        elif "company" in lc: col_map[c] = "company"
        elif "status" in lc: col_map[c] = "status"
        elif "cycle" in lc: col_map[c] = "cycle"
        elif "renew" in lc: col_map[c] = "renewal"
        elif "receipt" in lc: col_map[c] = "receipt"
        elif "desc" in lc: col_map[c] = "desc"
        else: col_map[c] = c

    df = df.rename(columns=col_map)

    # ENSURE REQUIRED COLUMNS
    for col in ["date","tool","used_by","department","amount","company","status","cycle","renewal","receipt","desc"]:
        if col not in df.columns:
            df[col] = ""

    # FORMAT DATA
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["amount"] = (
        df["amount"]
        .astype(str)
        .str.replace(r"[^\d\.]", "", regex=True)
        .astype(float)
        .fillna(0)
    )

    return df


def last_months(df, n=6):
    df = df.dropna(subset=["date"])
    df["month"] = df["date"].dt.to_period("M")
    trend = df.groupby("month")["amount"].sum().reset_index()
    trend = trend.sort_values("month").tail(n)
    trend["label"] = trend["month"].dt.strftime("%b %Y")
    return trend


# ============================
# ROUTES
# ============================
@app.route("/")
def dashboard():
    return render_template("dashboard.html")


@app.route("/add", methods=["GET", "POST"])
def add():
    if request.method == "POST":
        form = request.form
        file = request.files.get("receipt")

        receipt_path = ""
        if file and file.filename:
            filename = secure_filename(datetime.utcnow().strftime("%Y%m%d%H%M%S") + "_" + file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)
            receipt_path = "/" + filepath.replace("\\", "/")

        payload = {
            "date": form.get("date", ""),
            "tool": form.get("tool", ""),
            "used_by": form.get("used_by", ""),
            "department": form.get("department", ""),
            "amount": form.get("amount", "0"),
            "currency": form.get("currency", "USD"),  # ✅ TAMBAH
            "company": form.get("company", ""),
            "status": form.get("status", ""),
            "cycle": form.get("cycle", ""),
            "renewal": form.get("renewal", ""),
            "receipt": receipt_path or form.get("receipt", ""),
            "desc": form.get("desc", "")
        }

        send_to_sheet(payload)
        return redirect("/")

    return render_template("add.html")


@app.route("/api/data")
def api_data():
    month = request.args.get("month", "All")

    df = read_sheet()

    # FILTER MONTH
    if month != "All":
        try:
            df = df[df["date"].dt.month == int(month)]
        except:
            pass

    # KPI
    total = float(df["amount"].sum())

    vendor = (
        df.groupby("tool")["amount"]
        .sum()
        .reset_index()
        .sort_values("amount", ascending=False)
        .fillna("")
    )

    department = (
        df.groupby("department")["amount"]
        .sum()
        .reset_index()
        .sort_values("amount", ascending=False)
        .fillna("")
    )

    # TREND
    trend = last_months(read_sheet(), n=6)

    return jsonify({
        "total": total,
        "vendor": vendor.to_dict("records"),
        "department": department.to_dict("records"),
        "trend_months": trend["label"].tolist(),
        "trend_values": trend["amount"].astype(float).tolist()
    })


@app.route("/download/csv")
def download_csv():
    try:
        path = os.path.join("static", "export.csv")
        read_sheet().to_csv(path, index=False)
        return send_file(path, as_attachment=True, download_name="expenses.csv")
    except Exception as e:
        return str(e), 500


# ============================
# RUN
# ============================
if __name__ == "__main__":
    app.run(port=5000, debug=True)
