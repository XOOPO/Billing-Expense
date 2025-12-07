import os
from flask import Flask, render_template, request, redirect, jsonify, send_file
import pandas as pd
import requests
from werkzeug.utils import secure_filename
from datetime import datetime

# ============================
# CONFIG
# ============================
API_URL = "https://script.google.com/macros/s/AKfycbyHmv30wUZ6zyjEt2l6JwhNnPQg6Ig9QpHLFhjNHjWWtdDxn7XbaD2cE1aVJ9kc6rBM/exec"

UPLOAD_FOLDER = os.path.join("static", "receipts")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB


# ============================
# READ ALL SHEETS (NEW)
# ============================
def read_all_sheets():
    """
    Ambil semua data expense dari semua sheet:
    - XOO Expense of October
    - XOO Expense of November
    - XOO Expense of December
    - dll...
    """
    try:
        data = requests.get(API_URL + "?mode=all").json()
    except Exception as e:
        print("Failed to load multi-sheet data:", e)
        return pd.DataFrame()

    if not data:
        return pd.DataFrame()

    df = pd.DataFrame(data)

    # --- Normalisasi nama kolom dari Google Sheet ---
    col_map = {}
    for c in df.columns:
        lc = c.lower()

        if "date" in lc: col_map[c] = "date"
        elif "tool" in lc or "service" in lc: col_map[c] = "tool"
        elif "use" in lc: col_map[c] = "used_by"
        elif "depart" in lc: col_map[c] = "department"
        elif "amount" in lc: col_map[c] = "amount"
        elif "company" in lc: col_map[c] = "company"
        elif "status" in lc: col_map[c] = "status"
        elif "cycle" in lc: col_map[c] = "cycle"
        elif "renew" in lc: col_map[c] = "renewal"
        elif "receipt" in lc: col_map[c] = "receipt"
        elif "desc" in lc: col_map[c] = "desc"
        else: col_map[c] = c

    df = df.rename(columns=col_map)

    # Pastikan semua kolom wajib ada
    required = ["date","tool","used_by","department","amount","company",
                "status","cycle","renewal","receipt","desc"]
    for col in required:
        if col not in df.columns:
            df[col] = ""

    # Format tanggal + amount
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["amount"] = (
        df["amount"]
        .astype(str)
        .str.replace(r"[^\d\.]", "", regex=True)
        .astype(float)
        .fillna(0)
    )

    return df


# ============================
# TREND HELPER
# ============================
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


# =============================================
# ADD DATA → tetap kirim ke December (Apps Script)
# =============================================
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
            "currency": form.get("currency", "USD"),
            "company": form.get("company", ""),
            "status": form.get("status", ""),
            "cycle": form.get("cycle", ""),
            "renewal": form.get("renewal", ""),
            "receipt": receipt_path or form.get("receipt", ""),
            "desc": form.get("desc", "")
        }

        try:
            requests.post(API_URL, json=payload, timeout=8)
        except:
            print("Failed to send new data")

        return redirect("/")

    return render_template("add.html")


# =============================================
# API DATA → sudah multi-sheet otomatis
# =============================================
@app.route("/api/data")
def api_data():
    month = request.args.get("month", "All")
    df = read_all_sheets()

    # Filter bulan (kalau user pilih)
    if month != "All":
        try:
            df = df[df["date"].dt.month == int(month)]
        except:
            pass

    total = float(df["amount"].sum())

    vendor = (
        df.groupby("tool")["amount"]
        .sum()
        .reset_index()
        .sort_values("amount", ascending=False)
    )

    department = (
        df.groupby("department")["amount"]
        .sum()
        .reset_index()
        .sort_values("amount", ascending=False)
    )

    trend = last_months(read_all_sheets(), n=6)

    return jsonify({
        "total": total,
        "vendor": vendor.to_dict("records"),
        "department": department.to_dict("records"),
        "trend_months": trend["label"].tolist(),
        "trend_values": trend["amount"].astype(float).tolist()
    })


# =============================================
# DOWNLOAD CSV
# =============================================
@app.route("/download/csv")
def download_csv():
    try:
        path = os.path.join("static", "export.csv")
        read_all_sheets().to_csv(path, index=False)
        return send_file(path, as_attachment=True, download_name="expenses.csv")
    except Exception as e:
        return str(e), 500


# ============================
# RUN
# ============================
if __name__ == "__main__":
    app.run(port=5000, debug=True)
