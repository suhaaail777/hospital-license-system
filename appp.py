from flask import Flask, render_template, redirect, url_for, flash, request, session
import pandas as pd
from datetime import datetime
import yagmail
from apscheduler.schedulers.background import BackgroundScheduler
import os
import webbrowser
import threading
import socket

app = Flask(__name__)
app.secret_key = "your_secret_key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "license renewal details.xlsx")

ALERT_DAYS = 30
EMAIL_SENDER = "suhaaail777@gmail.com"
EMAIL_PASSWORD = "beqf qrdn xgyu qiet"

USERS = {
    "admin": {"password": "admin123", "role": "admin", "department": None},
    "hr_user": {"password": "hr123", "role": "user", "department": "hr"},
    "it_user": {"password": "it123", "role": "user", "department": "it"}, 
    "bio_user": {"password": "bio123", "role": "user", "department": "bio"},
}

def normalize(txt):
    return str(txt).strip().lower()

def internet_available():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except:
        return False

def get_data():
    """Read data from Excel only (offline-friendly)"""
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
    except FileNotFoundError:
        df = pd.DataFrame()

    if df.empty:
        return df

    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    df = df.rename(columns={
        "email": "alert_email_id",
        "person_incharge": "responsible_person"
    })
    if "department" in df.columns:
        df["department"] = df["department"].apply(normalize)
    return df

def update_excel(df):
    """Save changes back to Excel"""
    df.to_excel(EXCEL_FILE_PATH, index=False)

def send_mail(to_email, subject, message):
    if not to_email:
        return
    if not internet_available():
        print(f"No internet: cannot send email to {to_email}")
        return
    try:
        yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
        yag.send(to=to_email, subject=subject, contents=message)
    except Exception as e:
        print("Email error:", e)

def check_expiry(user=None, filter_days=None):
    df = get_data()
    today = datetime.today()
    alerts = []

    for _, row in df.iterrows():
        dept_sheet = normalize(row.get("department", ""))

        if user and user["role"] == "user":
            if normalize(user["department"]) != dept_sheet:
                continue

        try:
            expiry = pd.to_datetime(row["expiry_date"], dayfirst=True)
        except:
            continue

        days_left = (expiry - today).days

        if filter_days == "expired" and days_left >= 0:
            continue
        elif isinstance(filter_days, int):
            if not (0 <= days_left <= filter_days):
                continue

        alerts.append({
            "license_name": row.get("license_type", ""),
            "department": dept_sheet.title(),
            "person": row.get("responsible_person", ""),
            "email": row.get("alert_email_id", ""),
            "expiry_date": expiry.strftime("%d-%m-%Y"),
            "days_left": days_left
        })

    return alerts

def auto_send_alerts():
    alerts = check_expiry()
    for a in alerts:
        if a["days_left"] <= ALERT_DAYS:
            send_mail(
                a["email"],
                f"License Expiry Alert: {a['license_name']}",
                f"Expires on {a['expiry_date']} ({a['days_left']} days left)"
            )

@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = USERS.get(request.form["username"])
        if user and user["password"] == request.form["password"]:
            session["user"] = {
                "username": request.form["username"],
                "role": user["role"],
                "department": user["department"]
            }
            return redirect(url_for("dashboard"))
        flash("Invalid credentials")
    return render_template("login.html")

@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    user = session["user"]
    filter_value = request.args.get("filter")

    if user["role"] == "user":
        alerts = check_expiry(user, filter_value if filter_value == "expired" else None)
        departments = [user["department"].title()]

    else:
        if filter_value == "expired":
            alerts = check_expiry(None, "expired")
        elif filter_value:
            alerts = check_expiry(None, int(filter_value))
        else:
            alerts = check_expiry()

        departments = sorted(set(a["department"] for a in alerts))

    return render_template(
        "dashboard.html",
        alerts=alerts,
        departments=departments,
        user=user,
        filter_value=filter_value
    )


@app.route("/send_alerts")
def send_alerts():
    if "user" not in session:
        return redirect(url_for("login"))

    alerts = check_expiry(session["user"])
    sent = 0

    for a in alerts:
        if a["days_left"] <= ALERT_DAYS:
            send_mail(
                a["email"],
                f"License Expiry Alert: {a['license_name']}",
                f"Expiry Date: {a['expiry_date']}"
            )
            sent += 1

    flash(f"{sent} alert emails sent")
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

scheduler = BackgroundScheduler()
scheduler.add_job(auto_send_alerts, "cron", hour=8)
scheduler.start()

def open_browser():
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
