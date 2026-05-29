from flask import Flask, flash, render_template, request, redirect, url_for, send_file, send_from_directory, abort
import pandas as pd
import time
import os
from io import BytesIO
import secrets
from datetime import datetime

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

REMOTE_LOG = os.path.join(BASE_DIR, "remote_log.xlsx")
FACULTY_FILE = os.path.join(BASE_DIR, "faculty_data.xlsx")
ROOM_FILE = os.path.join(BASE_DIR, "room_data.xlsx")
DAY_FILE = os.path.join(BASE_DIR, "day_order.xlsx")

ISSUE_FOLDER = os.path.join(BASE_DIR, "issued_files")
os.makedirs(ISSUE_FOLDER, exist_ok=True)


# ================= SAFE EXCEL =================
def safe_read(path, cols):
    if not os.path.exists(path):
        df = pd.DataFrame(columns=cols)
        df.to_excel(path, index=False)
        return df
    df = pd.read_excel(path)
    df.columns = df.columns.str.strip()
    return df


def load_remote_log():
    cols = ["Faculty Id", "Faculty Name", "School", "Mobile Number",
            "Room Number", "Remote ID", "Date of Issue", "Time of Issue",
            "Date of Return", "Time of Return", "Return Status"]
    return safe_read(REMOTE_LOG, cols)


def load_faculty_data():
    cols = ["Faculty ID", "Faculty Name", "Mobile Number", "school"]
    return safe_read(FACULTY_FILE, cols)


def load_room_data():
    cols = ["Faculty ID", "Faculty Name", "Room Number", "Slot Name"]
    return safe_read(ROOM_FILE, cols)


def load_day_order_data():
    cols = ["Day", "Slot Name", "Start Time", "End Time"]
    return safe_read(DAY_FILE, cols)


# ================= HOME =================
@app.route('/')
def index():
    df = load_remote_log()
    return render_template('index.html', remotes=df.iloc[::-1].to_dict('records'))


# ================= ISSUE REMOTE =================
@app.route('/remote_collection', methods=['GET', 'POST'])
def remote_collection():

    df = load_remote_log()
    faculty_df = load_faculty_data()

    error_message = None
    success_message = None
    slot_message = None

    if request.method == 'POST':

        faculty_id = request.form.get('faculty_id', '').strip()
        remote_id = request.form.get('remote_id', '').strip()

        issue_time = time.strftime("%H:%M:%S")
        issue_date = time.strftime("%Y-%m-%d")

        # Faculty check
        faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str) == faculty_id]

        if faculty_info.empty:
            error_message = "Faculty ID not found!"

        else:
            faculty_name = faculty_info.iloc[0]["Faculty Name"]
            mobile_number = faculty_info.iloc[0]["Mobile Number"]
            school_name = faculty_info.iloc[0]["school"]

            # duplicate check
            existing = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")]

            if not existing.empty:
                error_message = "Remote already issued!"

            else:

                success_message = f"Remote {remote_id} issued to {faculty_name}"

                new_entry = {
                    "Faculty Id": faculty_id,
                    "Faculty Name": faculty_name,
                    "School": school_name,
                    "Mobile Number": mobile_number,
                    "Room Number": "N/A",
                    "Remote ID": remote_id,
                    "Date of Issue": issue_date,
                    "Time of Issue": issue_time,
                    "Date of Return": "",
                    "Time of Return": "",
                    "Return Status": "Not Returned"
                }

                # FIXED append
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                df.to_excel(REMOTE_LOG, index=False)

    return render_template(
        'remote_collection.html',
        remotes=load_remote_log().tail(10).iloc[::-1].to_dict('records'),
        error_message=error_message,
        success_message=success_message,
        slot_message=slot_message
    )


# ================= RETURN =================
@app.route('/remote_return', methods=['GET', 'POST'])
def remote_return():

    df = load_remote_log()

    if request.method == 'POST':
        remote_id = request.form.get('remote_id', '').strip()

        idx = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")].index

        if not idx.empty:
            df.loc[idx, "Return Status"] = "Returned"
            df.loc[idx, "Date of Return"] = time.strftime("%Y-%m-%d")
            df.loc[idx, "Time of Return"] = time.strftime("%H:%M:%S")

            df.to_excel(REMOTE_LOG, index=False)
            flash("Remote returned successfully!", "success")

        return redirect(url_for('remote_return'))

    return render_template(
        'remote_return.html',
        remotes=df[df["Return Status"] == "Not Returned"].tail(10).iloc[::-1].to_dict('records')
    )


# ================= RUN =================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)