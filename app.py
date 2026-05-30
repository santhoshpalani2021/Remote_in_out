from flask import Flask, flash, render_template, request, redirect, url_for, send_file, send_from_directory, abort  # type: ignore
import pandas as pd # type: ignore
import time
import os
from io import BytesIO
import secrets
from datetime import datetime

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Load the remote log (or create an empty one)
def load_remote_log():
    try:
        df = pd.read_excel("remote_log.xlsx")
        expected_columns = ["Faculty Id", "Faculty Name", "School", "Mobile Number", 
                            "Room Number", "Remote ID", "Date of Issue", "Time of Issue", 
                            "Date of Return", "Time of Return", "Return Status"]
        df.columns = [col.strip() for col in df.columns]
        return df[expected_columns]
    except Exception:
        df = pd.DataFrame(columns=["Faculty Id", "Faculty Name", "School", "Mobile Number", 
                                   "Room Number", "Remote ID", "Date of Issue", "Time of Issue", 
                                   "Date of Return", "Time of Return", "Return Status"])
        df.to_excel("remote_log.xlsx", index=False)
        return df

def load_faculty_data():
    try:
        df = pd.read_excel("faculty_data.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty ID", "Faculty Name", "Mobile Number", "school"])
        df.to_excel("faculty_data.xlsx", index=False)
    df.columns = df.columns.str.strip()
    return df

def load_excel_data():
    try: slot_data = pd.read_excel('slot_data.xlsx')
    except Exception: slot_data = pd.DataFrame(columns=["Slot Name", "Slot Time"])
    try: day_order = pd.read_excel('day_order.xlsx')
    except Exception: day_order = pd.DataFrame(columns=["Time Slot", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
    try: room_data = pd.read_excel('room_data.xlsx')
    except Exception: room_data = pd.DataFrame(columns=["Faculty ID", "Faculty Name", "Room Number", "Slot Name"])
        
    return slot_data, day_order, room_data

# FIX: ROBUST TIMETABLE PARSER USING TRUE CLOCK STRINGS
def determine_current_slot(df):
    if df is None or df.empty:
        return "No Schedule Available"

    now = datetime.now()
    current_day = now.strftime('%A')  # Captures "Saturday", "Sunday", etc.
    current_time_str = now.strftime("%H:%M")
    
    df.columns = df.columns.str.strip()

    if current_day not in df.columns:
        return f"Day column '{current_day}' missing"

    slots = df['Time Slot'].dropna().reset_index(drop=True)
    slot_names_for_day = df[current_day].dropna().reset_index(drop=True)

    # Convert a string like "7.30" or "14.30" directly to a clean datetime time object
    def clean_time_obj(t_str):
        t_str = t_str.strip().replace(':', '.')
        hours, minutes = map(int, t_str.split('.'))
        return now.replace(hour=hours, minute=minutes, second=0, microsecond=0)

    for index, slot in enumerate(slots):
        try:
            slot_str = str(slot).strip()
            if '-' not in slot_str:
                continue
                
            start_str, end_str = slot_str.split('-')
            start_time = clean_time_obj(start_str)
            end_time = clean_time_obj(end_str)

            # If testing a slot gap like 8.30 vs 8.31, pad the boundary by 1 minute
            start_time = start_time.replace(minute=max(0, start_time.minute - 1))

            if start_time <= now <= end_time:
                if index < len(slot_names_for_day):
                    return f"Current Slot: {slot_names_for_day[index]} ({slot})"
        except Exception as e:
            print(f"Error parsing slot matrix row: {e}")
            continue
    
    return "No current slot found"

@app.route('/')
def index():
    df = load_remote_log()
    remotes = df.iloc[::-1].to_dict(orient='records')  
    return render_template('index.html', remotes=remotes)

@app.route('/remote_collection', methods=['GET', 'POST'])
def remote_collection():
    df = load_remote_log()
    faculty_df = load_faculty_data()
    remotes = df.tail(10).iloc[::-1].to_dict(orient='records')
    
    error_message = None
    success_message = ""
    faculty_details = None
    slot_message = "N/A"

    slot_name = "N/A"
    room_number = "N/A"
    faculty_name = "Unknown Faculty"
    mobile_number = "N/A"
    school_name = "N/A"

    if request.method == 'POST':
        faculty_id = str(request.form['faculty_id']).strip()
        remote_id = str(request.form['remote_id']).strip()
        issue_time = time.strftime("%H:%M:%S")
        issue_date = time.strftime("%Y-%m-%d")
        
        _, day_order, room_data = load_excel_data()
        room_data.columns = room_data.columns.str.strip() 

        # Evaluate live schedule window
        slot_message = determine_current_slot(day_order)
        
        if "Current Slot:" in slot_message:
            slot_name = slot_message.split(":")[1].split("(")[0].strip()
        
        # Match Room Allocations from user input files
        if slot_name.upper() not in ["NIL", "FREE", "EMPTY", "N/A", ""]:
            room_data['Faculty ID'] = room_data['Faculty ID'].astype(str).str.strip()
            room_assignment = room_data[(room_data['Faculty ID'] == faculty_id) & (room_data['Slot Name'].str.contains(slot_name))]
            if not room_assignment.empty:
                room_number = room_assignment.iloc[0]['Room Number']
            else:
                error_message = f"Notice: Room allocation matching slot '{slot_name}' wasn't found in system registers. Logged as N/A."
        else:
            error_message = f"Notice: No active timetable slots right now (Slot is {slot_name}). Logged as N/A."

        # Fetch Faculty details
        faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str).str.strip() == faculty_id]
        if not faculty_info.empty:
            faculty_name = str(faculty_info.iloc[0]["Faculty Name"]).strip()
            mobile_number = faculty_info.iloc[0]["Mobile Number"]
            school_name = str(faculty_info.iloc[0]["school"]).strip()

        # Check checkout log safety constraints
        existing_remote = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")]
        if not existing_remote.empty:
            error_message = f"Error: Remote ID {remote_id} is already issued and not returned!"
            return render_template('remote_collection.html', remotes=remotes, error_message=error_message, slot_message=slot_message)

        # Log entry assembly
        new_entry = {
            "Faculty Id": faculty_id,
            "Faculty Name": faculty_name,
            "School": school_name,
            "Mobile Number": mobile_number,
            "Room Number": room_number,
            "Remote ID": remote_id,
            "Date of Issue": issue_date,
            "Time of Issue": issue_time,
            "Date of Return": "",
            "Time of Return": "",
            "Return Status": "Not Returned"
        }
        
        new_row_df = pd.DataFrame([new_entry])
        df = pd.concat([df, new_row_df], ignore_index=True)
        df.to_excel("remote_log.xlsx", index=False)

        # FIX: FIXES DOUBLE PREFIX ERROR
        # Clean potential duplicate prefixes already saved in source Excel strings
        clean_name = faculty_name.replace("Prof.", "").replace("Prof", "").strip()
        if school_name.upper() in ["ETHENUS", "SIXPHRASE", "FACE"]:
            success_message = f"Remote ID {remote_id} issued to {clean_name}!"
        else:
            success_message = f"Remote ID {remote_id} issued to Prof. {clean_name}!"

        faculty_details = {
            "faculty_id": faculty_id,
            "faculty_name": f"Prof. {clean_name}" if school_name.upper() not in ["ETHENUS", "SIXPHRASE", "FACE"] else clean_name,
            "mobile_number": mobile_number,
            "school_name": school_name,
            "remote_id": remote_id,
            "return_status": "Not Returned",
        }

        remotes = df.tail(10).iloc[::-1].to_dict(orient='records')
        return render_template('remote_collection.html', remotes=remotes, error_message=error_message, faculty_details=faculty_details, success_message=success_message, slot_message=slot_message)

    return render_template('remote_collection.html', remotes=remotes, error_message=error_message, slot_message=slot_message)

@app.route('/remote_return', methods=['GET', 'POST'])
def remote_return():
    df = load_remote_log()
    remotes = df[df["Return Status"] == "Not Returned"].tail(10).iloc[::-1].to_dict(orient='records')

    if request.method == 'POST':
        remote_id = request.form['remote_id']
        index = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")].index
        if not index.empty:
            df.loc[index, "Return Status"] = "Returned"
            df.loc[index, "Date of Return"] = time.strftime("%Y-%m-%d")
            df.loc[index, "Time of Return"] = time.strftime("%H:%M:%S")
            df.to_excel("remote_log.xlsx", index=False)
            flash(f"Remote {remote_id} successfully returned!", "success")
        return redirect(url_for('remote_return'))  

    return render_template('remote_return.html', remotes=remotes)

@app.route('/stationary', methods=['GET', 'POST'])
def stationary():
    success = False
    if request.method == 'POST':
        emp_id = request.form.get('emp_id', '').strip()
        category = request.form.get('category', '')
        
        data = {
            'Faculty ID': emp_id,
            'Faculty Name': 'Unknown',
            'School': 'Unknown',
            'Date': datetime.now().strftime("%Y-%m-%d"),
            'Time': datetime.now().strftime("%H:%M:%S")
        }

        os.makedirs('issued_files', exist_ok=True)
        file_path = os.path.join('issued_files', f"{category.replace(' ', '_')}.xlsx")

        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        else:
            df = pd.DataFrame([data])

        df.to_excel(file_path, index=False)
        success = True

    return render_template('stationary.html', success=success)

if __name__ == '__main__':
    print("\n" + "="*60)
    print(f"👉 DATA SAVED IN THE ACTIVE FOLDER:\n🎯 {os.getcwd()}")
    print("="*60 + "\n")
    app.run(host='0.0.0.0', port=5000, debug=True)
