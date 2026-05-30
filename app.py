from flask import Flask, flash, render_template, request, redirect, url_for, send_file, send_from_directory, abort  # type: ignore
import pandas as pd # type: ignore
import time
import os
from io import BytesIO
import secrets
from datetime import datetime

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)  # Unique runtime secret key

# Load the remote log (or create an empty one)
def load_remote_log():
    try:
        df = pd.read_excel("remote_log.xlsx")
        expected_columns = ["Faculty Id", "Faculty Name", "School", "Mobile Number", 
                            "Room Number", "Remote ID", "Date of Issue", "Time of Issue", 
                            "Date of Return", "Time of Return", "Return Status"]
        df.columns = [col.strip() for col in df.columns]  # Strip leading/trailing spaces
        return df[expected_columns]
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty Id", "Faculty Name", "School", "Mobile Number", 
                                   "Room Number", "Remote ID", "Date of Issue", "Time of Issue", 
                                   "Date of Return", "Time of Return", "Return Status"])
        df.to_excel("remote_log.xlsx", index=False)
        return df

# Load faculty data (or create an empty one)
def load_faculty_data():
    try:
        df = pd.read_excel("faculty_data.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty ID", "Faculty Name", "Mobile Number", "school"])
        df.to_excel("faculty_data.xlsx", index=False)
    df.columns = df.columns.str.strip()
    return df

# Load Slot Data (contains slot times and names)
def load_slot_data():
    try:
        df = pd.read_excel("slot_data.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Slot Name", "Slot Time"])
        df.to_excel("slot_data.xlsx", index=False)
    return df

# Load Room Data (contains room numbers, faculty IDs, slot names)
def load_room_data():
    try:
        df = pd.read_excel("room_data.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty ID", "Faculty Name", "Room Number", "Slot Name"])
        df.to_excel("room_data.xlsx", index=False)
    return df

# Load Day Order Data (contains days of the week, slot names, and hours)
def load_day_order_data():
    try:
        df = pd.read_excel("day_order.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Day", "Slot Name", "Start Time", "End Time"])
        df.to_excel("day_order.xlsx", index=False)
    return df

def load_excel_data():
    slot_data = pd.read_excel('slot_data.xlsx')
    day_order = pd.read_excel('day_order.xlsx')
    room_data = pd.read_excel('room_data.xlsx')
    return slot_data, day_order, room_data

# Function to read the timetable from an Excel file and return the DataFrame
def read_schedule_from_excel(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path)
        print("Excel file loaded successfully.")
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

# Function to determine the current slot based on the timetable and current time
def determine_current_slot(df):
    if df is None:
        return "Failed to load timetable from Excel.", 400

    current_time = datetime.now()
    current_day = current_time.strftime('%A')  # e.g., "Saturday"
    
    # ─── WEEKEND TESTING OVERRIDE ───
    if current_day in ["Saturday", "Sunday"]:
        print(f"Weekend detected ({current_day}). Overriding to Friday for slot calculations.")
        current_day = "Friday"  
    
    current_hour = current_time.hour + current_time.minute / 60
    print(f"Checking schedule layout for Day: {current_day}, Current Time Decimal: {current_hour:.2f}")

    # Clean space discrepancies from columns
    df.columns = df.columns.str.strip()

    if current_day not in df.columns:
        return f"Day '{current_day}' not found in the schedule columns.", 400

    slots = df['Time Slot'].dropna().reset_index(drop=True)
    slot_names_for_day = df[current_day].dropna().reset_index(drop=True)

    for index, slot in enumerate(slots):
        try:
            slot_str = str(slot).strip()
            if '-' not in slot_str:
                continue
                
            start_time, end_time = slot_str.split('-')
            
            # Helper logic to accurately parse trailing zeros in float values (e.g., 9.3 vs 9.30)
            def parse_time_to_decimal(time_val_str):
                parts = time_val_str.strip().split('.')
                hours = float(parts[0])
                minutes_string = parts[1] if len(parts) > 1 else "0"
                if len(minutes_string) == 1 and minutes_string != "0":
                    minutes_string += "0"  # Corrects "3" back into "30"
                minutes = float(minutes_string)
                return hours + (minutes / 60.0)

            start_time_decimal = parse_time_to_decimal(start_time)
            end_time_decimal = parse_time_to_decimal(end_time)

            if end_time_decimal < start_time_decimal:
                end_time_decimal += 24  # Wrap around midnight boundary condition

            print(f"Checking Slot Match [{index + 1}]: {start_time_decimal:.2f} <= {current_hour:.2f} < {end_time_decimal:.2f}")

            if start_time_decimal <= current_hour < end_time_decimal:
                slot_name = slot_names_for_day[index]
                return f"Current Slot: {slot_name} ({slot})"
        except Exception as e:
            print(f"Error processing matching matrix for token '{slot}': {e}")
            continue
    
    return "No current slot found", 404

# Function to get the room assignment for a given faculty and slot
def get_room_for_faculty(faculty_id, slot_name, room_data):
    filtered_data = room_data[room_data['Faculty ID'] == faculty_id]
    for slot in filtered_data['Slot Name']:
        slots = slot.split('+')
        if slot_name in slots:
            return filtered_data[filtered_data['Slot Name'] == slot].iloc[0]['Room Number']
    return None

# Home route - shows remote list
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
    faculty_details = None  
    success_message = ""  
    error_message = None
    slot_message = None  

    if request.method == 'POST':
        faculty_id = request.form['faculty_id']
        remote_id = request.form['remote_id']
        issue_time = time.strftime("%H:%M:%S")
        issue_date = time.strftime("%Y-%m-%d")
        slot_data, day_order, room_data = load_excel_data()

        room_data.columns = room_data.columns.str.strip() 

        # Evaluate current time slot bounds
        slot_message = determine_current_slot(day_order)
        print("System Evaluated Slot Status:", slot_message)  

        if isinstance(slot_message, tuple) or slot_message.startswith("No current slot found") or "Current Slot" not in slot_message:
            error_message = "No slot found for the current time."
            return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=str(slot_message))
        
        # Isolate the exact slot token string (e.g., "B1+TB1")
        slot_name = slot_message.split(":")[1].split("(")[0].strip()  
        print(f"Querying allocation indexes for Room Slot: {slot_name}")

        faculty_id = str(faculty_id)  
        room_data['Faculty ID'] = room_data['Faculty ID'].astype(str)  

        # Query room assignment matching Faculty ID and active Slot Name string boundary
        room_assignment = room_data[(room_data['Faculty ID'] == faculty_id) & (room_data['Slot Name'].str.contains(slot_name))]

        if room_assignment.empty:
            error_message = "Room not found for this faculty in the current slot."
            return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=slot_message)
        
        room_number = room_assignment.iloc[0]['Room Number']
        
        # Verify if token is checked out
        existing_remote = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")]

        if not existing_remote.empty:
            error_message = f"Remote ID {remote_id} is not returned yet. Please take another remote."
            return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=slot_message)
        
        faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str) == faculty_id]

        if faculty_info.empty:
            error_message = "Faculty ID not found!"
            return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=slot_message)
        
        faculty_name = faculty_info.iloc[0]["Faculty Name"]
        mobile_number = faculty_info.iloc[0]["Mobile Number"]
        school_name = faculty_info.iloc[0]["school"]
        
        if school_name in ["ETHENUS", "SIXPHRASE", "FACE"]:
            success_message = f"Remote ID {remote_id} issued to {faculty_name}!"
        else:
            success_message = f"Remote ID {remote_id} issued to Prof. {faculty_name}!"

        # Create active issue transaction entry dictionary
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
        
        # Concatenate record safely into tracking frame without deprecated append method
        new_row_df = pd.DataFrame([new_entry])
        df = pd.concat([df, new_row_df], ignore_index=True)
        df.to_excel("remote_log.xlsx", index=False)

        faculty_details = {
            "faculty_id": faculty_id,
            "faculty_name": faculty_name,
            "mobile_number": mobile_number,
            "school_name": school_name,
            "remote_id": remote_id,
            "return_status": "Not Returned",
        }

        remotes = df.tail(10).iloc[::-1].to_dict(orient='records')
        return render_template('remote_collection.html', remotes=remotes, error_message=error_message, faculty_details=faculty_details, success_message=success_message, slot_message=slot_message)

    return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=slot_message)


# Route to display faculty details + remote log
@app.route('/faculty_details', methods=['POST'])
def faculty_details():
    faculty_id = request.form['faculty_id']

    faculty_df = load_faculty_data()
    remote_df = load_remote_log()

    faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str) == faculty_id]
    
    if faculty_info.empty:
        return "Faculty ID not found!", 404

    faculty_name = faculty_info.iloc[0]["Faculty Name"]
    mobile_number = faculty_info.iloc[0]["Mobile Number"]
    school_name = faculty_info.iloc[0]["school"]

    remote_info = remote_df[remote_df["Faculty ID"].astype(str) == faculty_id]
    if not remote_info.empty:
        remote_id = remote_info.iloc[0]["Remote ID"]
        return_status = remote_info.iloc[0]["Return Status"]
    else:
        remote_id, return_status = "N/A", "N/A"

    return render_template(
        'faculty_details.html', 
        faculty_id=faculty_id, 
        faculty_name=faculty_name, 
        mobile_number=mobile_number, 
        school_name=school_name, 
        remote_id=remote_id, 
        return_status=return_status
    )

# Route for Remote Return
@app.route('/remote_return', methods=['GET', 'POST'])
def remote_return():
    df = load_remote_log()
    remotes = df[df["Return Status"] == "Not Returned"].tail(10).iloc[::-1].to_dict(orient='records')

    if request.method == 'POST':
        remote_id = request.form['remote_id']
        
        index = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")].index
        if not index.empty:
            return_time = time.strftime("%H:%M:%S")
            return_date = time.strftime("%Y-%m-%d")
            df.loc[index, "Return Status"] = "Returned"
            df.loc[index, "Date of Return"] = return_date
            df.loc[index, "Time of Return"] = return_time
            df.to_excel("remote_log.xlsx", index=False)

            faculty_name = df.loc[index, "Faculty Name"].iloc[0]
            flash(f"Remote {remote_id} returned by {faculty_name}!", "success")

        return redirect(url_for('remote_return'))  

    return render_template('remote_return.html', remotes=remotes)


# Route to download filtered excel data
@app.route('/download_excel', methods=['GET', 'POST'])
def download_excel():
    if request.method == 'POST':
        start_datetime_str = request.form['start_datetime']
        end_datetime_str = request.form['end_datetime']

        start_datetime = pd.to_datetime(start_datetime_str, format="%d-%m-%Y %H:%M")
        end_datetime = pd.to_datetime(end_datetime_str, format="%d-%m-%Y %H:%M")

        df = load_remote_log()
        df['datetime'] = pd.to_datetime(df['Date of Issue'] + ' ' + df['Time of Issue'], format="%Y-%m-%d %H:%M:%S")

        df_filtered = df[(df['datetime'] >= start_datetime) & (df['datetime'] <= end_datetime)]

        if df_filtered.empty:
            return "No data found for the specified range."

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtered.to_excel(writer, index=False)
        
        output.seek(0)
        return send_file(output, download_name='remote_log_filtered.xlsx', as_attachment=True)

    return render_template('download_excel.html')


FACULTY_DATA_PATH = os.path.join('faculty_data.xlsx')
ISSUE_FOLDER = 'issued_files'
os.makedirs(ISSUE_FOLDER, exist_ok=True)

@app.route('/stationary', methods=['GET', 'POST'])
def stationary():
    success = False

    if request.method == 'POST':
        emp_id = request.form.get('emp_id', '').strip()
        category = request.form.get('category', '')
        now = datetime.now()
        date = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")

        faculty_name = 'Unknown'
        school = 'Unknown'

        try:
            faculty_df = pd.read_excel(FACULTY_DATA_PATH)
            faculty_df.columns = faculty_df.columns.str.strip()
            
            faculty_df['Faculty ID'] = faculty_df['Faculty ID'].astype(str)
            emp_id = str(emp_id)  
          
            match = faculty_df.loc[faculty_df['Faculty ID'] == emp_id]
            if not match.empty:
                faculty_name = match.iloc[0]['Faculty Name']
                school = match.iloc[0]['school']
        except Exception as e:
            print("Error reading faculty data structure matrix:", e)

        data = {
            'Faculty ID': emp_id,
            'Faculty Name': faculty_name,
            'School': school,
            'Date': date,
            'Time': time_str
        }

        if category in ['Xerox', 'Notes Paper']:
            data['Number of Copies'] = request.form.get('copies', '')
            data['Purpose'] = request.form.get('purpose', '')
        elif category in ['Marker', 'Duster']:
            data['Quantity'] = request.form.get('quantity', '')
        elif category == 'Lab Fat Paper':
            data['Number of Sheets'] = request.form.get('sheets', '')

        filename = f"{category.replace(' ', '_')}.xlsx"
        file_path = os.path.join(ISSUE_FOLDER, filename)

        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        else:
            df = pd.DataFrame([data])

        df.to_excel(file_path, index=False)
        success = True

    return render_template('stationary.html', success=success)


# New download route for Excel files
@app.route('/download/<filename>')
def download_file(filename):
    allowed_files = [
        "Duster.xlsx",
        "Lab_Fat_Paper.xlsx",
        "Marker.xlsx",
        "Notes_Paper.xlsx",
        "Xerox.xlsx"
    ]
    if filename not in allowed_files:
        abort(404)
    try:
        return send_from_directory(ISSUE_FOLDER, filename, as_attachment=True)
    except Exception as e:
        abort(404)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
