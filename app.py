from flask import Flask, flash, render_template, request, redirect, url_for, send_file, send_from_directory, abort  # type: ignore
import pandas as pd # type: ignore
import time
import os
from io import BytesIO
import secrets
from datetime import datetime

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)  # Or use your own secret key

# Load the remote log (or create an empty one)
def load_remote_log():
    try:
        df = pd.read_excel("remote_log.xlsx")
        # Ensure column names are consistent
        expected_columns = ["Faculty Id", "Faculty Name", "School", "Mobile Number", 
                            "Room Number", "Remote ID", "Date of Issue", "Time of Issue", 
                            "Date of Return", "Time of Return", "Return Status"]
        df.columns = [col.strip() for col in df.columns]  # Strip any leading/trailing spaces
        return df[expected_columns]  # Ensure we load only the expected columns
    except FileNotFoundError:
        # If the file doesn't exist, create it with the correct columns
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
    df.columns = df.columns.str.strip()  # Strip any leading/trailing spaces from columns
    return df

# Load Slot Data (contains slot times and names)
def load_slot_data():
    try:
        df = pd.read_excel("slot_data.xlsx")  # Slot allocation data
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Slot Name", "Slot Time"])
        df.to_excel("slot_data.xlsx", index=False)
    return df

# Load Room Data (contains room numbers, faculty IDs, slot names)
def load_room_data():
    try:
        df = pd.read_excel("room_data.xlsx")  # Room allocation data
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Faculty ID", "Faculty Name", "Room Number", "Slot Name"])
        df.to_excel("room_data.xlsx", index=False)
    return df

# Load Day Order Data (contains days of the week, slot names, and hours)
def load_day_order_data():
    try:
        df = pd.read_excel("day_order.xlsx")  # Day wise order data (Monday to Friday)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Day", "Slot Name", "Start Time", "End Time"])
        df.to_excel("day_order.xlsx", index=False)
    return df

def load_excel_data():
    # Load the slot data
    slot_data = pd.read_excel('slot_data.xlsx')
    # Load the day order data
    day_order = pd.read_excel('day_order.xlsx')
    # Load the room data
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
    current_day = current_time.strftime('%A')  # Get the day of the week (e.g., Monday)
    current_hour = current_time.hour + current_time.minute / 60  # Include minutes for accurate comparison
    
    print(f"Checking for Day: {current_day}, Time: {current_hour:.2f}")  # Debugging output

    # Ensure the current_day exists as a column in the DataFrame
    if current_day not in df.columns:
        return f"Day '{current_day}' not found in the schedule data.", 400

    # Extract the time slots and corresponding names for today
    slots = df['Time Slot'].dropna().reset_index(drop=True)  # Time slots (e.g., '7.50-8.50')
    slot_names_for_day = df[current_day].dropna().reset_index(drop=True)  # Slot names for the day

    # Debugging: Print slot data for the day
    print(f"Slots for {current_day}: {slots}")
    print(f"Slot names for {current_day}: {slot_names_for_day}")

    # Find the matching slot based on the current time
    for index, slot in enumerate(slots):
        try:
            # Split start and end time and handle edge cases like crossing midnight
            start_time, end_time = slot.split('-')
            start_hour, start_minute = map(float, start_time.split('.'))
            end_hour, end_minute = map(float, end_time.split('.'))
            
            # If the end time is less than the start time, assume it spans midnight
            if end_hour < start_hour:
                end_hour += 24  # Adding 24 hours to the end time

            # Convert times into decimal hours (e.g., 12:30 becomes 12.5)
            start_time_decimal = start_hour + start_minute / 60
            end_time_decimal = end_hour + end_minute / 60

            print(f"Slot {index + 1}: {start_time}-{end_time} (Start: {start_time_decimal}, End: {end_time_decimal})")

            # Check if the current time is between the start and end time
            if start_time_decimal <= current_hour < end_time_decimal:
                slot_name = slot_names_for_day[index]
                return f"Current Slot: {slot_name} ({slot})"
        except ValueError as ve:
            print(f"Error parsing time slot '{slot}': {ve}")
            continue  # Skip invalid time slots
    
    return "No current slot found", 404

# Function to get the room assignment for a given faculty and slot
def get_room_for_faculty(faculty_id, slot_name, room_data):
    # Filter room data based on Faculty ID and Slot Name
    filtered_data = room_data[room_data['Faculty ID'] == faculty_id]
    for slot in filtered_data['Slot Name']:
        # Split slot names by "+" to handle multiple slots
        slots = slot.split('+')
        if slot_name in slots:
            return filtered_data[filtered_data['Slot Name'] == slot].iloc[0]['Room Number']
    return None

# Home route - shows remote list
@app.route('/')
def index():
    df = load_remote_log()
    remotes = df.iloc[::-1].to_dict(orient='records')  # Convert DataFrame to a list of dicts
    return render_template('index.html', remotes=remotes)


@app.route('/remote_collection', methods=['GET', 'POST'])
#@app.route('/faculty_scan', methods=['GET', 'POST'])
def remote_collection():
    df = load_remote_log()
    faculty_df = load_faculty_data()

    remotes = df.tail(10).iloc[::-1].to_dict(orient='records')
    faculty_details = None  # Initialize faculty details
    success_message = ""  # Initialize success message
    error_message = None
    slot_message = None  # Add a variable for slot information

    if request.method == 'POST':
        faculty_id = request.form['faculty_id']
        remote_id = request.form['remote_id']
        issue_time = time.strftime("%H:%M:%S")
        issue_date = time.strftime("%Y-%m-%d")
        slot_data, day_order, room_data = load_excel_data()

        # Strip spaces from column names in room_data
        room_data.columns = room_data.columns.str.strip()  # Ensure no extra spaces

        # Print room_data for debugging to ensure the correct format
        print("Room Data:")
        print(room_data)

        # Determine the current slot
        slot_message = determine_current_slot(day_order)
        print("Determined Slot:", slot_message)  # Debugging output

        if slot_message.startswith("No current slot found"):
            error_message = "No slot found for the current time."
        
        # Print faculty_id and slot_name for debugging
        print(f"Looking for Faculty ID: {faculty_id}, Slot: {slot_message}")

        # Clean the slot_name string to extract only the actual slot name
        # This will extract the part before the time range, e.g., "B1+TB1"
        if "Current Slot" in slot_message:
            slot_name = slot_message.split(":")[1].split("(")[0].strip()  # Extract the part before '('
        else:
            slot_name = slot_message  # Fallback if the format doesn't match

        print(f"Looking for Room for Slot: {slot_name}")

        # Fetch the room number based on faculty ID and slot
        faculty_id = str(faculty_id)  # Ensure it's a string
        room_data['Faculty ID'] = room_data['Faculty ID'].astype(str)  # Convert column to string

        # Search for the room assignment for the faculty_id and slot_name
        room_assignment = room_data[(room_data['Faculty ID'] == faculty_id) & (room_data['Slot Name'].str.contains(slot_name))]

        # Debugging output of the room assignment
        print("Room Assignment Data:", room_assignment)

        if room_assignment.empty:
            error_message = "Room not found for this faculty in the current slot."
        
        # Extract room number
        room_number = room_assignment.iloc[0]['Room Number'] if not room_assignment.empty else "N/A"
        
        # Check if the remote is already issued and not returned
        existing_remote = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")]

        if not existing_remote.empty:
            error_message = f"Remote ID {remote_id} is not returned yet. Please take another remote."
        else:
            faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str) == faculty_id]

            if faculty_info.empty:
                error_message = "Faculty ID not found!"
                faculty_name = "Unknown"
                mobile_number = "Unknown"
                school_name = "Unknown"
            else:
                faculty_name = faculty_info.iloc[0]["Faculty Name"]
                mobile_number = faculty_info.iloc[0]["Mobile Number"]
                school_name = faculty_info.iloc[0]["school"]
                # if not faculty_info.empty else ""
            if school_name in ["ETHENUS", "SIXPHRASE","FACE"]:
    # If faculty belongs to specific departments, format the success message without "Prof."
                success_message = f"Remote ID {remote_id} issued to {faculty_name}!"  # No "Prof." here
            else:
             # Default message for other faculties (using Prof.)
                success_message = f"Remote ID {remote_id} issued to  {faculty_name}!"  # Using "Prof." ADD THE WORD

            # Log the issue
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
            df = df.append(new_entry, ignore_index=True)
            df.to_excel("remote_log.xlsx", index=False)

            faculty_details = {
                "faculty_id": faculty_id,
                "faculty_name": faculty_name,
                "mobile_number": mobile_number,
                "school_name": school_name,
                "remote_id": remote_id,
                "return_status": "Not Returned",
            }
            #success_message = f"Remote ID {remote_id} issued to Prof. {faculty_name}!"

        # Reload the remote logs again after the post request and before redirect
        remotes = df.tail(10).iloc[::-1].to_dict(orient='records')
        return render_template('remote_collection.html', remotes=remotes, error_message=error_message, faculty_details=faculty_details, success_message=success_message, slot_message=slot_message)

    return render_template('remote_collection.html', remotes=remotes, error_message=error_message, success_message=success_message, slot_message=slot_message)




# Route to display faculty details + remote log
@app.route('/faculty_details', methods=['POST'])
def faculty_details():
    faculty_id = request.form['faculty_id']

    faculty_df = load_faculty_data()
    remote_df = load_remote_log()

    # Fetch faculty details
    faculty_info = faculty_df[faculty_df["Faculty ID"].astype(str) == faculty_id]
    
    if faculty_info.empty:
        return "Faculty ID not found!", 404

    faculty_name = faculty_info.iloc[0]["Faculty Name"]
    mobile_number = faculty_info.iloc[0]["Mobile Number"]
    school_name = faculty_info.iloc[0]["school"]

    # Fetch Remote ID and Status
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

    # Only include remotes that are "Not Returned"
    remotes = df[df["Return Status"] == "Not Returned"].tail(10).iloc[::-1].to_dict(orient='records')

    if request.method == 'POST':
        remote_id = request.form['remote_id']
        
        # Log the return
        index = df[(df["Remote ID"] == remote_id) & (df["Return Status"] == "Not Returned")].index
        if not index.empty:
            return_time = time.strftime("%H:%M:%S")
            return_date = time.strftime("%Y-%m-%d")
            df.loc[index, "Return Status"] = "Returned"
            df.loc[index, "Date of Return"] = return_date
            df.loc[index, "Time of Return"] = return_time
            df.to_excel("remote_log.xlsx", index=False)

            # Get the Faculty Name
            faculty_name = df.loc[index, "Faculty Name"].iloc[0]

            # Success message using flash
            flash(f"Remote {remote_id} returned by  {faculty_name}!", "success")

        return redirect(url_for('remote_return'))  # Redirect to show the success message

    return render_template('remote_return.html', remotes=remotes)


# Route to download filtered excel data
@app.route('/download_excel', methods=['GET', 'POST'])
def download_excel():
    if request.method == 'POST':
        start_datetime_str = request.form['start_datetime']
        end_datetime_str = request.form['end_datetime']

        # Convert the string datetime to pandas datetime format for form inputs
        start_datetime = pd.to_datetime(start_datetime_str, format="%d-%m-%Y %H:%M")
        end_datetime = pd.to_datetime(end_datetime_str, format="%d-%m-%Y %H:%M")

        df = load_remote_log()

        # Combine 'Date of Issue' and 'Time of Issue' columns into a single datetime column
        df['datetime'] = pd.to_datetime(df['Date of Issue'] + ' ' + df['Time of Issue'], format="%Y-%m-%d %H:%M:%S")

        # Filter the DataFrame based on the specified datetime range
        df_filtered = df[(df['datetime'] >= start_datetime) & (df['datetime'] <= end_datetime)]

        # Check if the filtered DataFrame is empty
        if df_filtered.empty:
            return "No data found for the specified range."

        # Create a new Excel file in memory
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
def statinary():
    success = False

    if request.method == 'POST':
        emp_id = request.form.get('emp_id', '').strip()
        category = request.form.get('category', '')
        now = datetime.now()
        date = now.strftime("%Y-%m-%d")
        time = now.strftime("%H:%M:%S")

        # Default values
        faculty_name = 'Unknown'
        school = 'Unknown'

        # Load and clean faculty data
        try:
            faculty_df = pd.read_excel(FACULTY_DATA_PATH)
            faculty_df.columns = faculty_df.columns.str.strip()  # Clean column names
            print("Available columns:", faculty_df.columns.tolist())

    # Convert Faculty ID column to string for safe comparison
            faculty_df['Faculty ID'] = faculty_df['Faculty ID'].astype(str)
            emp_id = str(emp_id)  # Ensure form ID is string too
          
            match = faculty_df.loc[faculty_df['Faculty ID'] == emp_id]
            if not match.empty:
                faculty_name = match.iloc[0]['Faculty Name']
                school = match.iloc[0]['school']
        except Exception as e:
            print("Error reading faculty data:", e)

        # Prepare data to store
        data = {
            'Faculty ID': emp_id,
            'Faculty Name': faculty_name,
            'School': school,
            'Date': date,
            'Time': time
        }

        # Add category-specific fields
        if category in ['Xerox', 'Notes Paper']:
            data['Number of Copies'] = request.form.get('copies', '')
            data['Purpose'] = request.form.get('purpose', '')
        elif category in ['Marker', 'Duster']:
            data['Quantity'] = request.form.get('quantity', '')
        elif category == 'Lab Fat Paper':
            data['Number of Sheets'] = request.form.get('sheets', '')

        # Save to Excel
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
    # Optional: Validate the filename to allow only specific files if needed.
    # For example, you may restrict to known categories:
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
        # Serve the file from the issued_files directory
        return send_from_directory(ISSUE_FOLDER, filename, as_attachment=True)
    except Exception as e:
        abort(404)



# Main entry point to run the app
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)  # Adjust the port as needed
