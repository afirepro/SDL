from flask import Flask, request
import pandas as pd
import pywhatkit
import time
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Get the uploaded Excel file
            attendance_file = request.files['attendance_file']

            if attendance_file:
                # Save the uploaded file to 'uploads' directory
                file_path = os.path.join('./uploads', attendance_file.filename)
                attendance_file.save(file_path)

                # Read the Excel file starting from the 9th row (0-indexed, so skip first 8)
                df = pd.read_excel(file_path, skiprows=8)

                # Rename the relevant columns for easy access
                df.columns = ['S.No.', 'Enrollment_No', 'Name', 'Phone'] + df.columns.tolist()[4:]

                # Convert the 'Phone' column to string and clean up the values
                df['Phone'] = df['Phone'].astype(str).str.split('.').str[0]

                # Initialize a list to store students with low attendance
                low_attendance_students = []

                # Filter for columns that are numeric and likely represent percentages
                percentage_columns = [col for col in df.columns if df[col].dtype in ['float64', 'int64']]

                # Check if any percentage value is below 75
                for index, row in df.iterrows():
                    if any(row[col] < 75 for col in percentage_columns if not pd.isna(row[col])):
                        low_attendance_students.append({'Name': row['Name'], 'Phone': row['Phone']})
                print("!!!!",low_attendance_students,"!!!!!")


                # Save the low attendance students to an Excel file
                low_attendance_df = pd.DataFrame(low_attendance_students).drop_duplicates()
                low_attendance_df.to_excel('low_attendance_students.xlsx', index=False)

                # Get the current time
                current_time = time.localtime()
                send_hour = current_time.tm_hour
                send_minute = current_time.tm_min + 1  # Buffer time

                # Send WhatsApp messages
                for student in low_attendance_students:
                    number = f"+91{student['Phone']}"
                    name = student['Name']
                    message = f"Dear Parent, your ward {name} has an attendance below 75%. Please ensure they attend classes regularly."

                    # If minute exceeds 59, adjust the hour
                    if send_minute > 59:
                        send_minute = 0
                        send_hour += 1

                    try:
                        pywhatkit.sendwhatmsg(number, message, send_hour, send_minute)
                        print(f"Message scheduled to {number} at {send_hour}:{send_minute:02d}")
                    except Exception as e:
                        print(f"Failed to send message to {number}: {str(e)}")

                    send_minute += 1  # Increment minute for the next message

                return f"<h2>Messages scheduled successfully! File processed: {attendance_file.filename}</h2>"

        except Exception as e:
            return f"<h2>An error occurred: {str(e)}</h2>"

    # HTML for uploading the Excel file
    return '''
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Attendance Notification System</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f4f4f4;
                    text-align: center;
                    margin: 0;
                    padding: 0;
                }
                .container {
                    margin-top: 100px;
                }
                h1 {
                    color: #333;
                }
                form {
                    display: inline-block;
                    margin-top: 20px;
                    padding: 20px;
                    background-color: #fff;
                    border-radius: 5px;
                    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                }
                label {
                    display: block;
                    margin-bottom: 10px;
                    font-size: 1.2em;
                    color: #555;
                }
                input[type="file"] {
                    margin-bottom: 20px;
                }
                button {
                    padding: 10px 20px;
                    font-size: 1.1em;
                    background-color: #007bff;
                    color: #fff;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                }
                button:hover {
                    background-color: #0056b3;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Upload Attendance File</h1>
                <form action="/" method="post" enctype="multipart/form-data">
                    <label for="attendance_file">Choose an Excel file:</label>
                    <input type="file" name="attendance_file" id="attendance_file" accept=".xlsx" required>
                    <button type="submit">Process and Send Messages</button>
                </form>
            </div>
        </body>
        </html>
    '''

if  __name__ == '__main__':

    if not os.path.exists('./uploads'):
        os.makedirs('./uploads')
    app.run(debug=True)