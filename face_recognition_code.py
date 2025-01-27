from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import face_recognition
import cv2
import numpy as np
import os
import sqlite3
import xlwt
from datetime import datetime
import xlrd
from xlutils.copy import copy as xl_copy

# Authenticate user before starting
username = input("Enter Username: ")
password = input("Enter Password: ")

if username.lower() != "orchid" or password != "123456":
    print("Invalid credentials! Exiting program.")
    exit()

# Function to send emails to parents of absent students
def send_email(student_email, parent_email, student_name, subject_name, status):
    from_email = "mominamaan001@gmail.com"  # Replace with your email
    from_password = "nxvh olbe mwiy fvde"  # Replace with your App Password

    # Subject for email
    subject = "Attendance Status of Your Ward"

    # Body for parent email
    parent_body = f"""Subject: Attendance Status of Your Ward

Dear Parent,

Please be informed that your ward, {student_name}, has been marked {status} for the {subject_name} class held today ({datetime.today().strftime('%Y-%m-%d')}). Kindly ensure regular attendance.

Warm Regards,
NKOCET, Artificial Intelligence and Data Science Smart Attendance System
"""

    # Email to parent
    parent_msg = MIMEMultipart()
    parent_msg['From'] = from_email
    parent_msg['To'] = parent_email
    parent_msg['Subject'] = subject
    parent_msg.attach(MIMEText(parent_body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, parent_email, parent_msg.as_string())
        server.quit()
        print(f"Email sent to parent ({parent_email}) of absent student: {student_name}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Prompt for subject name
subject_name = input("Enter the subject name for attendance: ")

# Define target directory and create file path with current date, day, and subject name
target_directory = r"C:\Users\momin\OneDrive\Desktop\smart"
today = datetime.today()
file_name = f"{today.strftime('%Y-%m-%d')}{today.strftime('%A')}{subject_name}.xls"
file_path = os.path.join(target_directory, file_name)

# Check if file already exists
if os.path.exists(file_path):
    print(f"Sheet already exists for {subject_name} on {today.strftime('%Y-%m-%d')}. File path: {file_path}")
else:
    # Create workbook and initial sheet
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet(subject_name)
    # Add headers
    headers = ['Name', 'Roll No', 'Email ID', 'Branch', 'Year', 'Division', 'Date', 'Status']
    for col, header in enumerate(headers):
        sheet1.write(0, col, header)
    wb.save(file_path)
    print(f"Excel file created successfully for {subject_name} at: {file_path}")

# Fetch students' details from the database
def fetch_students_from_db():
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name, roll_no, email, branch, year, division, image_path, parent_email FROM students")
    students = cursor.fetchall()
    conn.close()
    return students

# Load face encodings for all known students
students = fetch_students_from_db()
known_face_encodings = []
known_face_names = []
student_data = {}

for student in students:
    name, roll_no, email, branch, year, division, image_path, parent_email = student
    try:
        student_image = face_recognition.load_image_file(image_path)
        student_face_encoding = face_recognition.face_encodings(student_image)
        if student_face_encoding:
            known_face_encodings.append(student_face_encoding[0])
            known_face_names.append(name)
            student_data[name] = {
                "roll_no": roll_no,
                "email": email,
                "parent_email": parent_email,
                "branch": branch,
                "year": year,
                "division": division,
                "image_path": image_path
            }
        else:
            print(f"No faces found in the image for {name}.")
    except Exception as e:
        print(f"Error processing image for {name}: {e}")

# Start video capture
video_capture = cv2.VideoCapture(0)
attendance_log = set()

while True:
    ret, frame = video_capture.read()
    if not ret:
        print("Failed to grab frame from the camera.")
        break

    # Resize frame for faster processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = small_frame[:, :, ::-1]

    # Detect faces in the frame
    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_names = []

    if face_locations:
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        for face_encoding in face_encodings:
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"

            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            if matches[best_match_index]:
                name = known_face_names[best_match_index]

            face_names.append(name)

            if name != "Unknown" and name not in attendance_log:
                attendance_log.add(name)
                print(f"{name} marked as Present.")

    # Draw rectangles and names around detected faces
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 255, 0), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 0.5, (255, 255, 255), 1)

    cv2.imshow('Video', frame)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        confirm_exit = input("Are you sure you want to exit? Press 'y' for Yes and 'n' for No: ").lower()
        if confirm_exit == 'y':
            absentees = set(known_face_names) - attendance_log

            rb = xlrd.open_workbook(file_path, formatting_info=True)
            wb = xl_copy(rb)
            sheet = wb.get_sheet(0)

            row = 1
            for name in known_face_names:
                student = student_data[name]
                status = "Present" if name in attendance_log else "Absent"
                sheet.write(row, 0, name)
                sheet.write(row, 1, student["roll_no"])
                sheet.write(row, 2, student["email"])
                sheet.write(row, 3, student["branch"])
                sheet.write(row, 4, student["year"])
                sheet.write(row, 5, student["division"])
                sheet.write(row, 6, today.strftime('%Y-%m-%d'))
                sheet.write(row, 7, status)


                if status == "Absent":
                    send_email(student["email"], student["parent_email"], name, subject_name, "Absent")

                row += 1

            wb.save(file_path)
            print("Attendance process completed. Program exited.")
            video_capture.release()
            cv2.destroyAllWindows()
            break
        else:
            print("Resuming video capture.")
