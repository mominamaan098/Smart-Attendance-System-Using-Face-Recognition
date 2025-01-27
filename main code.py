import face_recognition
import cv2
import numpy as np
import os
import xlwt
from datetime import datetime
import xlrd
from xlutils.copy import copy as xl_copy
from email_notification import send_email  # Import the send_email function from the email_notification file

# Prompt for subject name
subject_name = input("Enter the subject name for attendance: ")

# Define target directory and create file path with current date, day, and subject name
target_directory = r"C:\Users\momin\OneDrive\Desktop\SAS"
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
    sheet1.write(0, 0, 'Name')
    sheet1.write(0, 1, 'Date')
    sheet1.write(0, 2, 'Status')
    wb.save(file_path)
    print(f"Excel file created successfully for {subject_name} at: {file_path}")

# Initialize known images
CurrentFolder = os.getcwd()
image1_path = os.path.join(CurrentFolder, 'Amaan_62.png')
image2_path = os.path.join(CurrentFolder, 'Nidhish_61.png')
image3_path = os.path.join(CurrentFolder, 'Aawaiz_38.png')
image4_path = os.path.join(CurrentFolder, 'Ajinkya_22.png')

# Load face encodings for all known images
person1_name = "Momin Amaan"
person1_image = face_recognition.load_image_file(image1_path)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

person2_name = "Nidhish"
person2_image = face_recognition.load_image_file(image2_path)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

person3_name = "Aawaiz"
person3_image = face_recognition.load_image_file(image3_path)
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]

person4_name = "Ajinkya"
person4_image = face_recognition.load_image_file(image4_path)
person4_face_encoding = face_recognition.face_encodings(person4_image)[0]

# Add all known face encodings and names to the lists
known_face_encodings = [
    person1_face_encoding, 
    person2_face_encoding, 
    person3_face_encoding, 
    person4_face_encoding
]
known_face_names = [
    person1_name, 
    person2_name, 
    person3_name, 
    person4_name
]

# Start video capture
video_capture = cv2.VideoCapture(0)
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True
attendance_log = set()

while True:
    ret, frame = video_capture.read()
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = small_frame[:, :, ::-1]

    if process_this_frame:
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
        face_names = []

        for face_encoding in face_encodings:
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)
            if matches[best_match_index]:
                name = known_face_names[best_match_index]

            face_names.append(name)

            # Mark attendance if recognized
            if name != "Unknown" and name not in attendance_log:
                # Open and edit the workbook to mark attendance
                rb = xlrd.open_workbook(file_path, formatting_info=True)
                wb = xl_copy(rb)
                sheet = wb.get_sheet(0)
                
                # Write name and status
                row = len(attendance_log) + 1
                sheet.write(row, 0, name)
                sheet.write(row, 1, today.strftime('%Y-%m-%d'))
                sheet.write(row, 2, "Present")
                
                # Save the updated attendance file
                wb.save(file_path)
                attendance_log.add(name)
                print(f"{name}'s attendance marked as present.")

                # Send email notification to the student
                student_email = "student_email@example.com"  # Replace with actual student email
                send_email(student_email, name)  # Send email to the student
    
    process_this_frame = not process_this_frame

    # Display results
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Display the resulting frame
    cv2.imshow('Video', frame)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release the capture
video_capture.release()
cv2.destroyAllWindows()