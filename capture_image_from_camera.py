import cv2
import os
import sqlite3

# Path to save captured images
images_directory = os.path.join(os.getcwd(), 'student_images')
os.makedirs(images_directory, exist_ok=True)

# Connect to the database
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Input student details
name = input("Enter student's name: ")
roll_no = int(input("Enter student's roll number: "))
email = input("Enter student's email: ")
branch = input("Enter student's branch: ")
year = input("Enter student's year: ")
division = input("Enter student's division: ")
parent_email = input("Enter parent's email: ")

# Initialize DroidCam (usually mapped as device ID 0 or 1)
video_capture = cv2.VideoCapture(2)  # Change device ID if needed (e.g., to 1 or 2)
if not video_capture.isOpened():
    print("Error: Unable to access the webcam. Check DroidCam settings.")
    exit()

print("Press 's' to capture the face and save the image. Press 'q' to quit.")

while True:
    ret, frame = video_capture.read()
    if not ret:
        print("Failed to grab frame. Check the DroidCam connection.")
        break

    cv2.imshow('Capture Face', frame)

    key = cv2.waitKey(1)
    if key & 0xFF == ord('s'):  # Save image
        image_path = os.path.join(images_directory, f"{name}_{roll_no}.png")
        cv2.imwrite(image_path, frame)
        print(f"Image saved at {image_path}")

        # Insert data into the database
        try:
            cursor.execute('''
                INSERT INTO students (name, roll_no, email, branch, year, division, parent_email, image_path) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (name, roll_no, email, branch, year, division, parent_email, image_path))
            conn.commit()
            print(f"Student {name}'s details added to the database.")
        except sqlite3.IntegrityError:
            print("Error: Roll number must be unique. Student not added.")
        break
    elif key & 0xFF == ord('q'):  # Quit without saving
        print("Quitting without saving.")
        break

# Release resources
video_capture.release()
cv2.destroyAllWindows()
conn.close()
