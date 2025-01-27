import sqlite3

# Connect to the database
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Fetch all records from the students table
cursor.execute("SELECT * FROM students")
rows = cursor.fetchall()

if rows:
    print("Student Records in Database:")
    for row in rows:
        print(f"ID: {row[0]}, Name: {row[1]}, Roll No: {row[2]}, Email: {row[3]}, Image Path: {row[4]}")
else:
    print("No records found in the database.")

# Close the database connection
conn.close()
