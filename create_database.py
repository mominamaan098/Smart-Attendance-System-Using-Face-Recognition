import sqlite3

# Connect to the database
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Drop the existing `students` table if it exists
cursor.execute("DROP TABLE IF EXISTS students")

# Create a new table for storing student details
cursor.execute('''CREATE TABLE students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    roll_no INTEGER UNIQUE NOT NULL,
    email TEXT NOT NULL,
    branch TEXT NOT NULL,
    year TEXT NOT NULL,
    division TEXT NOT NULL,
    parent_email TEXT NOT NULL,
    image_path TEXT NOT NULL
)''')

conn.commit()
conn.close()

print("Database cleared and table created successfully!")
