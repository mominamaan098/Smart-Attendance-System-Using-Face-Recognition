import sqlite3

# Connect to the database
conn = sqlite3.connect('attendance.db')
cursor = conn.cursor()

# Function to display data from a table
def display_table_data(table_name):
    print(f"Data in table '{table_name}':")
    try:
        cursor.execute(f"SELECT * FROM {table_name}")
        rows = cursor.fetchall()
        for row in rows:
            print(row)
        if not rows:
            print(f"No data found in '{table_name}'.\n")
    except sqlite3.OperationalError as e:
        print(f"Error accessing table '{table_name}': {e}\n")

# List of tables to display
tables = ['students', 'attendance', 'faculty']

# Display data for each table
for table in tables:
    display_table_data(table)

# Close the connection
conn.close()

import sqlite3

# # Connect to the database
# conn = sqlite3.connect('attendance.db')
# cursor = conn.cursor()

# # Clear only the data from the `students` table
# cursor.execute("DELETE FROM students")

# # Optional: Reset the `id` auto-increment sequence (if needed)
# cursor.execute("DELETE FROM sqlite_sequence WHERE name='students'")

# conn.commit()
# conn.close()

# print("Student data cleared successfully!")

