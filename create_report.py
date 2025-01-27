import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

# Step 1: Read Data
def find_matching_file(directory, keyword):
    matching_files = [f for f in os.listdir(directory) if keyword.lower() in f.lower() and f.endswith(('.xls', '.xlsx'))]
    if not matching_files:
        print(f"No file found containing '{keyword}'.")
        return None
    if len(matching_files) > 1:
        print(f"Multiple files found: {', '.join(matching_files)}. Using the first one.")
    return os.path.join(directory, matching_files[0])

def read_attendance_sheet(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

# Step 2: Analyze Attendance
def analyze_attendance(data):
    present_count = data[data['Status'] == 'Present'].shape[0]
    absent_count = data[data['Status'] == 'Absent'].shape[0]
    total_students = data.shape[0]
    attendance_percentage = (present_count / total_students) * 100
    return present_count, absent_count, attendance_percentage

# Step 3: Create Graphs
def create_graphs(present, absent, file_name):
    labels = ['Present', 'Absent']
    sizes = [present, absent]
    colors = ['#4CAF50', '#FF5252']
    
    plt.figure(figsize=(8, 8))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=140)
    plt.title(f'Attendance Distribution for {file_name}')
    graph_path = f"{file_name}_attendance_chart.png"
    plt.savefig(graph_path)
    plt.close()
    return graph_path

# Step 4: Generate PDF
def generate_pdf_report(file_name, data, present, absent, percentage, graph_path):
    pdf = FPDF()
    pdf.add_page()
    
    # Title
    pdf.set_font("Arial", size=16)
    pdf.cell(200, 10, txt="Attendance Report By SmartMark", ln=True, align='C')
    
    # Greeting
    pdf.set_font("Arial", size=12)
    pdf.ln(10)
    pdf.cell(200, 10, txt="Respected Authorities,", ln=True)
    pdf.cell(200, 10, txt="Please find the attendance details below:", ln=True)
    
    # Summary
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Subject: {file_name}", ln=True)
    pdf.cell(200, 10, txt=f"Total Students: {data.shape[0]}", ln=True)
    pdf.cell(200, 10, txt=f"Present: {present}", ln=True)
    pdf.cell(200, 10, txt=f"Absent: {absent}", ln=True)
    pdf.cell(200, 10, txt=f"Attendance Percentage: {percentage:.2f}%", ln=True)
    
    # Add Graph
    pdf.ln(10)
    pdf.image(graph_path, x=50, y=None, w=100)
    
    # Save PDF
    pdf_file_name = f"{file_name}_attendance_report.pdf"
    pdf.output(pdf_file_name)
    print(f"PDF report generated: {pdf_file_name}")
    return pdf_file_name

# Step 5: Send Email with PDF Attachment
def send_email_with_attachment(pdf_file_name, recipient_email):
    # Email settings (replace with actual details)
    sender_email = "mominamaan001@gmail.com"  # Replace with your email
    sender_password = "nxvh olbe mwiy fvde"     # Replace with your email password (or use app password)
    
    # Create message object
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Attendance Report"
    
    # Add body
    body = "Please find attached the attendance report."
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach the PDF
    attachment = MIMEBase('application', 'octet-stream')
    try:
        with open(pdf_file_name, 'rb') as file:
            attachment.set_payload(file.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename={pdf_file_name}')
            msg.attach(attachment)
    except Exception as e:
        print(f"Error attaching file: {e}")
        return

    # Send the email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Step 6: Main Function
def main():
    directory = r"C:\Users\momin\OneDrive\Desktop\smart"  # Replace with your directory
    keyword = input("Enter the file name (without extension): ")
    
    file_path = find_matching_file(directory, keyword)
    if file_path:
        data = read_attendance_sheet(file_path)
        if data is not None:
            present, absent, percentage = analyze_attendance(data)
            graph_path = create_graphs(present, absent, keyword)
            pdf_file_name = generate_pdf_report(keyword, data, present, absent, percentage, graph_path)
            
            # List of recipient emails
            recipients = ['mominamaan619@gmail.com', 'Patelaawaiz2003@gmail.com']
            for recipient in recipients:
                send_email_with_attachment(pdf_file_name, recipient)

if __name__ == "__main__":
    main()
