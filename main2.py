import cv2
import face_recognition
import pandas as pd
import sqlite3
import csv
import smtplib
import os
from datetime import datetime
from tkinter import *
from tkinter import messagebox
from email.mime.text import MIMEText

# Initialize database
def init_db():
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT, roll_no TEXT, department TEXT,
                  date TEXT, time TEXT)''')
    conn.commit()
    conn.close()

# Load student data from CSV
def load_student_data():
    student_data = {}
    try:
        with open('students.csv', mode='r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                student_data[row['image_file']] = {
                    'name': row['name'],
                    'roll_no': row['roll_no'],
                    'department': row['department']
                }
    except FileNotFoundError:
        print("students.csv not found. Using default data")
        student_data = {
            "RAZA SHAUD.jpg": {"name": "RAZA SHAUD", "Roll No": "35501622018", "department": "EE"}
        }
    return student_data

# Initialize student data
student_data = load_student_data()
known_faces_dir = r'C:\Users\razas\Documents\Attendance-System\known_faces'
known_faces = []
known_names = []
known_roll_numbers = []
known_departments = []

# Load known faces
for filename in os.listdir(known_faces_dir):
    if filename.endswith(('.jpg', '.png')):
        image = face_recognition.load_image_file(os.path.join(known_faces_dir, filename))
        encoding = face_recognition.face_encodings(image)[0]
        known_faces.append(encoding)
        details = student_data.get(filename, {})
        known_names.append(details.get("name", "Unknown"))
        known_roll_numbers.append(details.get("roll_no", "N/A"))
        known_departments.append(details.get("department", "N/A"))

# Image capture with real-time detection
def capture_image():
    cam = cv2.VideoCapture(0)
    while True:
        ret, frame = cam.read()
        if not ret:
            break
            
        face_locations = face_recognition.face_locations(frame)
        for (top, right, bottom, left) in face_locations:
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
        
        cv2.imshow('Press SPACE to capture (ESC to exit)', frame)
        key = cv2.waitKey(1)
        if key == ord(' '):
            break
        elif key == 27:
            frame = None
            break
            
    cam.release()
    cv2.destroyAllWindows()
    return frame

# Face recognition
def recognize_faces(captured_image):
    face_encodings = face_recognition.face_encodings(captured_image)
    if not face_encodings:
        return []
        
    results = []
    for encoding in face_encodings:
        matches = face_recognition.compare_faces(known_faces, encoding)
        if True in matches:
            first_match_index = matches.index(True)
            results.append((
                known_names[first_match_index],
                known_roll_numbers[first_match_index],
                known_departments[first_match_index]
            ))
    return results

# Database operations
def mark_attendance_db(student_name, roll_no, department):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute("INSERT INTO attendance (name, roll_no, department, date, time) VALUES (?, ?, ?, ?, ?)",
              (student_name, roll_no, department,
               datetime.now().strftime("%Y-%m-%d"),
               datetime.now().strftime("%H:%M:%S")))
    conn.commit()
    conn.close()

# Report generation
def generate_report(date=None, department=None):
    conn = sqlite3.connect('attendance.db')
    query = "SELECT * FROM attendance"
    conditions = []
    if date:
        conditions.append(f"date = '{date}'")
    if department:
        conditions.append(f"department = '{department}'")
    
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    df = pd.read_sql(query, conn)
    conn.close()
    
    if not df.empty:
        report_name = f"Attendance_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        df.to_excel(report_name, index=False)
        return report_name
    return None

def send_email(to, subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = "rshaudx3@gmail.com"
    msg['To'] = "ushnishskmd@gmail.com"
    
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.ehlo()
            server.starttls()
            # Provide both username and password arguments
            server.login("rshaudx3@gmail.com", "nali ggjm zcyr wwwg")  # Fixed line
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Email failed: {e}")
        return False
    
# GUI Application
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Face Recognition Attendance System")
        self.init_db()
        
        Label(root, text="Attendance System", font=("Arial", 16)).pack(pady=20)
        
        Button(root, text="Take Attendance", command=self.take_attendance).pack(pady=10)
        Button(root, text="Generate Report", command=self.generate_report).pack(pady=10)
        Button(root, text="Exit", command=root.quit).pack(pady=10)
        
    def init_db(self):
        conn = sqlite3.connect('attendance.db')
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS attendance
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      name TEXT, roll_no TEXT, department TEXT,
                      date TEXT, time TEXT)''')
        conn.commit()
        conn.close()
        
    def take_attendance(self):
        image = capture_image()
        if image is None:
            messagebox.showwarning("Warning", "No image captured")
            return
            
        recognized = recognize_faces(image)
        if not recognized:
            messagebox.showwarning("Warning", "No faces recognized")
            return
            
        for name, roll_no, dept in recognized:
            mark_attendance_db(name, roll_no, dept)
            messagebox.showinfo("Success", f"Attendance marked for {name}")
            
            # Send email notification
            email_body = f"Attendance recorded for {name} ({roll_no}) at {datetime.now().strftime('%H:%M:%S')}"
            if send_email("admin@example.com", "Attendance Recorded", email_body):
                print("Email notification sent")
        
    def generate_report(self):
        report_file = generate_report()
        if report_file:
            messagebox.showinfo("Success", f"Report generated: {report_file}")
        else:
            messagebox.showwarning("Warning", "No attendance records found")

# Main execution
if __name__ == "__main__":
    init_db()
    root = Tk()
    app = AttendanceApp(root)
    root.mainloop()