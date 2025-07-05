import cv2
import face_recognition
import pandas as pd
import sqlite3
import csv
import smtplib
import os
import pyttsx3
from datetime import datetime
from tkinter import *
from tkinter import ttk, messagebox
from email.mime.text import MIMEText
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize database
def init_db():
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT, roll_no TEXT, department TEXT,
                  date TEXT, time TEXT, UNIQUE(name, roll_no, date))''')
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
                    'department': row['department'],
                    'email': row.get('email', '')
                }
    except FileNotFoundError:
        print("students.csv not found. Using default data")
        student_data = {
            "RAZA SHAUD.jpg": {
                "name": "RAZA SHAUD", 
                "roll_no": "35501622018", 
                "department": "EE",
                "email": "ushnishskmd@gmail.com"
            }
        }
    return student_data

# Initialize student data
student_data = load_student_data()
known_faces_dir = r'known_faces'
known_faces = []
known_names = []
known_roll_numbers = []
known_departments = []
known_emails = []

# Load known faces
for filename in os.listdir(known_faces_dir):
    if filename.endswith(('.jpg', '.png', '.jpeg')):
        image = face_recognition.load_image_file(os.path.join(known_faces_dir, filename))
        encodings = face_recognition.face_encodings(image)
        if encodings:
            known_faces.append(encodings[0])
            details = student_data.get(filename, {})
            known_names.append(details.get("name", "Unknown"))
            known_roll_numbers.append(details.get("roll_no", "N/A"))
            known_departments.append(details.get("department", "N/A"))
            known_emails.append(details.get("email", ""))

# Enhanced Image capture with real-time detection
def capture_image():
    cam = cv2.VideoCapture(0)
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    
    while True:
        ret, frame = cam.read()
        if not ret:
            break
            
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 5)
        
        for (x, y, w, h) in faces:
            cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
        
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

# Face recognition with confidence threshold
def recognize_faces(captured_image, tolerance=0.6):
    face_encodings = face_recognition.face_encodings(captured_image)
    if not face_encodings:
        return []
        
    results = []
    for encoding in face_encodings:
        matches = face_recognition.compare_faces(known_faces, encoding, tolerance)
        face_distances = face_recognition.face_distance(known_faces, encoding)
        
        if True in matches:
            best_match_index = face_distances.argmin()
            if face_distances[best_match_index] < tolerance:
                results.append((
                    known_names[best_match_index],
                    known_roll_numbers[best_match_index],
                    known_departments[best_match_index],
                    known_emails[best_match_index]
                ))
    return results

# Database operations with duplicate prevention
def mark_attendance_db(student_name, roll_no, department):
    conn = sqlite3.connect('attendance.db')
    c = conn.cursor()
    
    try:
        c.execute("""INSERT OR IGNORE INTO attendance 
                  (name, roll_no, department, date, time) 
                  VALUES (?, ?, ?, ?, ?)""",
                  (student_name, roll_no, department,
                   datetime.now().strftime("%Y-%m-%d"),
                   datetime.now().strftime("%H:%M:%S")))
        conn.commit()
        return c.rowcount > 0
    except sqlite3.Error as e:
        print(f"Database error: {e}")
        return False
    finally:
        conn.close()

# Enhanced Report generation
def generate_report(date=None, department=None, output_format='excel'):
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
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if output_format == 'excel':
            report_name = f"Attendance_Report_{timestamp}.xlsx"
            df.to_excel(report_name, index=False)
        else:
            report_name = f"Attendance_Report_{timestamp}.csv"
            df.to_csv(report_name, index=False)
        return report_name
    return None

# Secure Email function
def send_email(to, subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = os.getenv("GMAIL_USER")
    msg['To'] = "ushnishskmd@gmail.com"
    
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.login(
                os.getenv("GMAIL_USER"),
                os.getenv("GMAIL_APP_PASSWORD")
            )
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Email failed: {e}")
        return False

# GUI Application with all enhancements
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Face Recognition Attendance System")
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        self.init_db()
        self.create_widgets()
        
    def init_db(self):
        conn = sqlite3.connect('attendance.db')
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS attendance
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      name TEXT, roll_no TEXT, department TEXT,
                      date TEXT, time TEXT, UNIQUE(name, roll_no, date))''')
        conn.commit()
        conn.close()
        
    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(N, S, E, W))
        
        ttk.Label(main_frame, text="Advanced Attendance System", 
                 font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        
        ttk.Button(main_frame, text="Take Attendance", 
                  command=self.take_attendance).grid(row=1, column=0, pady=10, sticky=E)
        ttk.Button(main_frame, text="Generate Report", 
                  command=self.show_report_options).grid(row=1, column=1, pady=10, sticky=W)
        ttk.Button(main_frame, text="Exit", 
                  command=self.root.quit).grid(row=2, column=0, columnspan=2, pady=10)
        
    def show_report_options(self):
        report_window = Toplevel(self.root)
        report_window.title("Generate Report")
        
        ttk.Label(report_window, text="Report Options", font=("Arial", 12)).grid(row=0, column=0, columnspan=2, pady=10)
        
        ttk.Label(report_window, text="Date (YYYY-MM-DD):").grid(row=1, column=0, sticky=E)
        self.date_entry = ttk.Entry(report_window)
        self.date_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(report_window, text="Department:").grid(row=2, column=0, sticky=E)
        self.dept_entry = ttk.Entry(report_window)
        self.dept_entry.grid(row=2, column=1, padx=5, pady=5)
        
        self.report_type = StringVar(value='excel')
        ttk.Radiobutton(report_window, text="Excel", variable=self.report_type, value='excel').grid(row=3, column=0)
        ttk.Radiobutton(report_window, text="CSV", variable=self.report_type, value='csv').grid(row=3, column=1)
        
        ttk.Button(report_window, text="Generate", 
                  command=self.generate_report).grid(row=4, column=0, columnspan=2, pady=10)
        
    def generate_report(self):
        date = self.date_entry.get() or None
        department = self.dept_entry.get() or None
        output_format = self.report_type.get()
        
        report_file = generate_report(date, department, output_format)
        if report_file:
            messagebox.showinfo("Success", f"Report generated: {report_file}")
            os.startfile(report_file)
        else:
            messagebox.showwarning("Warning", "No matching records found")
        
    def take_attendance(self):
        if not self.calibrate_camera():
            messagebox.showerror("Error", "Camera calibration failed. Please check your camera.")
            return
            
        image = capture_image()
        if image is None:
            messagebox.showwarning("Warning", "No image captured")
            return
            
        recognized = recognize_faces(image)
        if not recognized:
            messagebox.showwarning("Warning", "No faces recognized")
            return
            
        for name, roll_no, dept, email in recognized:
            # if mark_attendance_db(name, roll_no, dept):
            #     self.speak(f"Attendance marked for {name}")
            #     messagebox.showinfo("Success", f"Attendance marked for {name}")
            if mark_attendance_db(name, roll_no, dept):
                self.speak(f"Hello {name}")
                messagebox.showinfo("Success", f"Attendance marked for {name}")

                
                # Send email notification
                if email:
                    email_body = f"""Attendance recorded for:
                    Name: {name}
                    Roll No: {roll_no}
                    Department: {dept}
                    Time: {datetime.now().strftime('%H:%M:%S')}"""
                    
                    if send_email(email, "Attendance Recorded", email_body):
                        print(f"Email sent to {email}")
            else:
                messagebox.showinfo("Info", f"Attendance already marked today for {name}")
    
    def calibrate_camera(self):
        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            return False
            
        ret, frame = cam.read()
        cam.release()
        
        if not ret or frame is None:
            return False
            
        blur_value = cv2.Laplacian(frame, cv2.CV_64F).var()
        return blur_value > 100

# Main execution
if __name__ == "__main__":
    init_db()
    root = Tk()
    root.geometry("400x200")
    root.resizable(False, False)
    app = AttendanceApp(root)
    root.mainloop()