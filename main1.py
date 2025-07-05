import cv2
import face_recognition
import pandas as pd
from datetime import datetime
import os

# Directory containing known face images
known_faces_dir = r'C:\Users\razas\Documents\Attendance-System\known_faces'

# Load known faces with additional details
known_faces = []
known_names = []
known_roll_numbers = []
known_departments = []

# Example student data (modify with your actual students)
student_data = {
    "RAZA SHAUD.jpg": {"name": "RAZA SHAUD", "Roll No": "35501622018", "department": "EE"},
    "RAJEEV KUMAR.jpg": {"name": "RAJEEV KUMAR", "Roll No": "35501622018", "department": "EE"},
}

for filename in os.listdir(known_faces_dir):
    if filename.endswith(('.jpg', '.png')):
        image = face_recognition.load_image_file(os.path.join(known_faces_dir, filename))
        encoding = face_recognition.face_encodings(image)[0]
        known_faces.append(encoding)
        
        # Get student details
        details = student_data.get(filename, {})
        known_names.append(details.get("name", "Unknown"))
        known_roll_numbers.append(details.get("roll_no", "N/A"))
        known_departments.append(details.get("department", "N/A"))

def capture_image():
    """Capture image from webcam"""
    cam = cv2.VideoCapture(0)
    while True:
        ret, frame = cam.read()
        if not ret:
            print("Failed to grab frame")
            break
        cv2.imshow('Press Space to capture', frame)
        if cv2.waitKey(1) & 0xFF == ord(' '):
            break
    cam.release()
    cv2.destroyAllWindows()
    return frame if ret else None

def recognize_face(captured_image):
    """Recognize face and return student details"""
    face_encodings = face_recognition.face_encodings(captured_image)
    if not face_encodings:
        print("No faces detected.")
        return None, None, None
    
    captured_encoding = face_encodings[0]
    matches = face_recognition.compare_faces(known_faces, captured_encoding)
    
    if True in matches:
        first_match_index = matches.index(True)
        return (
            known_names[first_match_index],
            known_roll_numbers[first_match_index],
            known_departments[first_match_index]
        )
    return None, None, None

def mark_attendance(student_name, roll_no, department, file='attendance.xlsx'):
    """Mark attendance in Excel file"""
    now = datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M:%S")
    
    try:
        df = pd.read_excel(file)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Name", "Roll No", "Department", "Date", "Time"])
    
    new_record = {
        "Name": student_name,
        "Roll No": roll_no,
        "Department": department,
        "Date": current_date,
        "Time": current_time
    }
    
    df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
    df.to_excel(file, index=False)

def main():
    """Main execution flow"""
    print("Press SPACE to capture image...")
    image = capture_image()
    if image is None:
        print("Failed to capture image")
        return
    
    name, roll_no, dept = recognize_face(image)
    if name is None:
        print("Student not recognized!")
        return
    
    mark_attendance(name, roll_no, dept)
    print(f"Attendance marked for {name} (Roll No: {roll_no}, Dept: {dept})")

if __name__ == "__main__":
    main()