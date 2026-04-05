from tkinter.constants import GROOVE, RIDGE
import cv2
import pyzbar.pyzbar as pyzbar
import time
from datetime import date, datetime
import tkinter as tk 
from tkinter import Frame, ttk, messagebox
import pandas as pd
from PIL import Image, ImageTk

# ---------------- GUI PART ---------------- #
window = tk.Tk()
window.title('Attendance System')
window.geometry('900x600')  
window.configure(bg='lightblue')

year = tk.StringVar()      
branch = tk.StringVar()
sec = tk.StringVar() 
period = tk.StringVar()

title = tk.Label(window, text="Attendance System Using QR Code",
                 bd=10, relief=tk.GROOVE,
                 font=("times new roman", 30),
                 bg="lavender", fg="black")
title.pack(side=tk.TOP, fill=tk.X)

Manage_Frame = Frame(window, bg="lavender")
Manage_Frame.place(x=0, y=80, width=480, height=530)

# Inputs
ttk.Label(window, text="Year", background="lavender", font=("Times", 15)).place(x=100, y=150)
ttk.Combobox(window, textvariable=year, values=('1','2','3','4'),
             state='readonly').place(x=250, y=150)

ttk.Label(window, text="Branch", background="lavender", font=("Times", 15)).place(x=100, y=200)
ttk.Combobox(window, textvariable=branch,
             values=("CSE","ECE","EEE","IT","MECH","ECM"),
             state='readonly').place(x=250, y=200)

ttk.Label(window, text="Section", background="lavender", font=("Times", 15)).place(x=100, y=250)
ttk.Combobox(window, textvariable=sec,
             values=('A','B','C','D'),
             state='readonly').place(x=250, y=250)

ttk.Label(window, text="Period", background="lavender", font=("Times", 15)).place(x=100, y=300)
ttk.Combobox(window, textvariable=period,
             values=('1','2','3','4','5','6','7'),
             state='readonly').place(x=250, y=300)

def checkk():
    if year.get() and branch.get() and period.get() and sec.get():
        window.destroy()
    else:
        messagebox.showwarning("Warning", "All fields required!!")

tk.Button(window, text="Submit", font=("Times", 15),
          command=checkk, bd=2, relief=RIDGE).place(x=300, y=380)

window.mainloop()

# ---------------- DATA PART ---------------- #

# Load student database - Ensure student.csv exists in the folder
try:
    student_df = pd.read_csv("student.csv")
    student_dict = dict(zip(student_df["Reg No."], student_df["Name"]))
except FileNotFoundError:
    print("Error: student.csv not found! Please create it with 'Reg No.' and 'Name' columns.")
    exit()

cap = cv2.VideoCapture(0)
names = []

today = date.today()
d = today.strftime("%b-%d-%Y")

attendance_data = pd.DataFrame(columns=[
    "Reg No.", "Name", "Class & Sec", "Year", "Period", "In Time"
])

# ---------------- FUNCTIONS ---------------- #

def enterData(reg_no):   
    global attendance_data
    if reg_no not in names:
        it = datetime.now()
        names.append(reg_no)
        intime = it.strftime("%H:%M:%S")

        student_name = student_dict.get(reg_no, "Unknown Student")

        new_entry = pd.DataFrame({
            "Reg No.": [reg_no],
            "Name": [student_name],
            "Class & Sec": [f"{branch.get()}-{sec.get()}"],
            "Year": [year.get()],
            "Period": [period.get()],
            "In Time": [intime]
        })

        attendance_data = pd.concat([attendance_data, new_entry], ignore_index=True)
        # Using .xlsx might require 'openpyxl' installed: pip install openpyxl
        attendance_data.to_excel(d + '.xlsx', index=False)
        print(f"{student_name} marked present")

def extract_id(data):
    data = data.decode('utf-8')
    # If QR contains URL → extract ID
    if data.startswith("http"):
        data = data.split("/")[-1]
    return data

def checkData(data):
    reg_no = extract_id(data)
    if reg_no in names:
        print(f"ID {reg_no}: Already Present")
    else:
        print(f"Attendance Count: {len(names)+1}")
        enterData(reg_no)

# ---------------- CAMERA LOOP ---------------- #

print("Scanning QR... Press 'g' to stop")

while True:
    ret, frame = cap.read()

    # SAFETY CHECK: If the camera fails to return a frame, skip this loop iteration
    if not ret or frame is None:
        continue

    decodedObjects = pyzbar.decode(frame)

    for obj in decodedObjects:
        # Visual feedback: Draw a box around the QR code
        (x, y, w, h) = obj.rect
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
        
        checkData(obj.data)
        
        # Short pause to prevent scanning the same code 100 times per second
        # We show the frame FIRST so the window doesn't look frozen
        cv2.imshow("QR Scanner", frame)
        cv2.waitKey(1000) 

    cv2.imshow("QR Scanner", frame)

    # Press 'g' to stop
    if cv2.waitKey(1) & 0xFF == ord('g'):
        break

cap.release()
cv2.destroyAllWindows()

# Final save to ensure data integrity
attendance_data.to_excel(d + '.xlsx', index=False)
print(f"Attendance saved successfully to {d}.xlsx!")