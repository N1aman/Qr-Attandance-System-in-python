from tkinter.constants import GROOVE, RAISED, RIDGE
import cv2
import pyzbar.pyzbar as pyzbar
import time
from datetime import date, datetime
import tkinter as tk 
from tkinter import Frame, ttk, messagebox
from tkinter import *
import os
import pandas as pd
from PIL import Image, ImageTk

# Initialize the main window
window = tk.Tk()
window.title('Attendance System ')
window.geometry('900x600')  
window.configure(bg='lightblue')

# Variables for user input
year = tk.StringVar()      
branch = tk.StringVar()
sec = tk.StringVar() 
period = tk.StringVar()

# GUI Title
title = tk.Label(window, text="Attendance System Using Qr Code", bd=10, relief=tk.GROOVE, font=("times new roman", 40), bg="lavender", fg="black")
title.pack(side=tk.TOP, fill=tk.X)

# Manage Frame
Manage_Frame = Frame(window, bg="lavender")
Manage_Frame.place(x=0, y=80, width=480, height=530)

# Input fields
ttk.Label(window, text="Year", background="lavender", foreground="black", font=("Times New Roman", 15)).place(x=100, y=150)
combo_search = ttk.Combobox(window, textvariable=year, width=10, font=("times new roman", 13), state='readonly')
combo_search['values'] = ('1', '2', '3', '4') 
combo_search.place(x=250, y=150)

ttk.Label(window, text="Branch", background="lavender", foreground="black", font=("Times New Roman", 15)).place(x=100, y=200)
combo_search = ttk.Combobox(window, textvariable=branch, width=10, font=("times new roman", 13), state='readonly')
combo_search['values'] = ("CSE", "ECE", "EEE", "IT", "MECH", "ECM")
combo_search.place(x=250, y=200)

ttk.Label(window, text="Section", background="lavender", foreground="black", font=("Times New Roman", 15)).place(x=100, y=250)
combo_search = ttk.Combobox(window, textvariable=sec, width=10, font=("times new roman", 13), state='readonly')
combo_search['values'] = ('A', 'B', 'C', 'D')
combo_search.place(x=250, y=250)

ttk.Label(window, text="Period", background="lavender", foreground="black", font=("Times New Roman", 15)).place(x=100, y=300)
combo_search = ttk.Combobox(window, textvariable=period, width=10, font=("times new roman", 13), state='readonly')
combo_search['values'] = ('1', '2', '3', '4', '5', '6', '7')
combo_search.place(x=250, y=300)

# Right Frame (for image)
right_frame = Frame(window, bg="white")
right_frame.place(x=480, y=80, width=420, height=530)

img_path = "QR_Attendance/bg.jpg"
# Load and display image
try:
    # Replace "your_image.jpg" with your image file path
    img = Image.open(img_path)
    img = img.resize((400, 500), Image.LANCZOS)  # Resize to fit
    img_tk = ImageTk.PhotoImage(img)
    
    img_label = tk.Label(right_frame, image=img_tk, bg="white")
    img_label.image = img_tk  # Keep a reference
    img_label.pack(padx=10, pady=10)
    
except Exception as e:
    # Fallback if image can't be loaded
    fallback_label = tk.Label(right_frame, text="Attendance System\n(Image Here)", 
                             bg="white", fg="gray", font=("Times New Roman", 24))
    fallback_label.pack(expand=True)

def checkk():
    if year.get() and branch.get() and period.get() and sec.get():
        window.destroy()
    else:
        messagebox.showwarning("Warning", "All fields required!!")

exit_button = tk.Button(window, width=13, text="Submit", font=("Times New Roman", 15), command=checkk, bd=2, relief=RIDGE)
exit_button.place(x=300, y=380)

# Start the Tkinter main loop
window.mainloop()

# Start video capture
cap = cv2.VideoCapture(0)
names = []
today = date.today()
d = today.strftime("%b-%d-%Y")

# Create a DataFrame to store attendance data
attendance_data = pd.DataFrame(columns=["Reg No.", "Class & Sec", "Year", "Period", "In Time"])

def enterData(z):   
    if z not in names:
        it = datetime.now()
        names.append(z)
        intime = it.strftime("%H:%M:%S")
        
        # Add data to DataFrame
        global attendance_data
        new_entry = pd.DataFrame({
            "Reg No.": [z],
            "Class & Sec": [f"{branch.get()}-{sec.get()}"],
            "Year": [year.get()],
            "Period": [period.get()],
            "In Time": [intime]
        })
        attendance_data = pd.concat([attendance_data, new_entry], ignore_index=True)
        
        # Save to Excel file (overwrites each time)
        attendance_data.to_excel(d + '.xlsx', index=False)
        
    return names

print('Reading...')

def checkData(data):
    data = data.decode('utf-8')  # Convert bytes to string
    if data in names:
        print('Already Present')
    else:
        print(f'\n{len(names)+1}\nPresent...')
        enterData(data)

while True:
    _, frame = cap.read()         
    decodedObjects = pyzbar.decode(frame)
    for obj in decodedObjects:
        checkData(obj.data)
        time.sleep(1)
       
    cv2.imshow("Frame", frame)

    if cv2.waitKey(1) & 0xFF == ord('g'):
        cv2.destroyAllWindows()
        break

# Final save when exiting
attendance_data.to_excel(d + '.xlsx', index=False)
print("Attendance data saved to Excel.")
