import pandas as pd
import qrcode
import os

# Load student data
df = pd.read_csv("student.csv")

# Create folder to store QR codes
folder = "Student_QR"
os.makedirs(folder, exist_ok=True)

# Generate QR for each student
for index, row in df.iterrows():
    reg_no = row["Reg No."]
    name = row["Name"]



    # Create QR (only Reg No.)
    qr = qrcode.make(reg_no)

    # Save QR image
    filename = f"{folder}/{reg_no}.png"
    qr.save(filename)

    print(f"QR generated for {name} ({reg_no})")

print("All QR codes generated successfully!")