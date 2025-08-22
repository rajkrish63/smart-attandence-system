from openpyxl import Workbook

# Create a new workbook and select the active sheet
wb = Workbook()
sheet = wb.active



# Set column headers
sheet.append(["S/No", "Roll No", "Student Name", "Parent No"])

# Add sample student data
students = [
    (1,"101", "RAJESH", "9876543210"),
   
]

# Insert student records
for student in students:
    sheet.append(student)

# Save the file
wb.save("data.xlsx")
print("✅ data.xlsx created successfully!")
