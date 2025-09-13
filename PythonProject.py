import openpyxl  # reads excel file
import tkinter as tk # Python's built in GUI
from tkinter import messagebox # model from tkinter which shows alert boxes

# -------- Student Class --------
class Student:
    def __init__(self, row): # take list og objects from row
        self.roll = str(row[0].value).strip() # takes the first column (roll number)
        self.name = row[1].value # takes second column (name)
        self.attendance = [cell.value for cell in row[2:] if cell.value in ("P", "A")] # from the remaining columns it will take the list of "P" and "A"

    def stats(self):
        total = len(self.attendance) # takes total entries (13) 13 days
        present = self.attendance.count("P") # takes the count of "P"
        absent = self.attendance.count("A") # takes the count of "A"
        percent = round((present / total) * 100, 2) if total else 0.0 # calculation part also rounds to 2 decimals 
        return total, present, absent, percent

# -------- GUI Function --------
def generate_report(): # this function takes whenever we click on generate report
    roll_input = entry.get().strip() # takes the input as roll number
    if not roll_input:
        messagebox.showwarning("Input Error", "Please enter a roll number.") # if roll no does not exist then it shows
        return

    filename = "attendance.xlsx" #reads excel file

    try:
        wb = openpyxl.load_workbook(filename) # takes the access of excel sheet
        ws = wb.active # active means current sheet

        found = False
        for row in ws.iter_rows(min_row=2): # iteration starts from 2 row since first row consists of headings
            student = Student(row) # creates a student object from each row
            if student.roll == roll_input: # it will check if the entered roll no exists in student class and equal then proceeds
                total, present, absent, percent = student.stats() # if exist then it prints the output in this format
                result = (
                    f"----- ATTENDANCE REPORT -----\n"
                    f"Name: {student.name}\n"
                    f"Roll Number: {student.roll}\n"
                    f"Total Days: {total}\n"
                    f"Days Present: {present}\n"
                    f"Days Absent: {absent}\n"
                    f"Attendance %: {percent}%"
                )
                output_label.config(text=result)
                found = True
                break

        if not found: # if roll no does not exists then it shows below line
            output_label.config(text="")
            messagebox.showinfo("Not Found", "Roll number not found.")

    except FileNotFoundError: # also shows errors if file is not found
        messagebox.showerror("File Error", f"File '{filename}' not found!")

# -------- GUI Setup --------
window = tk.Tk() # creates main window ,sets the title,size and background colour
window.title("Attendance Report Generator")
window.geometry("400x350")
window.configure(bg="#f2f2f2")

tk.Label(window, text="Enter Roll Number:", font=("Arial", 12), bg="#f2f2f2").pack(pady=10) # label to ask roll number with certain texts and bg
entry = tk.Entry(window, font=("Arial", 12), width=30) # creates a blank space and will allow entry of a roll no
entry.pack()

tk.Button(window, text="Generate Report", command=generate_report, font=("Arial", 12), bg="#2CF409", fg="white").pack(pady=10) # creates a button 

output_label = tk.Label(window, text="", font=("Courier", 10), justify="left", bg="#f2f2f2") # gives the output in this format
output_label.pack(pady=10)

window.mainloop() # keeps the GUI running until user closes