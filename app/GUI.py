import tkinter as tk
from tkinter import filedialog, messagebox
from app import extract_job_costing_from_raw_excel  # <-- Update this to your actual script/module name

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xls *.xlsx")]
    )
    if file_path:
        file_var.set(file_path)

def process_file():
    input_path = file_var.get()
    if not input_path:
        messagebox.showwarning("No File", "Please select a file first.")
        return

    try:
        output_path = "formatted_output.xlsx"
        extract_job_costing_from_raw_excel(input_path, output_path)
        messagebox.showinfo("Success", f"Formatted file saved as {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# Set up GUI
root = tk.Tk()
root.title("Job Costing Formatter")
root.geometry("500x200")

file_var = tk.StringVar()

# UI Layout
label = tk.Label(root, text="Select Payroll Excel File", font=("Arial", 12))
label.pack(pady=10)

file_entry = tk.Entry(root, textvariable=file_var, width=50)
file_entry.pack(pady=5)

browse_button = tk.Button(root, text="Browse", command=select_file)
browse_button.pack(pady=5)

process_button = tk.Button(root, text="Process File", command=process_file)
process_button.pack(pady=15)

root.mainloop()
