## My first Babysteps in Python. 
## I learned to write this code with partially use of ChatGPT 3.5.
## This little tool converts a csv file to a xlsx file
## The tables are specified to an explicit file for my music school
## You have to alter the table names in line 25 if you want to customize it for your purpose

### Importing the necessary modules

import pandas as pd
import os
from tkinter import Tk, Label, Button, filedialog, messagebox
from PIL import Image, ImageTk

## Input

def convert_csv_to_xlsx(input_file, output_file):
    try:
        # Check if the input file exists
        if not os.path.isfile(input_file):
            messagebox.showerror("Error", f"The file '{input_file}' does not exist.")
            return

        # Read the CSV file
        df = pd.read_csv(input_file)
        messagebox.showinfo("Info", f"Successfully read the CSV file: {input_file}")

        # Ensure the column has the right order and names
        expected_columns = ["Lehrer Vorname", "Lehrer Nachname", "Sch. Nachname", "Sch. Vorname", "Stück Name", "Stück Komponist", "Stück Länge"]
        if len(df.columns) != len(expected_columns):
            messagebox.showerror("Error", f"The input file does not have the expected number of columns ({len(expected_columns)}).")
            messagebox.showinfo("Info", f"Found columns: {df.columns}")
            return
        df.columns = expected_columns
        messagebox.showinfo("Info", "Column names have been set correctly.")

        # Save to XLSX file
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Info", f"Conversion completed: {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        input_label.config(text=f"Input File: {file_path}")
        input_label.file_path = file_path

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        output_label.config(text=f"Output File: {file_path}")
        output_label.file_path = file_path

def run_conversion():
    input_file = getattr(input_label, 'file_path', None)
    output_file = getattr(output_label, 'file_path', None)

    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    convert_csv_to_xlsx(input_file, output_file)

## frontend: Hovering over buttons

def on_enter(event):
    event.widget.config(bg="lightblue")

def on_leave(event):
    event.widget.config(bg="goldenrod")

def on_enter_convert(event):
    event.widget.config(bg="lightblue")

def on_leave_convert(event):
    event.widget.config(bg="green")

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        app.destroy()   

## frontend

app = Tk()
app.title("CSV to XLSX Converter")
app.geometry("750x550")
app.configure(bg="azure")

title_label = Label(app, text="CSV to XLSX Converter", font=("Helvetica", 20), bg="lightblue", bd=5)
title_label.pack(fill="both", expand=True, pady=20)

input_label = Label(app, text="Input File: Not Selected", width=90, height=4)
input_label.pack(pady=10)

output_label = Label(app, text="Output File: Not Selected", width=90, height=4)
output_label.pack(pady=10)

select_input_button = Button(app, text="Select Input .CSV File", command=select_input_file, width=40, height=4, bg="goldenrod")
select_input_button.pack(pady=10)
select_input_button.bind("<Enter>", on_enter)
select_input_button.bind("<Leave>", on_leave)

select_output_button = Button(app, text="Select Output .XLSX File", command=select_output_file, width=40, height=4, bg="goldenrod")
select_output_button.pack(pady=10)
select_output_button.bind("<Enter>", on_enter)
select_output_button.bind("<Leave>", on_leave)

convert_button = Button(app, text="Convert", font=("Helvetica", 20), command=run_conversion, width=40, height=4, bg="green")
convert_button.pack(pady=20)
convert_button.bind("<Enter>", on_enter_convert)
convert_button.bind("<Leave>", on_leave_convert)

## Quit Message Box

app.protocol("WM_DELETE_WINDOW", on_closing)

## Necessary for Tkinter

app.mainloop()
