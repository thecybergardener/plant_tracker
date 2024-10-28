import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
from openpyxl import load_workbook
from datetime import datetime
import os
import shutil
from dateutil.relativedelta import relativedelta # To handle date difference in years, months, days
import config

# Global variable to called plant data
plant_data = {}

# Function to get plant details from the "Basic info" sheet
def get_plant_info():
    wb = load_workbook(config.EXCEL_FILE_PATH)  # Use path from config
    sheet = wb['Basic info']
    plants = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Assumes plant names start from row 2
        plant_name = row[1]  # Unique plant nickname
        plant_species = row[0]  # Actual plant species
        purchase_date = row[2]  # Purchase date
        description = row[3] if row[3] else ""  # Description
        plants[plant_name] = {"plant": plant_species, "purchase_date": purchase_date, "description": description}
    return plants

# Function to calculate years, months, and days of ownership
def calculate_ownership(purchase_date):
    today = datetime.now().date()
    if purchase_date:
        ownership_duration = relativedelta(today, purchase_date)
        years = ownership_duration.years
        months = ownership_duration.months
        days = ownership_duration.days
        # Format the purchase date as MM/DD/YYYY
        formatted_purchase_date = purchase_date.strftime("%m/%d/%Y")
        return f"{years} years, {months} months, {days} days (Purchased on {formatted_purchase_date})"
    return "N/A"

# Function to get the last watering date when "Y" is selected
def get_last_watering(plant_name):
    wb = load_workbook(config.EXCEL_FILE_PATH)  # Use path from config
    last_watering = "N/A"
    
    # Loop through all the sheets (monthly sheets)
    for sheet_name in wb.sheetnames:
        if sheet_name != "Basic info":  # Ignore the "Basic info" sheet
            sheet = wb[sheet_name]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[1] == plant_name and row[2] == "Y":  # Column B is plant name, Column C is Watering (Y/N)
                    last_watering = row[0]  # Date in column A
    
    return last_watering

# Function to get the last photo path for the selected plant
def get_last_photo(plant_name):
    wb = load_workbook(config.EXCEL_FILE_PATH)  # Use path from config
    last_photo = "N/A"
    
    # Loop through all the sheets (monthly sheets)
    for sheet_name in wb.sheetnames:
        if sheet_name != "Basic info":  # Ignore the "Basic info" sheet
            sheet = wb[sheet_name]
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[1] == plant_name and row[6]:  # Column B is plant name, Column G is photo path
                    last_photo = row[6]
    
    return last_photo

# Function to update Excel based on available data and handle optional photo upload
def update_excel(plant_name, watering, humidity, temperature, notes, photo_path, selected_date):
    # Load the existing workbook
    wb = load_workbook(config.EXCEL_FILE_PATH)  # Use path from config
    
    # Get the appropriate sheet for the current month
    current_month = selected_date.strftime("%B_%Y")
    if current_month not in wb.sheetnames:
        wb.create_sheet(current_month)
        new_sheet = wb[current_month]
        new_sheet.append(["Date", "Plant Name", "Watering (Y/N)", "Humidity (%)", "Room Temperature (°F)", "Notes", "Photo Path"])
    
    monthly_sheet = wb[current_month]
    
    # Find the last row in the monthly sheet to append new data
    date_today = selected_date.now().strftime("%m/%d/%Y")
    
    # Handle photo file and create path if photo is provided
    photo_destination = ""
    if photo_path:
        year_folder = os.path.join(config.PLANT_PHOTO_DIR, selected_date.now().strftime("%Y"))
        month_folder = os.path.join(year_folder, selected_date.now().strftime("%m-%B"))
        os.makedirs(month_folder, exist_ok=True)

        # Format photo name
        photo_name = f"{plant_name.lower().replace(' ', '_')}_{selected_date}.jpg"
        photo_destination = os.path.join(month_folder, photo_name)

        # Copy photo to the destination folder
        shutil.copy(photo_path, photo_destination)

    # Append the data to the monthly sheet
    row_data = [selected_date, plant_name, watering if watering else "", 
                humidity if humidity else "", 
                temperature if temperature else "", 
                notes if notes else "", 
                photo_destination]
    
    monthly_sheet.append(row_data)
    
    # Save the workbook
    wb.save(config.EXCEL_FILE_PATH)  # Use path from config

# Function to update the displayed info when a plant is selected
def on_plant_select(event):
    selected_plant_name = plant_name.get()
    
    # Display species and description for the selected plant
    plant_species = plant_data[selected_plant_name]["plant"]
    species_label.config(text=f"Species: {plant_species}")
    # description = plant_data[selected_plant_name]["description"]
    # description_label.config(text=f"Description1: {description}")
    
    # Update description box
    description_text.config(state=tk.NORMAL)
    description_text.delete("1.0", tk.END)
    description_text.insert(tk.END, plant_data[selected_plant_name]["description"])
    description_text.config(state=tk.DISABLED)  # Make it read-only
    
    # Get last watering date
    last_watering = get_last_watering(selected_plant_name)
    last_watering_label.config(text=f"Last Watered: {last_watering}")
    
    # Get last photo path
    last_photo = get_last_photo(selected_plant_name)
    photo_label.config(text=f"Last Photo Path: {last_photo}")
    
    # Calculate and display ownership duration
    purchase_date = plant_data[selected_plant_name]["purchase_date"]
    ownership_duration = calculate_ownership(purchase_date)
    ownership_label.config(text=f"Age Select: {ownership_duration}")

# Function to edit the plant name
def edit_plant_name():
    selected_plant_name = plant_name.get()
    new_plant_name = new_plant_name_entry.get()

    if new_plant_name in plant_data:
        messagebox.showerror("Error", "Plant nickname must be unique.")
        return

    # Load the workbook and update the Basic info sheet
    wb = load_workbook(config.EXCEL_FILE_PATH)  # Use path from config
    sheet = wb['Basic info']
    
    # Update the plant name in Basic info sheet
    for row in sheet.iter_rows(min_row=2, values_only=False):
        if row[1].value == selected_plant_name:  # Update the nickname
            row[1].value = new_plant_name
    
    # Update the plant name in all monthly sheets
    for sheet_name in wb.sheetnames:
        if sheet_name != "Basic info":
            month_sheet = wb[sheet_name]
            for row in month_sheet.iter_rows(min_row=2, values_only=False):
                if row[1].value == selected_plant_name:
                    row[1].value = new_plant_name

    # Save the workbook
    wb.save(config.EXCEL_FILE_PATH)  # Use path from config
    
    # Update the plant_data dictionary
    plant_data[new_plant_name] = plant_data.pop(selected_plant_name)
    
    # Update the dropdown menu for plant names
    plant_name.set("")
    plant_menu['menu'].delete(0, 'end')
    for plant in plant_data.keys():
        plant_menu['menu'].add_command(label=plant, command=tk._setit(plant_name, plant, on_plant_select))
    
    # Clear the new plant name entry
    new_plant_name_entry.delete(0, tk.END)
    messagebox.showinfo("Success", f"Plant name updated to {new_plant_name}")

# Create the main window for data entry
root = tk.Tk()
root.title("PlantDaddy's Plant Data")

# Set window size (width x height)
root.geometry("800x600")

# Add a label to the window
label = tk.Label(root, text="Plant Tracker", font=("Arial", 16))
label.grid(pady=20)

# Row 0: Name Frame
name_frame = tk.Frame(root)
name_frame.grid(row=0, column=0, columnspan=4, pady=10)

plant_label = tk.Label(name_frame, text="Plant Name:")
plant_label.grid(row=0, column=0, sticky="W", padx=10)

# Fetch plant data from Excel
plant_data = get_plant_info()

# Dropdown for plant name (loaded from Excel sheet)
plant_name = tk.StringVar(root)
plant_menu = tk.OptionMenu(name_frame, plant_name, *plant_data.keys(), command=on_plant_select)
plant_menu.grid(row=0, column=1, columnspan=2, padx=10, pady=5)

# Row 1: Detail Frame
detail_frame = tk.Frame(root)
detail_frame.grid(row=1, column=0, columnspan=2, pady=10)

ownership_label = tk.Label(detail_frame, text="Age Before:")
ownership_label.grid(row=0, column=0, sticky="W", padx=10)
ownership_data = tk.Label(detail_frame, text="")
ownership_data.grid(row=0, column=1, sticky="W")

species_label = tk.Label(detail_frame, text="Species:")
species_label.grid(row=1, column=0, sticky="W", padx=10)
species_data = tk.Label(detail_frame, text="")
species_data.grid(row=1, column=1, sticky="W")

last_watering_label = tk.Label(detail_frame, text="Last Watering:")
last_watering_label.grid(row=2, column=0, sticky="W", padx=10)
last_watering_data = tk.Label(detail_frame, text="")
last_watering_data.grid(row=2, column=1, sticky="W")

# Row 2: Input Frame
input_frame = tk.Frame(root)
input_frame.grid(row=2, column=0, columnspan=2, pady=10)

# Input for humidity (optional)
humidity_label = tk.Label(input_frame, text="Humidity (%)")
humidity_label.grid(row=0, column=0, sticky="W", padx=10)
humidity_entry = tk.Entry(input_frame)
humidity_entry.grid(row=0, column=1, padx=10, pady=5)

# Input for room temperature (optional)
temperature_label = tk.Label(input_frame, text="Temperature (°F)")
temperature_label.grid(row=1, column=0, sticky="W", padx=10)
temperature_entry = tk.Entry(input_frame)
temperature_entry.grid(row=1, column=1, padx=10, pady=5)

# Input for notes (optional)
notes_label = tk.Label(input_frame, text="Notes:")
notes_label.grid(row=2, column=0, sticky="W", padx=10)
# Notes Entry Box for multiline input
notes_entry = tk.Text(input_frame, height=4, width=30)
notes_entry.grid(row=2, column=1, padx=10, pady=5)

watering_label = tk.Label(input_frame, text="Watering (Y/N):")
watering_label.grid(row=3, column=0, sticky="W", padx=10)
watering_dropdown = tk.StringVar(input_frame)
watering_menu = tk.OptionMenu(input_frame, watering_dropdown, "Y", "N")
watering_menu.grid(row=3, column=1, padx=10, pady=5)
watering_dropdown.set("N")  # Default to "No"

# Row 1 & 2: Photo Frame for selecting and displaying a photo
photo_frame = tk.Frame(root)
photo_frame.grid(row=1, column=2, rowspan=2, columnspan=2, pady=10, padx=10)
photo_label = tk.Label(photo_frame, text="Photo Here", relief="solid", width=20, height=10)
photo_label.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

description_label = tk.Label(photo_frame, text="Description2:")
description_label.grid(row=1, column=0, sticky="W", padx=10, pady=5)
# Scrollable Description Box (Read-only, will scroll if text is too long)
description_text = tk.Text(photo_frame, wrap="word", height=4, width=30, state=tk.DISABLED)
description_text.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
# Add a scrollbar for the description box
scrollbar = tk.Scrollbar(photo_frame, command=description_text.yview)
description_text.config(yscrollcommand=scrollbar.set)
scrollbar.grid(row=2, column=2, sticky="ns")

# Date Picker (Date Selector)
date_label = tk.Label(root, text="Select Date:")
date_label.grid(row=4, column=0, pady=10)

cal = Calendar(root, selectmode="day", date_pattern="mm/dd/yyyy")
cal.grid(row=4, column=1, pady=10)

# Row 3: Submit Button (Centered across all columns)
# submit_button = tk.Button(root, text="Submit", width=20, command=lambda: submit(plant_name.get(), watering_dropdown.get()))
submit_button = tk.Button(root, text="Submit", width=20, command=lambda: submit())
submit_button.grid(row=3, column=0, columnspan=4, pady=20)

# Button to upload a photo (optional)
photo_path = ""
def upload_photo():
    global photo_path
    photo_path = filedialog.askopenfilename(title="Select Plant Photo", filetypes=[("Image files", "*.jpg *.jpeg *.png")])
    if photo_path:
        photo_label.config(text=os.path.basename(photo_path))

photo_button = tk.Button(root, text="Upload Photo (Optional)", command=upload_photo)
photo_button.grid(row=9, column=0)

photo_label = tk.Label(root, text="No photo selected")
photo_label.grid(row=9, column=1)

# Submit button function to gather inputs and update Excel
def submit():
    selected_date = cal.get_date()
    selected_plant_name = plant_name.get()
    watering = watering_dropdown.get() if watering_dropdown.get() else None
    humidity = humidity_entry.get() if humidity_entry.get() else None
    temperature = temperature_entry.get() if temperature_entry.get() else None
    notes = notes_entry.get("1.0", tk.END).strip() if notes_entry.get("1.0", tk.END).strip() else None

    print(f"Submitted: Plant={plant_name}, Watering={watering}")

    # Call the update_excel function to append the data
    update_excel(selected_plant_name, watering, humidity, temperature, notes, photo_path, selected_date)

    # Clear the inputs after submission
    cal.set("")
    plant_name.set("")
    watering_dropdown.set("")
    humidity_entry.delete(0, tk.END)
    temperature_entry.delete(0, tk.END)
    notes_entry.delete(0, tk.END)
    photo_label.config(text="No photo selected")

# Run Main application
root.mainloop()