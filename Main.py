import camelot
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

# Create a root window but keep it hidden
root = tk.Tk()
root.withdraw()

# Ask the user to select the input PDF file
pdf_path = filedialog.askopenfilename(title="Select the DTC PDF file", filetypes=[("PDF files", "*.pdf")])
if not pdf_path:  # If user cancels the file dialog
    print("No PDF file selected. Exiting...")
    exit()

# Ask the user to select the output XLSX file path
xlsx_path = filedialog.asksaveasfilename(title="Save the output as", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
if not xlsx_path:  # If user cancels the file dialog
    print("No output path selected. Exiting...")
    exit()
    
# Function to save data to Excel
def save_to_excel(data, module_name, writer):
    df = pd.DataFrame(data, columns=["DTC Number", "DTC Description"])
    
    # Parse the Description column
    df["DTC Description"] = df["DTC Description"].str.split(".").str[0]
    
    # Add the additional columns with empty data
    additional_cols = ["Concerned", "Responsible", "Comments", "Contacted SW Owner", "Raised On TeamWorks", "In Known DTC List"]
    for col in additional_cols:
        df[col] = ""
    
    df.to_excel(writer, sheet_name=module_name, index=False)

# Read PDF
tables = camelot.read_pdf(pdf_path, pages='all')

def get_acronym(name):
    # Predefined acronyms
    predefined = {
        "Electronic Power Steering Module": "EPS",
        "Telematic Gateway Head Unit" : "TGW",
        "Common Powertrain Control Module" : "ECM",
        "Electronic Diff" : "EDIFF", 
        "HVAC Module" : "HVAC",
        "Drive Unit" : "DRVU",
        "Tyre Pressure Monitor Module" : "TPMS",
        "Solo Distronic Radar" : "Radar",
        "Ambient Lighting Control" : "ALCM",
        "Amplifier" : "AMP",
        "Front HVAC Panel" : "HVAC_F",
        "Park and Surround View System" : "ParkMan",
        "Front Long Range Radar" : "ACC"

    }

    # If the name has a predefined acronym, return that
    if name in predefined:
        return predefined[name]

    # Otherwise, generate an acronym from the first letter of each word
    return ''.join(word[0] for word in name.split())


# Create an Excel writer object
with pd.ExcelWriter(xlsx_path) as writer:
    module_name = None
    data = []

    for table in tables:
        df = table.df

        for index, row in df.iterrows():
            if "Reading DTCs from" in row[0]:
                # If we've encountered a new module and have previous data, save it
                if module_name and data:
                    save_to_excel(data, module_name, writer)
                    data.clear()

                module_name = get_acronym(row[1])  # Use the get_acronym function here
            else:
                # Store the DTC data
                data.append(row.tolist())

    # Handle any leftover data after loop ends
    if module_name and data:
        save_to_excel(data, module_name, writer)
