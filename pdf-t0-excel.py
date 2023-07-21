import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime


def convert_pdf_tables_to_excel(pdf_filepath):
    all_tables = pd.DataFrame()
    with pdfplumber.open(pdf_filepath) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                df = pd.DataFrame(table)
                all_tables = all_tables.append(df, ignore_index=True)

    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    excel_filepath = f'output-{current_time}.xlsx'

    # Write the DataFrame to an Excel file
    all_tables.to_excel(excel_filepath, index=False)

    print(f"Excel file was saved under: {excel_filepath}")


# Create the tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the main window

# Open the file dialog and get the path of the selected PDF file
pdf_filepath = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])

# Call the function to convert the PDF to Excel
convert_pdf_tables_to_excel(pdf_filepath)

# Close the tkinter root window
root.destroy()
