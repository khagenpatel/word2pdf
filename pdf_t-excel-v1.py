import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime


def convert_pdf_tables_to_excel(pdf_filepath):
    all_tables = pd.DataFrame()
    with pdfplumber.open(pdf_filepath) as pdf:
        for page_num, page in enumerate(pdf.pages):
            for table_num, table in enumerate(page.extract_tables()):
                df = pd.DataFrame(table)
                print(f'Page {page_num + 1}, Table {table_num + 1}, Columns: {df.columns}')
                print(df.head())
                if all_tables.empty:
                    all_tables = df
                else:
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
