import pandas as pd
from docx import Document


def docx_to_excel(docx_path, excel_path):
    # Load the .docx file
    doc = Document(docx_path)
    tables = doc.tables
    print(len(tables))
    
    # Create a Pandas Excel writer using openpyxl as the engine
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Iterate over tables in the document
        for i, table in enumerate(tables):
            # Initialize a list to hold the data of each table
            data = []
            # Iterate over each row in the table
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                data.append(row_data)
            # Convert the table data into a pandas DataFrame
            df = pd.DataFrame(data)
            df = pd.DataFrame(data[1:], columns=data[0])  # Use the first row as column names
            # Write the DataFrame to a sheet in the Excel file
            # Each table goes into a different sheet
            sheet_name = f'Table {i+1}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
# Specify the path to your .docx file and the desired Excel file path
docx_path = 'report/report.docx'
excel_path = 'result/table_report.xlsx'

# Call the function with the paths
docx_to_excel(docx_path, excel_path)

print("Tables extracted and saved to Excel file.")
