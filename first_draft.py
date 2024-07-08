import pandas as pd;
import openpyxl
from openpyxl.styles import PatternFill

filepath = "client.xlsx"

df = pd.read_excel(filepath)

# Create a new Excel workbook to save the flagged results
wb = openpyxl.load_workbook(filepath)
ws = wb.active

# Group by PR AWARD NUMBER
grouped = df.groupby('PR Award Number')

# Define the red fill for highlighting
red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

# Function to compare rows and highlight discrepancies
def highlight_discrepancies(group):
    columns = group.columns
    for col in columns:
        if col == 'PR Award Number':
            continue
        if not all(group[col] == group[col].iloc[0]):
            for idx in group.index:
                cell = ws.cell(row=idx+2, column=columns.get_loc(col)+1)  # +2 to adjust for header and 0-indexing
                cell.fill = red_fill

# Apply the function to each group
for name, group in grouped:
    highlight_discrepancies(group)

# Save the workbook with the highlighted discrepancies
output_file_path = 'red_flagged.xlsx'
wb.save(output_file_path)

print("Discrepancies highlighted and saved to", output_file_path)