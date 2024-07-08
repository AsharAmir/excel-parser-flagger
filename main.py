import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

# Load the Excel file
filepath = 'client.xlsx'
df = pd.read_excel(filepath)

# Create a new Excel workbook to save the flagged results
wb = openpyxl.load_workbook(filepath)
ws = wb.active

# Group by PR Award Number
grouped = df.groupby('PR Award Number')

# Define a list of fill colors for highlighting
colors = [
    "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
    "FF800000", "FF808000", "FF008000", "FF800080", "FF008080", "FF000080"
]
fills = [PatternFill(start_color=color, end_color=color, fill_type="solid") for color in colors]

# Dictionary to keep track of colors used for each PR Award Number
color_map = {}

# Function to compare rows and highlight discrepancies
def highlight_discrepancies(group, color_fill):
    columns = group.columns
    for col in columns:
        if col == 'PR Award Number':
            continue
        values = group[col]
        first_value = values.iloc[0]
        if any(values != first_value):
            for idx in group.index:
                current_value = values.loc[idx]
                if pd.isna(current_value) or str(current_value).lower() == 'comments':
                    continue
                cell = ws.cell(row=idx+2, column=columns.get_loc(col)+1)  # +2 to adjust for header and 0-indexing
                cell.fill = color_fill

# Apply the function to each group
color_idx = 0
for name, group in grouped:
    color_fill = fills[color_idx % len(fills)]
    color_map[name] = colors[color_idx % len(colors)]
    highlight_discrepancies(group, color_fill)
    color_idx += 1

# Add a legend at the bottom of the sheet
start_row = ws.max_row + 2
ws.cell(row=start_row, column=1).value = "PR Award Number"
ws.cell(row=start_row, column=2).value = "Color Code"
for i, (award_number, color) in enumerate(color_map.items()):
    ws.cell(row=start_row + i + 1, column=1).value = award_number
    ws.cell(row=start_row + i + 1, column=2).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# Save the workbook with the highlighted discrepancies
output_file_path = 'flagged_colorcoded.xlsx'
wb.save(output_file_path)

print("Discrepancies highlighted and saved to", output_file_path)
