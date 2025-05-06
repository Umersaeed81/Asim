# [Umer Saeed](https://www.linkedin.com/in/engumersaeed/)  
**Senior RF Planning & Optimization Engineer**  


📍 **Location:** Dream Gardens, Defence Road, Lahore  
📞 **Mobile:** +92 301 8412180  
✉ **Email:** [umersaeed81@hotmail.com](mailto:umersaeed81@hotmail.com)  

## **Education**  
🎓 **BSc Telecommunications Engineering** – School of Engineering  
🎓 **MS Data Science** – School of Business and Economics  
**University of Management & Technology** 

------------------------------------

## Import required Libarries


```python
import os  # 📁 Used to interact with the operating system (e.g., file paths, directories)
import pandas as pd  # 🐼 Importing pandas for data manipulation and analysis
```

## Input File Path


```python
working_directory = 'D:/Advance_Data_Sets/Asim'  # 📂 Define the target working directory
os.chdir(working_directory)  # 🔄 Change the current working directory to the specified path
```

## Import Input File


```python
df = pd.read_excel('Input.xlsx', parse_dates=['Date'], usecols=['Date','Vehicle Number','Location'])  
# 📄 Read the 'Input.xlsx' file into a DataFrame (df)
# 📅 Parse the 'Date' column as datetime objects
# 🗂️ Only read the 'Date', 'Vehicle Number', and 'Location' columns to save memory and time
```

## Extract Day name


```python
# 🗓️ Extract the day name (e.g., Monday) from the 'Date' column and 🏷️ store it in a new 'Day' column
df['Day'] = df['Date'].dt.day_name()  
```

## Re-Shape Data Set


```python
# 📊 Create a pivot table from the DataFrame
df1 = df.pivot_table(index=['Date','Day'],\
                    columns='Vehicle Number',values='Location',\
                    aggfunc=lambda x: ' '.join(str(v) for v in x))\
                    .reset_index()\
                    .fillna('Idle')\
                    .replace('Blank', 'Idle')

# 📅 Use 'Date' and 'Day' as the row indices
# 🚗 Create separate columns for each vehicle
#📍 Use 'Location' as the values to aggregate
# 🔄 Combine multiple location entries into a single string
# 🔁 Reset the index to turn 'Date' and 'Day' back into columns
# 💤 Replace missing values (NaN) with 'Idle'
# 🔄 Replace any 'Blank' entries with 'Idle'
```

## Export Output


```python
# 📅 Convert 'Date' column to datetime and extract only the date part (drop time)
df1['Date'] = pd.to_datetime(df1['Date']).dt.date  
df1.to_excel('Tracker.xlsx', index=False, sheet_name='Vehicle_Tracking')  
# 💾 Export the DataFrame to an Excel file named 'Tracker.xlsx' 📊, without the index, and with the sheet name 'Vehicle_Tracking'
```


```python
# 🔄⚠️ Forcefully reset the IPython environment by removing all user-defined variables and imports without asking for confirmation
%reset -f
```

## Import required Libarries


```python
import os  # 📁 Provides functions to interact with the operating system (e.g., file paths, directories)
import openpyxl  # 📘 A library for reading and writing Excel (.xlsx) files
from openpyxl import load_workbook  # 📖 Used specifically to load existing Excel workbooks for editing
```

## Set Input File Path


```python
working_directory = 'D:/Advance_Data_Sets/Asim'  # 📂 Define the path to the target working directory
os.chdir(working_directory)  # 🔄 Change the current working directory to the specified folder
```

## Load Excel Sheet


```python
# 📖 Load the existing Excel workbook named 'Tracker.xlsx' for editing
wb = load_workbook('Tracker.xlsx')  
```

## Set Tab Color (All the Tabs)


```python
# 🎨 List of colors for tab colors in hex format
colors = ["00B0F0", "0000FF", "ADD8E6", "87CEFA"]
# Loop through each sheet in the workbook and assign a tab color
for i, ws in enumerate(wb):
    ws.sheet_properties.tabColor = colors[i % len(colors)]        
# 🖍️ Set the tab color of the sheet, cycling through the color list using modulo
```

## Apply border (All the Sheets)


```python
# Define a thin border style for all sides (left, right, top, bottom)
border = openpyxl.styles.borders.Border(
    left=openpyxl.styles.borders.Side(style='thin'),  # 🖊️ Thin border on the left side
    right=openpyxl.styles.borders.Side(style='thin'),  # 🖊️ Thin border on the right side
    top=openpyxl.styles.borders.Side(style='thin'),  # 🖊️ Thin border on the top side
    bottom=openpyxl.styles.borders.Side(style='thin')  # 🖊️ Thin border on the bottom side
)
```

## Font, Alignment and Border (All the Sheets)


```python
from openpyxl import load_workbook  # 📖 Load the existing Excel workbook
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side  # ✨ Import styling components for the Excel sheet

# Define named styles for font, alignment, and border
style = NamedStyle(name="styled_cell")  # 🏷️ Create a named style called 'styled_cell'

# Set font properties: Calibri Light, size 11
style.font = Font(name='Calibri Light', size=11)  # 🖋️ Apply font style

# Set cell alignment to center both horizontally and vertically
style.alignment = Alignment(horizontal='center', vertical='center')  # 🔄 Align text in the center

# Define a thin border for all sides (left, right, top, bottom)
style.border = Border(left=Side(style='thin'),
                      right=Side(style='thin'),
                      top=Side(style='thin'),
                      bottom=Side(style='thin'))  # 🖊️ Apply a thin border to the cell

# Register the named style with the workbook so it can be applied to cells
wb.add_named_style(style)  # 📋 Add the 'styled_cell' style to the workbook

# Disable auto calculation to prevent automatic recalculations during editing
wb.calculation.calcMode = 'manual'  # ⏸️ Turn off auto calculations for performance

# Iterate through each worksheet in the workbook
for ws in wb:  # 🔄 Loop through each sheet
    # Apply the 'styled_cell' named style to all cells in the sheet
    for row in ws.iter_rows():  # 📋 Loop through all rows
        for cell in row:  # 🏷️ Loop through each cell in the row
            cell.style = "styled_cell"  # ✨ Apply the custom style to the cell

# Re-enable auto calculation after making changes
wb.calculation.calcMode = 'auto'  # 🔄 Turn auto calculations back on
```

## Set Date Format


```python
font_style = Font(name='Calibri Light', size=11)  # 🖋️ Define font style with 'Calibri Light' and size 11
alignment_style = Alignment(horizontal='center', vertical='center')  # 🔄 Define cell alignment (centered horizontally and vertically)

from datetime import datetime  # 📅 Import datetime for date handling
from openpyxl.styles import NamedStyle  # ✨ Import NamedStyle for custom styles

# Define a custom date style with a specific date format (DD-MM-YYYY)
date_style = NamedStyle(name='custom_date_style', number_format='DD-MM-YYYY')  # 🗓️ Define style for date formatting

# Loop through each sheet in the workbook
for ws in wb:  # 🔄 Iterate through each worksheet in the workbook
    # Loop through each column in the sheet
    for col in ws.iter_cols():  # 📊 Iterate through columns in the sheet
        for cell in col:  # 🏷️ Iterate through each cell in the column
            if isinstance(cell.value, datetime):  # 📅 Check if the cell contains a datetime value
                cell.style = date_style  # 🖋️ Apply the custom date style to the cell
                cell.border = border  # 🖊️ Apply the border to the cell
                cell.font = font_style  # 🖋️ Apply the font style to the cell
                cell.alignment = alignment_style  # 🔄 Apply the alignment style to the cell
```

## WrapText of header and Formatting (All the sheets)


```python
for ws in wb:  # 🔄 Loop through each sheet in the workbook
    for row in ws.iter_rows(min_row=1, max_row=1):  # 📊 Iterate through the first row of the sheet
        for cell in row:  # 🏷️ Iterate through each cell in the row
            # Set alignment properties for the cell
            cell.alignment = openpyxl.styles.Alignment(
                wrapText=False,  # 🚫 Prevent text wrapping
                horizontal='center',  # 🔄 Center text horizontally
                vertical='center',  # 🔄 Center text vertically
                textRotation=0  # 🔄 No text rotation
            )
            
            # Set fill color for the cell (red background)
            cell.fill = openpyxl.styles.PatternFill(
                start_color="C00000",  # 🔴 Set starting color to red
                end_color="C00000",  # 🔴 Set ending color to red (solid fill)
                fill_type="solid"  # 🟩 Fill the cell with a solid color
            )
            
            # Set font properties (white, bold, size 11, 'Calibri Light')
            font = openpyxl.styles.Font(
                color="FFFFFF",  # ⚪ Set font color to white
                bold=True,  # 🔥 Make the font bold
                size=11,  # 🔢 Set font size to 11
                name='Calibri Light'  # 🖋️ Set the font to 'Calibri Light'
            )
            cell.font = font  # Apply the font to the cell
```

## Set Filter on the Header (All the sheet)


```python
from openpyxl.utils import get_column_letter  # 🔠 Import function to convert column numbers to letters

for ws in wb:  # 🔄 Loop through each sheet in the workbook
    # Get the first row (row 1) of the worksheet
    first_row = ws[1]  # 📋 Reference to the first row
    
    # Apply the auto filter on the first row
    # The filter will be applied from column A to the last column in the first row
    ws.auto_filter.ref = f"A1:{get_column_letter(len(first_row))}1"  # 🔍 Set the filter range from A1 to the last column in the first row

```

## Set Zoom Size (All the sheets)


```python
for ws in wb:  # 🔄 Loop through each sheet in the workbook
    ws.sheet_view.zoomScale = 80  # 🔍 Set the zoom level of the sheet to 80% for better visibility    
```

## Set Column Width (All the sheets)


```python
# Iterate over all sheets in the workbook
for ws in wb.worksheets:  # 📄 Loop through each sheet in the workbook
    # Iterate over all columns in the current sheet
    for column in ws.columns:  # 📊 Loop through each column in the sheet
        # Get the current width of the column
        current_width = ws.column_dimensions[column[0].column_letter].width  # 📏 Retrieve the current column width
        
        # Get the maximum width of the cells in the column
        length = max(len(str(cell.value)) for cell in column)  # 🔢 Find the longest value (in terms of character length) in the column
        
        # Set the width of the column to fit the maximum width, if it's greater than the current width
        if length > current_width:  # 🔍 Check if the new length exceeds the current width
            ws.column_dimensions[column[0].column_letter].width = length  # 🖋️ Adjust the column width to fit the longest value
```

## Set Float Number Format (All the Tabs)


```python
for ws in wb:  # 🔄 Loop through each sheet in the workbook
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):  # 📊 Loop through each row in the sheet, from the first column to the last column
        for cell in row:  # 🏷️ Loop through each cell in the row
            if isinstance(cell.value, float):  # 🔢 Check if the cell contains a floating-point number
                cell.number_format = '0.00'  # 🖋️ Set the number format to show two decimal places (e.g., 12.34)
```

## Conditional Formatting


```python
from openpyxl import load_workbook  # 📚 Import load_workbook to work with existing workbooks
from openpyxl.styles import PatternFill  # 🎨 Import PatternFill to apply background colors
from openpyxl.formatting.rule import FormulaRule  # 📏 Import FormulaRule to apply conditional formatting rules
from openpyxl.utils import get_column_letter  # 🔠 Import get_column_letter to convert column indices to letters

# Define the fill color for conditional formatting (light red color)
fill = PatternFill(start_color='ffcccc', end_color='ffcccc', fill_type='solid')  # 🔴 Set a light red background for matching cells

for ws in wb.worksheets:  # 🔄 Loop through each worksheet in the workbook
    # max_col = ws.max_column  # 📊 Retrieve the maximum number of columns in the sheet (commented out)
    # max_row = ws.max_row  # 📅 Retrieve the maximum number of rows in the sheet (commented out)
    
    for col in range(3, ws.max_column + 1):  # 🔢 Loop through columns starting from the 3rd column to the last column
        col_letter = get_column_letter(col)  # 🔠 Convert the column index to a letter (e.g., 1 → 'A', 2 → 'B', etc.)
        
        # Define the formula for conditional formatting (check if "Idle" is in the cell)
        formula = f'ISNUMBER(SEARCH("Idle", {col_letter}1))'  # 🔍 Formula searches for the word "Idle" in the first row of the column
        
        # Define the range for the conditional formatting rule (entire column from row 1 to the last row)
        cell_range = f'{col_letter}1:{col_letter}{ws.max_row}'  # 📏 Define the cell range for conditional formatting
        
        # Create the conditional formatting rule with the formula and background color
        rule = FormulaRule(formula=[formula], fill=fill)  # 🖋️ Apply the rule with the light red fill if the condition is met
        
        # Apply the conditional formatting rule to the specified range
        ws.conditional_formatting.add(cell_range, rule)  # ✅ Apply the rule to the range
```

## Insert a New Sheet (as First Sheet)


```python
ws511 = wb.create_sheet("Title Page", 0)  # 📑 Create a new sheet named "Title Page" at the first position (index 0) in the workbook
```

## Merge Specific Row and Columns


```python
ws511.merge_cells(start_row=12, start_column=5, end_row=18, end_column=24)  
# 🔲 Merge cells from E12 to X18 (5th column, 12th row to 24th column, 18th row)
```

## Fill the Merge Cells


```python
ws511.cell(row=12, column=5).value = 'Car Location and Tracking Platform'  
# 🚗 Set the value of cell E12 to "Car Location and Tracking Platform"
```

## Formatting Tital Page Report


```python
# Access the first row starting from row 3 (row 12 in 0-indexed, hence the 11th row)
first_row1 = list(ws511.rows)[11]  # 📜 Get the 12th row (which corresponds to row 12 in the sheet)

# Iterate through the cells in the first row starting from column E
for cell in first_row1[4:]:  # 🔄 Loop through each cell in the row starting from column E (index 4)
    # Center the text horizontally and vertically in each cell
    cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')  # 🎯 Set alignment to center
    
    # Set a solid fill color for the background (dark red color)
    cell.fill = openpyxl.styles.PatternFill(start_color="CF0A2C", end_color="CF0A2C", fill_type="solid")  # 🟥 Apply dark red fill
    
    # Set the font color to white, make it bold, and set the font size to 65 with 'Calibri Light'
    font = openpyxl.styles.Font(color="FFFFFF", bold=True, size=65, name='Calibri Light')  # ✨ Apply bold white font with size 65
    cell.font = font  # 🖋️ Assign the font style to the cell
```

## Hide the gridlines


```python
ws511.sheet_view.showGridLines = False  # 🚫 Hide gridlines in the "Title Page" sheet for a cleaner look
```

## Hide the headings


```python
ws511.sheet_view.showRowColHeaders = False  # 🚫 Hide row and column headers for a cleaner sheet view
```

## Final Output


```python
wb.save('Tracker.xlsx')  
# 💾 Save the workbook as 'Tracker.xlsx' to preserve all changes made
```


```python
# 🔄⚠️ Forcefully reset the IPython environment by removing all user-defined variables and imports without asking for confirmation
%reset -f
```
