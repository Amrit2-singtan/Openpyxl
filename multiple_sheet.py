from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

# Sample data for each sheet
data_sheet1 = [
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles']
]

data_sheet2 = [
    ['Product', 'Price', 'Stock'],
    ['Laptop', 1000, 50],
    ['Phone', 500, 100]
]

# Create a new workbook and remove the default sheet
wb = Workbook()
default_sheet = wb.active
wb.remove(default_sheet)

# Function to create a sheet, write data, and set tab color
def add_sheet(wb, title, data, tab_color):
    ws = wb.create_sheet(title=title)
    for row in data:
        ws.append(row)
    ws.sheet_properties.tabColor = tab_color  # Hex color without #


# Add sheets with different headings and data
add_sheet(wb, 'People', data_sheet1, 'FF9999')      # Light Red tab
add_sheet(wb, 'Inventory', data_sheet2, '99CCFF')   # Light Blue tab

# Save the workbook
wb.save("multi_sheet_excel.xlsx")
