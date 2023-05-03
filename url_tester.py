import urllib.request
import json
import openpyxl

# Open the source workbook
source_wb = openpyxl.load_workbook('source_file.xlsx')
source_ws = source_wb.active

# Open the destination workbook
dest_wb = openpyxl.Workbook()
dest_ws = dest_wb.active

# Paste key here
key = ''
# Paste api endpoint here
url = ''

# Loop through the column of items
for cell in source_ws['A']:
    # Get the value of the current cell
    item = cell.value

    if item is None:
        break
    
    print(item)
    
    # Make the API request
    with urllib.request.urlopen(f'', timeout=None) as response:
        data = json.loads(response.read().decode())

    # Write the data to the destination worksheet
    dest_ws.append([item])
    if 'seg' in data:
        for segment in data['seg']:

            dest_ws.append([' ', segment['name']])
    else:
        dest_ws.append(['No score returned'])

# Save the destination workbook
dest_wb.save('destination_file.xlsx')