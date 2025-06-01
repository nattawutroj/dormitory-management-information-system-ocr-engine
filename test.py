import openpyxl
import shutil
import os

# Create result directory if it doesn't exist
os.makedirs('result', exist_ok=True)

template_path = 'filestemplate/slip.xlsx'
output_path = 'result/slip_result.xlsx'

# Data to be filled in the template
data = {
    'D6': '1',
    'F6': 'มกราคม',
    'H6': '2568',
    'A9': 'ค่าไฟฟ้า เดือน มกราคม 2568',
    'G9': 5000.00
}

try:
    # Copy template file
    shutil.copy2(template_path, output_path)
    
    # Load workbook
    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active
    
    # Fill in the data
    for cell, value in data.items():
        sheet[cell] = value
    
    # Save and close
    wb.save(output_path)
    print(f"Successfully generated pay report and saved to {output_path}")
    
except FileNotFoundError:
    print(f"Error: Template file not found at {template_path}")
    exit(1)
except Exception as e:
    print(f"Error occurred: {e}")
    # Clean up the output file if it exists
    if os.path.exists(output_path):
        os.remove(output_path)
    exit(1)