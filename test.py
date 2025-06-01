import xlwings as xw
import shutil
import os

# Create result directory if it doesn't exist
os.makedirs('result', exist_ok=True)

template_path = 'filestemplate/pay_report.xlsx'
output_path = 'result/pay_report_result.xlsx'

# Data to be filled in the template
data = {
    'B4': 'นาย สมชาย สมหญิง',
    'B5': '5555555555555',
    'B6': 'วิทยาศาสตร์ประยุกต์',
    'B7': 'คณะวิทยาศาสตร์',
    'F4': '2568-05-01',
    'F5': '1234567890123',
    'F6': '101',
    'F7': 'หอชาย',
    'C10': 'พฤษภาคม 2568',
    'F10': 4500.34,
}

try:
    # Copy template file
    shutil.copy2(template_path, output_path)
    
    # Initialize Excel application
    app = xw.App(visible=False)
    
    try:
        # Open workbook
        wb = app.books.open(output_path)
        sheet = wb.sheets[0]
        
        # Fill in the data
        for cell, value in data.items():
            sheet.range(cell).value = value
        
        # Save and close
        wb.save()
        wb.close()
        print(f"Successfully generated pay report and saved to {output_path}")
        
    except Exception as e:
        print(f"Error while processing Excel file: {e}")
        # Clean up the output file if it exists
        if os.path.exists(output_path):
            os.remove(output_path)
        raise
    
    finally:
        # Always quit Excel application
        app.quit()

except FileNotFoundError:
    print(f"Error: Template file not found at {template_path}")
    exit(1)
except Exception as e:
    print(f"Error occurred: {e}")
    exit(1)