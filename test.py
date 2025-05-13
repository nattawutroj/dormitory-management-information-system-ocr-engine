from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
template_data = {
    'common': {
        'font': Font(name='TH SarabunPSK', size=14),
        'font10': Font(name='TH SarabunPSK', size=10),
    },
    'base_info': {
        'header_title': 'พฤษภาคม 2568',
    },
    'data_list': {
        'room':[{
            'room_name': '1-101',
            'student_list': [
                {
                    'student_id': '1',
                    'name': 'นายทดสอบ 1',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '2',
                    'name': 'นายทดสอบ 2',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-102',
            'student_list': [
                {
                    'student_id': '3',
                    'name': 'นายทดสอบ 3',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-103',
            'student_list': [
                {
                    'student_id': '4',
                    'name': 'นายทดสอบ 4',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '5',
                    'name': 'นายทดสอบ 5',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '6',
                    'name': 'นายทดสอบ 6',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '7',
                    'name': 'นายทดสอบ 7',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-104',
            'student_list': [
                {   
                    'student_id': '8',
                    'name': 'นายทดสอบ 8',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '9',
                    'name': 'นายทดสอบ 9',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {   
                    'student_id': '10',
                    'name': 'นายทดสอบ 10',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '11',
                    'name': 'นายทดสอบ 11',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-105',
            'student_list': [
                {
                    'student_id': '12',
                    'name': 'นายทดสอบ 5',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '13',
                    'name': 'นายทดสอบ 5',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '14',
                    'name': 'นายทดสอบ 5',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '15',
                    'name': 'นายทดสอบ 5',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-104',
            'student_list': [
                {   
                    'student_id': '8',
                    'name': 'นายทดสอบ 8',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '9',
                    'name': 'นายทดสอบ 9',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {   
                    'student_id': '10',
                    'name': 'นายทดสอบ 10',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '11',
                    'name': 'นายทดสอบ 11',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-104',
            'student_list': [
                {   
                    'student_id': '8',
                    'name': 'นายทดสอบ 8',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '9',
                    'name': 'นายทดสอบ 9',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {   
                    'student_id': '10',
                    'name': 'นายทดสอบ 10',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '11',
                    'name': 'นายทดสอบ 11',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-104',
            'student_list': [
                {   
                    'student_id': '8',
                    'name': 'นายทดสอบ 8',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '9',
                    'name': 'นายทดสอบ 9',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {   
                    'student_id': '10',
                    'name': 'นายทดสอบ 10',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '11',
                    'name': 'นายทดสอบ 11',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        },
        {
            'room_name': '1-104',
            'student_list': [
                {   
                    'student_id': '8',
                    'name': 'นายทดสอบ 8',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '9',
                    'name': 'นายทดสอบ 9',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {   
                    'student_id': '10',
                    'name': 'นายทดสอบ 10',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                },
                {
                    'student_id': '11',
                    'name': 'นายทดสอบ 11',
                    'inital': "IT",
                    'electric_meter_before': 100,
                    'electric_meter_after': 200,
                    'used_unit': 100,
                    'price_per_unit': 6,
                    'total_price': 600,
                    'price_divide_student': 300,
                }
            ]
        }
        ]
    }
}

def copy_cell_style(source_cell, target_cell):
    """Copy cell style from source to target"""
    if source_cell.has_style:
        # Copy font
        if source_cell.font:
            target_cell.font = Font(
                name=source_cell.font.name,
                size=source_cell.font.size,
                bold=source_cell.font.bold,
                italic=source_cell.font.italic,
                vertAlign=source_cell.font.vertAlign,
                underline=source_cell.font.underline,
                strike=source_cell.font.strike,
                color=source_cell.font.color
            )
        
        # Copy fill
        if source_cell.fill:
            target_cell.fill = PatternFill(
                fill_type=source_cell.fill.fill_type,
                start_color=source_cell.fill.start_color,
                end_color=source_cell.fill.end_color
            )
        
        # Copy border
        if source_cell.border:
            target_cell.border = Border(
                left=source_cell.border.left,
                right=source_cell.border.right,
                top=source_cell.border.top,
                bottom=source_cell.border.bottom
            )
        
        # Copy alignment
        if source_cell.alignment:
            target_cell.alignment = Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                text_rotation=source_cell.alignment.text_rotation,
                wrap_text=source_cell.alignment.wrap_text,
                shrink_to_fit=source_cell.alignment.shrink_to_fit,
                indent=source_cell.alignment.indent
            )
        
        # Copy number format
        target_cell.number_format = source_cell.number_format

def copy_border(source_border):
    """Create a new border object from source border"""
    if not source_border:
        return None
    
    return Border(
        left=Side(style=source_border.left.style) if source_border.left else None,
        right=Side(style=source_border.right.style) if source_border.right else None,
        top=Side(style=source_border.top.style) if source_border.top else None,
        bottom=Side(style=source_border.bottom.style) if source_border.bottom else None
    )

def parse_range(range_string):
    """Parse a range string into min_col, min_row, max_col, max_row"""
    start, end = range_string.split(':')
    start_col, start_row = coordinate_from_string(start)
    end_col, end_row = coordinate_from_string(end)
    
    return (
        column_index_from_string(start_col),
        start_row,
        column_index_from_string(end_col),
        end_row
    )

def setup_sheet_template(sheet, template_sheet, start_row=1):
    """Copy template settings to new sheet"""
    # Copy print settings
    sheet.page_setup = template_sheet.page_setup
    sheet.print_options = template_sheet.print_options
    sheet.print_title_rows = template_sheet.print_title_rows
    sheet.print_title_cols = template_sheet.print_title_cols
    sheet.print_area = template_sheet.print_area

    # Store merged cells ranges in a list
    merged_ranges = []
    for merged_cells in template_sheet.merged_cells.ranges:
        min_row = merged_cells.min_row + start_row - 1
        max_row = merged_cells.max_row + start_row - 1
        min_col = merged_cells.min_col
        max_col = merged_cells.max_col
        new_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        merged_ranges.append(new_range)

    # Apply merged cells after collecting all ranges
    for new_range in merged_ranges:
        sheet.merge_cells(new_range)

    # Copy all cells from template with offset
    for row in template_sheet.rows:
        for cell in row:
            if not isinstance(cell, MergedCell):
                new_row = cell.row + start_row - 1
                col_letter = get_column_letter(cell.column)
                target_cell = sheet[f"{col_letter}{new_row}"]
                target_cell.value = cell.value
                copy_cell_style(cell, target_cell)

    # Copy column dimensions
    for col in list(template_sheet.column_dimensions.keys()):
        sheet.column_dimensions[col] = template_sheet.column_dimensions[col]

    # Copy row dimensions with offset
    for row in list(template_sheet.row_dimensions.keys()):
        new_row = row + start_row - 1
        sheet.row_dimensions[new_row] = template_sheet.row_dimensions[row]

    # Create thin border style
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Apply borders to all cells in the section
    for row in range(start_row, start_row + 25):
        for col in range(1, 21):  # A to T
            cell = sheet[f"{get_column_letter(col)}{row}"]
            # Copy border from template if it exists
            template_row = row - start_row + 1
            template_cell = template_sheet[f"{get_column_letter(col)}{template_row}"]
            if template_cell.border:
                cell.border = copy_border(template_cell.border)
            else:
                cell.border = thin_border

    # Ensure borders are applied to merged cells
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = parse_range(merged_range)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = sheet[f"{get_column_letter(col)}{row}"]
                if not cell.border:
                    cell.border = thin_border

def fill_room_names(sheet, rooms, start_index, start_row=1):
    """Fill room names in the specified positions on the sheet"""
    positions = ['A5', 'A9', 'A13', 'A17']
    for i, pos in enumerate(positions):
        if start_index + i < len(rooms):
            # Calculate new position with offset
            col = pos[0]
            row = int(pos[1:]) + start_row - 1
            cell = sheet[f"{col}{row}"]
            cell.value = rooms[start_index + i]['room_name']
            cell.font = Font(name='TH SarabunPSK', size=14)

def fill_student_data(sheet, room, start_row):
    """Fill student data for a room group starting at the specified row"""
    students = room['student_list']
    
    # Fill student data for up to 4 students
    for i, student in enumerate(students[:4]):
        row = start_row + i
        
        # Student ID
        sheet[f'B{row}'] = student['student_id']
        
        # Name
        sheet[f'C{row}'] = student['name']
        
        # Initial
        sheet[f'D{row}'] = student['inital']
        
        # Total price and price per student (only for first student)
        if i == 0:
            sheet[f'G{start_row}'] = student['electric_meter_before']
            sheet[f'H{start_row}'] = student['electric_meter_after']
            sheet[f'I{start_row}'] = student['used_unit']
        
        # Total price and price per student for each student
        sheet[f'J{row}'] = student['total_price']
        sheet[f'K{row}'] = student['price_divide_student']

def process_rooms(workbook, rooms):
    """Process rooms and create new sections in the same sheet"""
    rooms_per_section = 4
    total_sections = (len(rooms) + rooms_per_section - 1) // rooms_per_section
    sheet = workbook.active

    # Clear existing content except first section
    for row in range(26, sheet.max_row + 1):
        for col in range(1, 21):  # A to T
            cell = sheet[f"{get_column_letter(col)}{row}"]
            cell.value = None
            cell.border = None
            cell.fill = None
            cell.font = None
            cell.alignment = None

    # Process each section
    for section_num in range(total_sections):
        # Load fresh template for each section
        template_workbook = load_workbook('filestemplate/f2.xlsx')
        template_sheet = template_workbook.active
        
        start_row = section_num * 24 + 1
        setup_sheet_template(sheet, template_sheet, start_row)
        start_index = section_num * rooms_per_section
        
        # Add header title in A1 for each page
        header_cell = sheet[f'A{start_row}']
        header_cell.value = template_data['base_info']['header_title']
        header_cell.font = template_data['common']['font']
        
        # Fill room names and student data
        for i in range(rooms_per_section):
            if start_index + i < len(rooms):
                room = rooms[start_index + i]
                # Calculate the starting row for this room group
                room_start_row = start_row + (i * 4) + 4  # +4 because room names start at A5, A9, etc.
                fill_student_data(sheet, room, room_start_row)
        
        fill_room_names(sheet, rooms, start_index, start_row)
        
        # Close template workbook
        template_workbook.close()

# Load the workbook
try:
    workbook = load_workbook('filestemplate/f2.xlsx')
except FileNotFoundError as e:
    print(f"Error: {e}")
    exit(1)

# Process the rooms from template_data
rooms = template_data['data_list']['room']
process_rooms(workbook, rooms)

# Save the workbook to the result folder
output_path = 'result/f2_result.xlsx'
workbook.save(output_path)
print(f"Successfully generated {len(rooms)} rooms in {((len(rooms) + 3) // 4)} sections and saved to {output_path}")


#  group student by room
#  A5 group
    # student_id B5 - B8 max 4 student
    # name C5 - C8
    # initial D5 - D8
    # electric_meter_before[0] G5
    # electric_meter_after[0] H5
    # used_unit I5
    # total_price J5 - J8
    # price_divide_student K5 - K8
#  A9 group
    # student_id B9 - B12 max 4 student
    # name C9 - C12
    # initial D9 - D12
    # electric_meter_before[0] G9
    # electric_meter_after[0] H9
    # used_unit I9
    # total_price J9 - J12
    # price_divide_student K9 - K12
#  A13 group
    # student_id B13 - B16 max 4 student
    # name C13 - C16
    # initial D13 - D16
    # electric_meter_before[0] G13
    # electric_meter_after[0] H13
    # used_unit I13
    # total_price J13 - J16
    # price_divide_student K13 - K16
#  A17 group
    # student_id B17 - B20 max 4 student
    # name C17 - C20
    # initial D17 - D20
    # electric_meter_before[0] G17
    # electric_meter_after[0] H17
    # used_unit I17
    # total_price J17 - J20
    # price_divide_student K17 - K20
    