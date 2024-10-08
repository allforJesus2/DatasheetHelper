import openpyxl
from openpyxl.cell.text import InlineFont
#from openpyxl.cell.rich_text import TextBlock, CellRichText
import openpyxl.styles as styles
from openpyxl.styles import Font, Alignment
import os
import xlwings as xw
import re
from datetime import datetime
import math
import json

def find_cells(sheet, search_terms):
    cells = []
    # Find the row in the sheet that contains any of the specified search terms
    for term in search_terms:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and term.lower() == str(cell.value).lower():
                    cells.append(cell)
    return cells

def combine_tables(tables):
    if tables:
        tag_dict_combined = tables[0].copy()
    else:
        print('Empty tables')
        return

    for table in tables[1:]:
        for tag in table:
            if table[tag] is not None:
                # Use setdefault to ensure the tag exists in tag_dict_combined
                try:
                    tag_dict_combined[tag].update(table[tag])
                except Exception as e:
                    print('ERROR: ',e)
            else:
                print(f'Error: No match found for {tag}')
    print(tag_dict_combined)
    return tag_dict_combined

def get_table_length(sheet, header_cell):
    start_row = header_cell.row + 1
    end_value = header_cell.value

    # Initialize table_length to 0
    table_length = 0

    # Iterate down the column starting from the row of the given cell
    current_row = start_row
    current_value = sheet.cell(row=current_row, column=header_cell.column).value


    # Continue iterating until an empty cell is encountered
    max_row = sheet.max_row
    while current_value and current_row <= max_row:
        if current_value == end_value:
            break
        table_length += 1
        current_row += 1
        current_value = sheet.cell(row=current_row, column=header_cell.column).value
        print(current_value)

    return table_length

def add_datasheets(datasheet_path, source_sheet_name, tag_cell_values, datasheet_coord, ds_prefix,
                   tag_pattern=r'^[0-9]{3}-[A-Z]{2,3}-[0-9]{4}[A-Z]?$', rows_per_sheet=1, custom_sort=None):
    if rows_per_sheet > 1:
        count, tag_set = get_count(datasheet_path, tag_pattern, ds_prefix)  # count = 0
    else:
        count, tag_set = get_count(datasheet_path, tag_pattern, '')
    print(f'existing tag set:\n{tag_set}')
    datasheet = xw.Book(datasheet_path)
    source_sheet = datasheet.sheets[source_sheet_name]

    # Get the sorted list of keys from the filtered dictionary
    if custom_sort:
        sorted_keys = sorted(tag_cell_values, key=custom_sort)
    else:
        sorted_keys = sorted(tag_cell_values)

    print(sorted_keys)
    # tag_filters = [{'header':'TAG PFX','filter':'CV'}, {'header':'ACT TYPE', 'filter':'PNEUMATIC'}]
    # tag_filters = [['TAG PFX','CV'], [ACT TYPE', 'PNEUMATIC']]
    for tag in sorted_keys:
        if tag in tag_set:
            print(f'tag {tag} already in datasheet. Continuing...')
            continue

        if rows_per_sheet > 1:
            ds_name = ds_prefix + str((count // rows_per_sheet) + 1).zfill(2)
        else:  # its a one tag sheet
            ds_name = tag

        # need to make a new sheet?
        if count % rows_per_sheet == 0:

            try:
                datasheet.sheets[ds_name].delete()
            except:
                pass
            new_sheet = source_sheet.copy(name=ds_name)
            new_sheet.range(datasheet_coord).value = ds_prefix + str((count // rows_per_sheet) + 1).zfill(2)
        else:
            print(ds_name)
            new_sheet = datasheet.sheets[ds_name]
        # grab the cell values from the dictionary
        cell_values = tag_cell_values[tag]

        # its not going to change anything if rows per sheet is 1 cause that always has a remainder of 0
        incremented_cell_values = increment_cell_values_row(cell_values, count % rows_per_sheet)
        count += 1

        for cell, value in incremented_cell_values.items():
            try:
                new_sheet.range(cell).value = value
                # new_sheet.range(cell).api.Font.Color = 255  # 255 represents red color in Excel
            except Exception as e:
                print(e)

def add_datasheets2(datasheet_path, source_sheet_name, tag_cell_values, datasheet_coord, ds_prefix,
                    rows_per_sheet=1, custom_sort=None, key_coordinate='I12'):

    datasheet = xw.Book(datasheet_path)
    source_sheet = datasheet.sheets[source_sheet_name]

    # Get existing tags and their locations
    existing_tags = {}
    for sheet in datasheet.sheets:
        if sheet.name.startswith(ds_prefix) or ds_prefix == '':
            for i in range(rows_per_sheet):
                offset_coord = increment_cell_reference(key_coordinate, i)
                tag_value = sheet.range(offset_coord).value
                if tag_value:# and re.search(tag_pattern, str(tag_value)):
                    existing_tags[tag_value] = (sheet.name, offset_coord)
                    #we need to update the values here and not later

    print(f'Existing tags: {existing_tags}')

    if custom_sort:
        sorted_keys = sorted(tag_cell_values, key=custom_sort)
    else:
        sorted_keys = sorted(tag_cell_values)

    print(f'Tags to process: {sorted_keys}')

    count = len(existing_tags)
    for tag in sorted_keys:
        if tag in existing_tags:
            print(f'Updating existing tag {tag}')
            sheet_name, tag_coord = existing_tags[tag]
            target_sheet = datasheet.sheets[sheet_name]
        else:
            print(f'Adding new tag {tag}')
            if count % rows_per_sheet == 0:
                sheet_name = f"{ds_prefix}{str((count // rows_per_sheet) + 1).zfill(2)}"
                try:
                    datasheet.sheets[sheet_name].delete()
                except:
                    pass
                target_sheet = source_sheet.copy(name=sheet_name)
                target_sheet.range(datasheet_coord).value = sheet_name
            else:
                sheet_name = f"{ds_prefix}{str((count // rows_per_sheet) + 1).zfill(2)}"
                target_sheet = datasheet.sheets[sheet_name]

            tag_coord = increment_cell_reference(key_coordinate, count % rows_per_sheet)
            count += 1

        cell_values = tag_cell_values[tag]
        row_offset = int(tag_coord[1:]) - int(key_coordinate[1:])

        for cell, value in cell_values.items():
            try:
                target_cell = increment_cell_reference(cell, row_offset)
                target_sheet.range(target_cell).value = value
            except Exception as e:
                print(f"Error updating cell {cell} for tag {tag}: {e}")

    # Remove the workbook from xlwings' internal collection
    #datasheet.app.cleanup()
    return datasheet

def update_datasheets(datasheet_path, tag_cell_values, ds_prefix, rows_per_sheet=1, key_coordinate='I12'):

    datasheet = xw.Book(datasheet_path)

    # Get existing tags and their locations
    # here were baiscally machine-gunning the datasheet with data
    for sheet in datasheet.sheets:
        if sheet.name.startswith(ds_prefix) or ds_prefix == '':
            for i in range(rows_per_sheet):
                tag_coord = increment_cell_reference(key_coordinate, i)
                tag_value = sheet.range(tag_coord).value
                if tag_value in tag_cell_values:# and re.search(tag_pattern, str(tag_value)):
                    coord_values = tag_cell_values[tag_value]
                    # so coord values is going to be like:
                    #{'C38': 'LE-3101', 'E38': 'AT1-2000-PR-PID-111', 'F38': 'TK-3101A', 'H38': 'SOFTENING FEED TANK A'}
                    for coord, value in coord_values.items():
                        current_coord = increment_cell_reference(coord,i)
                        sheet.range(current_coord).value = value

                    #we need to update the values here and not later

def increment_cell_reference(cell_ref, increment):
    col = ''.join(filter(str.isalpha, cell_ref))
    row = int(''.join(filter(str.isdigit, cell_ref)))
    return f"{col}{row + increment}"

def get_count(workbook_path, tag_pattern, ds_prefix):
    # Open the Excel workbook in read-only mode
    workbook = openpyxl.load_workbook(workbook_path, read_only=True)
    sheet_tag_set = set()
    # Iterate through each sheet in the Excel workbook
    for sheet in workbook.sheetnames:
        # Check if the sheet name starts with the specified string
        if sheet.startswith(ds_prefix) or ds_prefix == '':
            # print(sheet)

            # Access the active sheet
            active_sheet = workbook[sheet]

            # Iterate through all cells in the active sheet
            for row in active_sheet.iter_rows(values_only=True):
                for cell_value in row:
                    cell_value = str(cell_value)

                    # Use the re.search() function to find matches with the tag_pattern
                    if re.search(tag_pattern, cell_value):
                        sheet_tag_set.add(cell_value)
                        # print(cell_value)
            # print(len(sheet_tag_set))

    return len(sheet_tag_set), sheet_tag_set

def increment_cell_values_row(cell_values, amount):
    new_cell_values = {}
    for key, value in cell_values.items():

        new_key = increment_coord(key)
        new_cell_values[new_key] = value

    return new_cell_values

def increment_coords_to_fields(coords_to_fields):
    new_coords_to_fields = {}
    for coord, field_name in coords_to_fields.items():

        new_coord = increment_coord(coord)
        new_coords_to_fields[new_coord] = field_name

    return new_coords_to_fields

def increment_coord(coord):
    letter_part, numeric_part = split_text_on_first_number(coord)
    new_coord = letter_part + str(int(numeric_part) + 1)
    return new_coord

def split_text_on_first_number(text):
    match = re.search(r'\d', text)  # Search for the first digit in the text
    if match:
        first_number_index = match.start()  # Get the index of the first digit
        text_before_number = text[:first_number_index]  # Text before the first number
        text_after_number = text[first_number_index:]  # Text from the first number onward
        return text_before_number, text_after_number
    else:
        # If there is no digit in the text, return the entire text
        return text, ''

def list_to_tag_dict(table_list, tag):
    # tag is a string
    tag_dict = {}
    # table_list is a list of dictionaries
    for i, row in enumerate(table_list):
        print(f'processing row {i}', end='\r')
        # row is a dictionary
        for key in row:
            if key == tag:
                tag_key = row[key]
                # tag_key is the line number

        tag_dict[tag_key] = row

    return tag_dict

def generate_dictionary_from_xlsx(wb_path, headers):
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    # process_conditions['09-PM-006-2'] = ...
    wb_tag_data = {}
    sheet_names = wb.sheetnames

    for i, sheet in enumerate(wb.worksheets):
        print(f'processing {sheet_names[i]}')

        tag_tables = process_sheet(sheet, headers)

        sheet_data = combine_tables(tag_tables)

        if sheet_data:
            wb_tag_data.update(sheet_data)

    return wb_tag_data

def process_sheet(sheet, headers):
    header_cells = find_cells(sheet, headers)
    tables = []
    print(header_cells)
    for cell in header_cells:
        print('starting loop')
        # Automatically determine table length based on the number of rows available in the sheet
        print('getting table length')
        table_length = get_table_length_rows(sheet, cell)
        left_col, right_col = get_table_cols(sheet, cell)
        # end = start + table_length
        print('table rows ', table_length)
        print('table cols ', right_col-left_col)
        print('getting table list')
        #we specify the start col and row as the header cell
        table_dict_list = table_to_list(sheet, cell.row, left_col, cell.row + table_length, right_col)
        # table_dict_list is a list of dictionaries (each item corresponding to a row) where the headers are the keys
        print(table_dict_list)
        print('getting tag dict')
        tag_dict = list_to_tag_dict(table_dict_list, cell.value)
        tables.append(tag_dict)

    return tables

def transpose_matrix(matrix):
    # Check if the input is a valid 2D matrix
    if not all(isinstance(row, list) for row in matrix):
        raise ValueError("Input must be a 2D matrix.")

    # Calculate the dimensions of the matrix
    num_rows = len(matrix)
    num_cols = len(matrix[0]) if num_rows else 0

    # Transpose the matrix
    transposed_matrix = [[matrix[j][i] for j in range(num_rows)] for i in range(num_cols)]

    return transposed_matrix

def table_to_list(sheet, start_row, start_col, end_row, end_col, column_headers=False):
    # This function converts a specified range of a spreadsheet sheet into a list of dictionaries,
    # where each dictionary represents a row in the sheet and its keys are the column headers.
    matrix = []
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True):
        print('checking row ', row)
        matrix.append(list(row))

    if column_headers:
        matrix = transpose_matrix(matrix)

    # Extract headers based on the start column
    headers = matrix.pop(0)

    table_list = []

    # Iterate over rows from start_row to end_row
    for row in matrix:
        # Convert each row to a dictionary using the extracted headers
        row_dict = dict(zip(headers, row))
        table_list.append(row_dict)

    return table_list

def get_table_length_rows(sheet, header_cell):
    start_row = header_cell.row + 1
    end_value = header_cell.value
    max_row = sheet.max_row
    current_row = start_row
    # Initialize table_length to 0
    table_length = 0

    # Use iter_rows to iterate over rows starting from the row after the header
    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        current_value = row[header_cell.column - 1]  # Adjusting for zero-based indexing

        # Break the loop if an empty cell is encountered or if the current value matches the end value
        if current_row >= max_row or current_value == end_value:
            break

        current_row += 1
        table_length += 1

    return table_length

def get_table_length_cols(sheet, header_cell):
    start_col = header_cell.column
    end_value = header_cell.value
    max_col = sheet.max_column
    current_col = start_col
    # Initialize table_length to 0
    table_length = 0

    # Iterate over each row in the sheet
    for row in sheet.iter_rows(values_only=True):


        current_col += 1
        table_length += 1

        # Access the cell in the column corresponding to the header_cell
        current_value = row[current_col - 1]  # Adjusting for zero-based indexing

        # Break the loop if an empty cell is encountered or if the current value matches the end value
        if current_col >= max_col or current_value == end_value:
            break

    return table_length

def get_table_cols(sheet, header_cell):
    start_col = header_cell.column
    end_value = header_cell.value
    row = header_cell.row

    # Find the leftmost column
    left_col = start_col
    for col in range(start_col - 1, 0, -1):
        if sheet.cell(row=row, column=col).value is None:
            break
        left_col = col

    # Find the rightmost column
    right_col = start_col
    for col in range(start_col + 1, sheet.max_column + 1):
        cell_value = sheet.cell(row=row, column=col).value
        if cell_value is None or cell_value == end_value:
            right_col = col - 1
            break
        right_col = col

    return left_col, right_col

def load_dict_from_json(file_path):
    """
    Load a dictionary from a JSON file.

    Args:
    file_path (str): The path to the JSON file.

    Returns:
    dict: The dictionary loaded from the JSON file.

    Raises:
    FileNotFoundError: If the specified file is not found.
    json.JSONDecodeError: If the file is not valid JSON.
    """
    try:
        with open(file_path, 'r') as json_file:
            return json.load(json_file)
    except FileNotFoundError:
        raise FileNotFoundError(f"The file {file_path} was not found.")
    except json.JSONDecodeError:
        raise json.JSONDecodeError(f"The file {file_path} is not valid JSON.")