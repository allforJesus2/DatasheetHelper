import openpyxl
import xlwings as xw
import json
from tkinter import messagebox, Toplevel, Listbox, Button, Frame, MULTIPLE, Checkbutton, IntVar, Label, Scrollbar, Canvas, Entry, Spinbox, END, BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, W, E, NW, SE, EW
from math import log10, floor
import tkinter as tk

# Standalone Functions

# region Data Transformation

def transform_dictionary(input_dict, transformation_code):
    """
    Transform all keys in a dictionary using the provided transformation code.
    Returns a new dictionary with transformed keys and original values.
    """
    return {
        translate(key, transformation_code): value
        for key, value in input_dict.items()
    }

def try_round_to_sigfigs(num, sig_figs=4, tolerance=1e-2):
    # Rounds a number to specified significant figures, with additional rounding for very close values.
    org_num = num
    try:
        num = float(str(num).strip().replace(',', ''))
    except:
        return org_num

    if num == 0:
        return 0

    # Get the order of magnitude
    magnitude = floor(log10(abs(num)))

    # Calculate the scaling factor
    scale = 10 ** (sig_figs - 1 - magnitude)

    # Round the scaled number
    rounded = round(num * scale) / scale

    # Additional rounding for very close values
    if abs(rounded - round(rounded)) < tolerance:
        rounded = round(rounded)

    return rounded

def translate(input_string, transformation_code):
    try:
        x_fn = eval(f'lambda x: {transformation_code}')
        transformed_string = x_fn(input_string)
        print(f'doing transformation {transformation_code} on {input_string} -> {transformed_string}')
        return transformed_string
    except Exception as e:
        print(f'error applying transformation: {e}')
        return input_string
# endregion

# region Excel Cell Operations

def find_cells(sheet, search_terms):
    """
    Find rows containing all specified search terms and return the header cells.
    Returns cells found and a dict of primary cells to their related adjacent cells.
    """
    cells = []
    concat_cells = {}

    # Scan each row to find one containing all search terms
    for row_idx, row in enumerate(sheet.iter_rows(), 1):
        found_terms = {}
        
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell_value = cell.value.lower()
                # Check if this cell contains any search term
                for term in search_terms:
                    if term.lower() == cell_value:
                        found_terms[term.lower()] = cell
        
        # If we found all search terms in this row
        if len(found_terms) == len(search_terms):
            # Add the first term's cell as the primary cell
            primary_cell = found_terms[search_terms[0].lower()]
            cells.append(primary_cell)
            
            # Store relationships to other header cells
            related_cells = [found_terms[term.lower()] for term in search_terms[1:]]
            if related_cells:
                concat_cells[primary_cell] = related_cells
                
    return cells, concat_cells

def find_cells_old(sheet, search_terms):
    cells = []
    # Find the row in the sheet that contains any of the specified search terms
    for term in search_terms:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and term.lower() == str(cell.value).lower():
                    cells.append(cell)
    return cells

def get_value_above(sheet, row, col):
    """Get value from merged cell above or regular cell."""
    current_row = row
    while current_row > 1:
        current_row -= 1
        for range_string in sheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_string.bounds
            if (min_row <= current_row <= max_row and
                min_col <= col <= max_col):
                print('mereged cell ', sheet.cell(min_row, min_col).value)
                return sheet.cell(min_row, min_col).value
        cell_value = sheet.cell(row=current_row, column=col).value
        if cell_value is not None:
            return cell_value
    return None

def increment_cell_reference(cell_ref, increment):
    col = ''.join(filter(str.isalpha, cell_ref))
    row = int(''.join(filter(str.isdigit, cell_ref)))
    return f"{col}{row + increment}"
# endregion

# region Table Processing

def table_to_list(sheet, start_row, start_col, end_row, end_col, column_headers=False):
    matrix = []
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row,
                             min_col=start_col, max_col=end_col,
                             values_only=False):
        row_values = [cell.value for cell in row]
        matrix.append(row_values)

    headers = matrix.pop(0)
    modified_headers = headers.copy()

    # Handle duplicates considering merged cells
    header_counts = {}
    for i, header in enumerate(headers):
        if header is None:
            continue
        header_counts[header] = header_counts.get(header, 0) + 1

    for i, header in enumerate(headers):
        print('header ',header)

        # Check if header is not None and if it is duplicated
        if header is not None and header_counts[header] > 1:
            col_index = start_col + i
            #print(sheet.cell(start_row-1, col_index).value)
            above_value = get_value_above(sheet, start_row, col_index)
            if above_value:
                modified_headers[i] = f"{above_value}_{header}"

    table_list = []
    for row in matrix:
        row_dict = dict(zip(modified_headers, row))
        table_list.append(row_dict)

    return table_list

def list_to_tag_dict(table_list, headers, delimiter="-"):
    """
    Convert a list of dictionaries to a dictionary keyed by concatenated header values.
    
    Args:
        table_list: List of dictionaries containing row data
        headers: List of header names to use for key generation
        delimiter: String to use between concatenated values (default="-")
        
    Returns:
        Dictionary with concatenated keys from header values
    """
    tag_dict = {}
    duplicate_counts = {}
    
    for row in table_list:
        # Check if all headers exist in the row
        if all(header in row for header in headers):
            # Concatenate the values for all headers with the delimiter
            key_parts = [str(row[header] or '') for header in headers]
            base_tag_key = delimiter.join(key_parts)
            
            if base_tag_key in tag_dict:
                if base_tag_key not in duplicate_counts:
                    duplicate_counts[base_tag_key] = 1
                duplicate_counts[base_tag_key] += 1
                tag_key = f"{base_tag_key}_{duplicate_counts[base_tag_key]}"
            else:
                tag_key = base_tag_key
                
            tag_dict[tag_key] = row
    
    return tag_dict

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

def get_table_cols(sheet, header_cell, max_empty_allowed=2):
    """Find the column span of a header row, tolerating blank/merged cells.
    
    Args:
        sheet: The openpyxl worksheet object.
        header_cell: The openpyxl cell object representing one cell in the header row.
        max_empty_allowed: The maximum number of consecutive empty/None cells to tolerate
                           before considering it the end of the table header row.
                           Defaults to 2.
    Returns:
        Tuple[int, int]: The start column index and end column index (1-based).
    """
    start_col = header_cell.column
    end_value = header_cell.value # This might be used as an explicit delimiter
    row = header_cell.row
    # max_empty_allowed = 2 # Now passed as an argument

    # --- Find the leftmost column ---
    left_col = start_col
    first_content_col = start_col
    empty_streak_left = 0

    # Scan left from start_col - 1 down to 1
    for col in range(start_col - 1, 0, -1):
        cell_value = sheet.cell(row=row, column=col).value
        is_merged = False
        effective_value = cell_value # Value to check for content

        # Check if the cell is part of a merged range
        for range_string in sheet.merged_cells.ranges:
            min_col_m, min_row_m, max_col_m, max_row_m = range_string.bounds
            if min_row_m <= row <= max_row_m and min_col_m <= col <= max_col_m:
                is_merged = True
                # Use the value from the top-left cell of the merged range
                effective_value = sheet.cell(row=min_row_m, column=min_col_m).value
                # If this merged cell has content, update the first known content column
                if effective_value is not None:
                    first_content_col = min(first_content_col, min_col_m)
                break # Found the merge range

        # Determine if the column effectively has content
        has_content = (effective_value is not None)

        # Update based on content
        if has_content:
            # If merged, first_content_col was already updated. If not merged, update here.
            if not is_merged:
                 first_content_col = col
            empty_streak_left = 0 # Reset streak
        else: # Cell is effectively empty
            empty_streak_left += 1
            if empty_streak_left >= max_empty_allowed:
                # Found the left boundary after a streak of empty cells
                left_col = first_content_col
                break # Exit loop

    # If loop finished without break (reached column 1)
    else:
        left_col = first_content_col


    # --- Find the rightmost column ---
    right_col = start_col
    last_content_col = start_col
    empty_streak_right = 0

    # Scan right from start_col + 1 up to max_column + 1 (to handle edge cases)
    for col in range(start_col + 1, sheet.max_column + 2):
        cell_value = None
        is_merged = False
        effective_value = None

        if col <= sheet.max_column:
            cell_value = sheet.cell(row=row, column=col).value
            effective_value = cell_value
            # Check if the cell is part of a merged range
            for range_string in sheet.merged_cells.ranges:
                min_col_m, min_row_m, max_col_m, max_row_m = range_string.bounds
                if min_row_m <= row <= max_row_m and min_col_m <= col <= max_col_m:
                    is_merged = True
                    effective_value = sheet.cell(row=min_row_m, column=min_col_m).value
                    # If this merged cell has content, update the last known content column to its right boundary
                    if effective_value is not None:
                        last_content_col = max(last_content_col, max_col_m)
                    break # Found the merge range
        else:
            # Treat columns beyond max_column as empty
            effective_value = None


        # Check for explicit end_value delimiter *before* checking for None
        # Use direct cell_value here, not effective_value, as end_value shouldn't trigger on merged cells unless it's the top-left
        if cell_value == end_value:
             right_col = last_content_col # Table ends *before* the delimiter
             break

        # Determine if the column effectively has content
        has_content = (effective_value is not None)

        # Update based on content
        if has_content:
            # If merged, last_content_col was already updated. If not merged, update here.
            if not is_merged:
                 last_content_col = col
            empty_streak_right = 0 # Reset streak
        else: # Cell is effectively empty
            empty_streak_right += 1
            if empty_streak_right >= max_empty_allowed:
                # Found the right boundary after a streak of empty cells
                right_col = last_content_col
                break # Exit loop

    # If loop finished without break (reached end of sheet)
    else:
        right_col = last_content_col

    return left_col, right_col

def process_sheet(sheet, headers, max_empty_allowed=2):
    """
    Process a sheet to extract tables based on headers.
    
    Args:
        sheet: Excel worksheet to process
        headers: List of headers to use for table identification and key generation
        max_empty_allowed: Max consecutive empty cells allowed in header detection.
        
    Returns:
        List of dictionaries with table data
    """
    header_cells, concat_pairs = find_cells(sheet, headers)
    tables = []

    for cell in header_cells:
        table_length = get_table_length_rows(sheet, cell)
        left_col, right_col = get_table_cols(sheet, cell, max_empty_allowed)

        table_dict_list = table_to_list(sheet, cell.row, left_col, 
                                        cell.row + table_length, right_col)

        # Use all headers for tag dictionary generation
        tag_dict = list_to_tag_dict(table_dict_list, headers)
        tables.append(tag_dict)

    return tables

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
# endregion

# region Datasheet Management

def get_unique_sheet_name(datasheet, ds_prefix, sheet_number):
    """Generate unique sheet name, incrementing suffix if needed"""
    base_name = f"{ds_prefix}{str(sheet_number).zfill(2)}"
    name = base_name
    suffix = 1

    while name in datasheet.sheets:
        name = f"{base_name}_{suffix}"
        suffix += 1

    return name

def apply_character_coloring(sheet, cell_address, old_text, new_text):
    """Apply character-level color coding to show changes using difflib"""
    try:
        import difflib
        
        # Get the cell
        cell = sheet.range(cell_address)
        
        # Check if cell is part of a merged range
        if cell.api.MergeCells:
            # Get the merged range address
            merged_range_address = cell.api.MergeArea.Address
            # Use the top-left cell of the merged range
            cell = sheet.range(merged_range_address.split(':')[0])
        
        # Use difflib to find differences
        matcher = difflib.SequenceMatcher(None, old_text, new_text)
        
        # Clear any existing formatting
        try:
            cell.api.Font.Color = (0, 0, 0)  # Reset to black
        except:
            pass
        
        # Apply character-level formatting
        char_index = 0
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Same text - keep black
                char_index += (i2 - i1)
            elif tag == 'replace':
                # Replaced text - color the new text red
                start_pos = char_index
                end_pos = char_index + (j2 - j1)
                
                try:
                    # Color the new text red
                    if end_pos > start_pos:
                        cell.characters[start_pos:end_pos].font.color = (255, 0, 0)  # Red
                except Exception as e:
                    print(f"Error coloring replaced text: {e}")
                
                char_index += (j2 - j1)
            elif tag == 'delete':
                # Deleted text - we can't color it since it's gone
                pass
            elif tag == 'insert':
                # Inserted text - color it red
                start_pos = char_index
                end_pos = char_index + (j2 - j1)
                
                try:
                    # Color the inserted text red
                    if end_pos > start_pos:
                        cell.characters[start_pos:end_pos].font.color = (255, 0, 0)  # Red
                except Exception as e:
                    print(f"Error coloring inserted text: {e}")
                
                char_index += (j2 - j1)
        
    except Exception as e:
        print(f"Error applying character coloring: {e}")
        # Fallback: just color the entire cell red if there are changes
        try:
            if old_text != new_text:
                cell.api.Font.Color = (255, 0, 0)  # Red
        except:
            pass

def apply_append_coloring(sheet, cell_address, current_value, new_value):
    """Apply append-style color coding (green for old, red for new)"""
    try:
        # Get the cell
        cell = sheet.range(cell_address)
        
        # Check if cell is part of a merged range
        if cell.api.MergeCells:
            # Get the merged range address
            merged_range_address = cell.api.MergeArea.Address
            # Use the top-left cell of the merged range
            cell = sheet.range(merged_range_address.split(':')[0])
        
        current_value_str = str(current_value).strip()
        
        if current_value != "":
            if new_value is not None:
                cell_values = current_value_str.split('\n')
                end = len(cell_values) - 1
                
                if end >= 0 and cell_values[end] != str(new_value).strip():
                    # Format differently using xlwings
                    new_value_str = f"{current_value_str}\n{str(new_value).strip()}"
                    cell.value = new_value_str
                    
                    # Apply formatting (green for old, red for new)
                    try:
                        last_line_pos = len(current_value_str)
                        # Ensure the cell has content before trying to format characters
                        if cell.value and len(str(cell.value)) > 0:
                            # Use Font.ColorIndex as a safer alternative
                            cell.characters[:last_line_pos].font.color = (0, 170, 0)  # Green (RGB)
                            cell.characters[last_line_pos:].font.color = (255, 0, 0)  # Red (RGB)
                    except Exception as e:
                        print(f"Warning: Could not format text colors: {e}")
                        # Fallback: try to set the entire cell to red if formatting fails
                        try:
                            cell.api.Font.Color = (255, 0, 0)  # Red (RGB)
                        except Exception as fallback_e:
                            print(f"Warning: Could not apply fallback color formatting: {fallback_e}")
        else:
            if new_value is not None:
                cell.value = new_value
                
    except Exception as e:
        print(f"Error applying append coloring: {e}")
        # Fallback: just set the value without formatting
        try:
            cell.value = new_value
        except:
            pass

def update_cell_xlwings(sheet, cell_address, value, cell_update_option=None):
    """Update cell using xlwings with similar functionality.
    This function works with merged cells too. Skips update if values are identical.
    
    Args:
        sheet: The worksheet to update
        cell_address: The cell address to update
        value: The new value to set
        cell_update_option: Color coding option - None (black), "new_red_old_green" (current method), or "new_red" (character-level)
    """
    
    try:
        # Convert value to string for comparison if it's not None
        value_str = str(value).strip() if value is not None else None
        
        # Get the cell
        cell = sheet.range(cell_address)
        
        # Check if cell is part of a merged range
        if cell.api.MergeCells:
            # Get the merged range address
            merged_range_address = cell.api.MergeArea.Address
            # Use the top-left cell of the merged range
            cell = sheet.range(merged_range_address.split(':')[0])
        
        # Process cell update with colors
        current_value = cell.value if cell.value is not None else ""
        
        # Convert current_value to string for comparison if it's not None
        current_value_str = str(current_value).strip()
        
        # Skip update if values are identical
        if current_value_str == value_str:
            print(f"Skipping update for {cell_address}: values are identical or there is no source value ({value_str})")
            return True
        
        # Handle different color coding options
        if cell_update_option == "new_red":
            # Character-level color coding - replace entire content and color changes red
            cell.value = value_str
            apply_character_coloring(sheet, cell_address, current_value_str, value_str)
            
        elif cell_update_option == "new_red_old_green":
            # Current method - append with new line and color old green, new red
            apply_append_coloring(sheet, cell_address, current_value, value)
            
        else:
            # No color coding - just update the value in black
            cell.value = value_str

                
        return True
    except Exception as e:
        print(f'Error updating cell {cell_address}: {e}')
        return False

def apply_green_highlighting(sheet, cell_address):
    """Apply green background highlighting to a cell"""
    try:
        # Get the cell
        cell = sheet.range(cell_address)
        
        # Check if cell is part of a merged range
        if cell.api.MergeCells:
            # Get the merged range address
            merged_range_address = cell.api.MergeArea.Address
            # Use the top-left cell of the merged range
            cell = sheet.range(merged_range_address.split(':')[0])
        
        # Apply green background color
        cell.color = (0, 255, 0)  # Green (RGB)
        
    except Exception as e:
        print(f"Error applying green highlighting to {cell_address}: {e}")


def add_update_datasheets(datasheet, source_sheet_name, tag_cell_values, datasheet_coord, ds_prefix,
                   rows_per_sheet=1, custom_sort=None, key_coordinate='I12',
                   sig_figs=4, tolerance=1e-2, halt_callback=None, cell_update_option=None, partial_match=False):
    """
    Manages Excel sheets by adding or updating data based on tags.
    
    Updates existing tags first, then creates new sheets for additional tags if a source sheet is provided.
    
    Args:
        halt_callback: Optional function that returns True if the process should be halted
        cell_update_option: Color coding option - None (black), "new_red_old_green" (current method), or "new_red" (character-level)
        partial_match: If True, matches tags if found anywhere in cell text; if False, requires exact match
    """
    # Determine if we can create new sheets
    can_create_new_sheets = False
    if source_sheet_name:
        try:
            source_sheet = datasheet.sheets[source_sheet_name]
            can_create_new_sheets = True
        except Exception as e:
            print(f"Error accessing source sheet: {e}")
            source_sheet = None
    
    # Find all existing tags
    existing_tags = []
    for sheet in datasheet.sheets:
        if sheet.name == source_sheet_name:
            continue
            
        if sheet.name.startswith(ds_prefix) or ds_prefix == '' :
            print(f"Processing sheet {sheet.name} since it starts with {ds_prefix}")
            print('key_coordinate', key_coordinate)
            for i in range(rows_per_sheet):
                offset_coord = increment_cell_reference(key_coordinate, i)
                tag_value = sheet.range(offset_coord).value
                if tag_value:
                    existing_tags.append((tag_value, sheet.name, offset_coord))
    
    print(f'Existing tags: {existing_tags}')
    
    # Helper function to check if a tag matches a cell value
    def matches_tag(cell_value, tag, partial=False):
        """Returns True if cell_value matches the tag based on partial_match setting"""
        if cell_value is None:
            return False
        if partial:
            # Partial match: check if tag is found in cell value (case-insensitive)
            return str(tag).lower() in str(cell_value).lower()
        else:
            # Exact match
            return cell_value == tag
    
    # Sort tags according to custom function or default alphabetical
    sorted_keys = sorted(tag_cell_values, key=custom_sort) if custom_sort else tag_cell_values
    
    added_sheets = set()
    # Process existing tags first
    for tag in sorted_keys:
        # Check if process should be halted
        if halt_callback and halt_callback():
            print("Process halted by user")
            return list(added_sheets)
            
        # Find all instances of this tag in existing_tags list
        tag_instances = []
        for tag_value, sheet_name, tag_coord in existing_tags:
            if matches_tag(tag_value, tag, partial_match):
                tag_instances.append((sheet_name, tag_coord))
        
        if tag_instances:
            # Update values for all instances of existing tag
            print(f'Updating existing tag {tag} found in {len(tag_instances)} instance(s)')
            for sheet_name, tag_coord in tag_instances:
                target_sheet = datasheet.sheets[sheet_name]
                
                # Update cells for this tag
                cell_values = tag_cell_values[tag]
                try:
                    row_offset = int(tag_coord[1:]) - int(key_coordinate[1:])
                except (ValueError, IndexError) as e:
                    print(f"Error calculating row offset for existing tag {tag}: {e}")
                    print(f"tag_coord: '{tag_coord}', key_coordinate: '{key_coordinate}'")
                    row_offset = 0  # Default to 0 if we can't calculate the offset
                
                for cell, value in cell_values.items():
                    try:
                        target_cell = increment_cell_reference(cell, row_offset)
                        value = try_round_to_sigfigs(value, sig_figs, tolerance)
                        update_cell_xlwings(target_sheet, target_cell, value, cell_update_option)
                    except Exception as e:
                        print(f"Error updating cell {cell} for tag {tag}: {e}")
    
    # Check for unmatched tags in existing sheets and highlight them in green
    # For partial match mode, we need to check if any existing tag was matched
    matched_existing_tags = set()
    for tag_value, _, _ in existing_tags:
        for tag in sorted_keys:
            if matches_tag(tag_value, tag, partial_match):
                matched_existing_tags.add(tag_value)
                break
    
    unmatched_tag_instances = []
    for tag_value, sheet_name, tag_coord in existing_tags:
        if tag_value not in matched_existing_tags:
            unmatched_tag_instances.append((tag_value, sheet_name, tag_coord))
    
    if unmatched_tag_instances:
        print(f"Found {len(unmatched_tag_instances)} unmatched tag instances that will be highlighted in green")
        for tag_value, sheet_name, tag_coord in unmatched_tag_instances:
            target_sheet = datasheet.sheets[sheet_name]
            print(f"Highlighting unmatched tag '{tag_value}' at {tag_coord} in sheet '{sheet_name}'")
            apply_green_highlighting(target_sheet, tag_coord)
    
    # Create new sheets for remaining tags if we have a source sheet
    if can_create_new_sheets:
        count = len(existing_tags)
        print(f"existing tags length: {len(existing_tags)}")
        # Find remaining tags - those that don't match any existing tags
        remaining_tags = []
        for tag in sorted_keys:
            has_match = any(matches_tag(tag_value, tag, partial_match) for tag_value, _, _ in existing_tags)
            if not has_match:
                remaining_tags.append(tag)
        print(f"remaining tags length: {len(remaining_tags)}")

        for tag in remaining_tags:
            # Check if process should be halted
            if halt_callback and halt_callback():
                print("Process halted by user")
                return list(added_sheets)
                
            print(f'Adding new tag {tag}')
            datasheet_no = get_unique_sheet_name(datasheet, ds_prefix, (count // rows_per_sheet) + 1)
            print(f'Datasheet number: {datasheet_no}, count: {count}, rows per sheet: {rows_per_sheet}')
            if rows_per_sheet == 1:
                sheet_name = tag
            else:
                sheet_name = datasheet_no
                
            if count % rows_per_sheet == 0:
                try:
                    datasheet.sheets[sheet_name].delete()  # Remove if exists
                except:
                    pass
                target_sheet = source_sheet.copy(name=sheet_name)
                # Ensure the copied sheet is visible (fix for sheets being hidden)
                target_sheet.visible = True
                added_sheets.add(sheet_name)
                update_cell_xlwings(target_sheet, datasheet_coord, datasheet_no, cell_update_option)
            else:
                target_sheet = datasheet.sheets[sheet_name]
                
            tag_coord = increment_cell_reference(key_coordinate, count % rows_per_sheet)
            update_cell_xlwings(target_sheet, tag_coord, tag, cell_update_option)
            
            # Update cells for this tag
            cell_values = tag_cell_values[tag]
            try:
                row_offset = int(tag_coord[1:]) - int(key_coordinate[1:])
            except (ValueError, IndexError) as e:
                print(f"Error calculating row offset for tag {tag}: {e}")
                print(f"tag_coord: '{tag_coord}', key_coordinate: '{key_coordinate}'")
                row_offset = 0  # Default to 0 if we can't calculate the offset
            
            for cell, value in cell_values.items():
                try:
                    target_cell = increment_cell_reference(cell, row_offset)
                    value = try_round_to_sigfigs(value, sig_figs, tolerance)
                    update_cell_xlwings(target_sheet, target_cell, value, cell_update_option)
                except Exception as e:
                    print(f"Error updating cell {cell} for tag {tag}: {e}")
                    
            count += 1
    
    return list(added_sheets)

# endregion

# region Dictionary and JSON Operations

def select_sheets_dialog(parent, sheet_names, title="Select Sheets to Process"):
    """
    Display a dialog for selecting sheets to process.
    
    Args:
        parent: Parent tkinter window
        sheet_names: List of sheet names to choose from
        title: Dialog title
        
    Returns:
        List of selected sheet names
    """
    dialog = Toplevel(parent)
    dialog.title(title)
    dialog.geometry("400x400")
    dialog.transient(parent)
    dialog.grab_set()
    
    Label(dialog, text="Select sheets to process:").pack(pady=5)
    
    # Create a frame for the listbox and scrollbar
    list_frame = Frame(dialog)
    list_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
    # Add a canvas with scrollbar
    canvas = tk.Canvas(list_frame)
    scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Pack the scrollbar and canvas
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    
    # Create checkbuttons for each sheet
    check_vars = {}
    for sheet in sheet_names:
        var = IntVar(value=0)
        check_vars[sheet] = var
        Checkbutton(scrollable_frame, text=sheet, variable=var).pack(anchor="w", fill="x")
    
    # Buttons frame
    button_frame = Frame(dialog)
    button_frame.pack(fill="x", padx=10, pady=10, side="bottom")
    
    selected_sheets = []
    
    def on_select_all():
        for var in check_vars.values():
            var.set(1)
    
    def on_select_none():
        for var in check_vars.values():
            var.set(0)
    
    def on_ok():
        nonlocal selected_sheets
        selected_sheets = [sheet for sheet, var in check_vars.items() if var.get() == 1]
        dialog.destroy()
    
    def on_cancel():
        dialog.destroy()
    
    Button(button_frame, text="Select All", command=on_select_all).pack(side="left", padx=5)
    Button(button_frame, text="Select None", command=on_select_none).pack(side="left", padx=5)
    Button(button_frame, text="OK", command=on_ok).pack(side="right", padx=5)
    Button(button_frame, text="Cancel", command=on_cancel).pack(side="right", padx=5)
    
    # Wait for the dialog to be closed
    parent.wait_window(dialog)
    
    return selected_sheets

def generate_dictionary_from_xlsx(wb_path, headers, parent=None, selected_sheets=None, max_empty_allowed=2):
    """
    Generate a dictionary from Excel file where keys are concatenated values 
    from multiple headers.
    
    Args:
        wb_path: Path to Excel workbook
        headers: List of headers to use for dictionary generation
        parent: Parent tkinter window for sheet selection dialog
        selected_sheets: List of sheet names to process (if None, will prompt)
        max_empty_allowed: Max consecutive empty cells allowed in header detection.
        
    Returns:
        Dictionary with concatenated header values as keys
    """
    wb = openpyxl.load_workbook(wb_path, read_only=False, data_only=True)
    wb_tag_data = {}
    all_sheet_names = wb.sheetnames
    
    # If no sheets are pre-selected and we have a parent window, show selection dialog
    if selected_sheets is None and parent is not None:
        selected_sheets = select_sheets_dialog(parent, all_sheet_names)
        
        # If user cancels or selects no sheets, use all sheets
        if not selected_sheets:
            selected_sheets = all_sheet_names
    elif selected_sheets is None:
        # If no parent window provided, use all sheets
        selected_sheets = all_sheet_names
    
    # Process only the selected sheets
    for sheet_name in selected_sheets:
        if sheet_name in wb:
            print(f'processing {sheet_name}')
            sheet = wb[sheet_name]
            
            # Pass max_empty_allowed to process_sheet
            tag_tables = process_sheet(sheet, headers, max_empty_allowed)
            sheet_data = combine_tables(tag_tables)
            
            if sheet_data:
                wb_tag_data.update(sheet_data)
        else:
            print(f'Sheet {sheet_name} not found in workbook')

    return wb_tag_data

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
# endregion

# region Analysis and Validation

def analyze_nested_dict_keys(dict_of_dicts):
    if not dict_of_dicts:
        return {
            'consistent': True,
            'inconsistencies': []
        }

    # Get first dictionary's keys as reference
    reference_keys = set(next(iter(dict_of_dicts.values())).keys())

    inconsistencies = []

    # Check each dictionary against the reference
    for key, d in dict_of_dicts.items():
        current_keys = set(d.keys())
        if current_keys != reference_keys:
            missing_keys = reference_keys - current_keys
            extra_keys = current_keys - reference_keys

            inconsistency = {
                'key': key,  # Using the dictionary key instead of index
                'missing_keys': list(missing_keys) if missing_keys else None,
                'extra_keys': list(extra_keys) if extra_keys else None
            }
            inconsistencies.append(inconsistency)

    return {
        'consistent': len(inconsistencies) == 0,
        'reference_keys': list(reference_keys),
        'total_dictionaries': len(dict_of_dicts),
        'inconsistent_count': len(inconsistencies),
        'inconsistencies': inconsistencies
    }

def show_nested_dict_analysis(dict_of_dicts):
    result = analyze_nested_dict_keys(dict_of_dicts)

    # Build the message string
    if result['consistent']:
        message = "All nested dictionaries have the same keys!"
    else:
        message = f"Found {result['inconsistent_count']} inconsistent dictionaries out of {result['total_dictionaries']}\n\n"
        message += f"Reference keys: {', '.join(result['reference_keys'])}\n\n"
        message += "Inconsistencies:\n"

        for item in result['inconsistencies']:
            message += f"\nDictionary with key '{item['key']}':\n"
            if item['missing_keys']:
                message += f"  Missing keys: {', '.join(item['missing_keys'])}\n"
            if item['extra_keys']:
                message += f"  Extra keys: {', '.join(item['extra_keys'])}\n"

    messagebox.showinfo("Nested Dictionary Key Analysis Results", message)
# endregion
