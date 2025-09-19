import tkinter as tk
from tkinter import simpledialog, messagebox
import xlwings as xw


class CoordsToFieldsGenerator:
    def __init__(self, parent, xlsx_path, initial_coords_dict):
        self.parent = parent
        self.xlsx_path = xlsx_path
        self.coords_dict = initial_coords_dict or {}
        self.wb = None
        self.coords_window = None

    def generate(self):
        if not self.xlsx_path:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        try:
            self.wb = xw.Book(self.xlsx_path)
            self.create_window()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file: {str(e)}")

    def create_window(self):
        self.coords_window = tk.Toplevel(self.parent)
        self.coords_window.title("Coordinate to Field Generator")

        frame1 = tk.Frame(self.coords_window)
        frame1.pack(fill=tk.X)

        label = tk.Label(frame1, text="Coordinate")
        label.pack(side=tk.LEFT)

        self.coord_entry = tk.Entry(frame1)
        self.coord_entry.bind('<Return>', self.add_implicit)
        self.coord_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        add_button = tk.Button(frame1, text="Add", command=self.add_implicit)
        add_button.pack(side=tk.LEFT)

        explicit_add_button = tk.Button(frame1, text="Add Explicit", command=self.add_explicit)
        explicit_add_button.pack(side=tk.LEFT)

        remove_button = tk.Button(frame1, text="Remove Entry", command=self.remove_entry)
        remove_button.pack(side=tk.LEFT)

        clear_button = tk.Button(frame1, text="Clear All", command=self.clear_all)
        clear_button.pack(side=tk.LEFT)

        frame2 = tk.Frame(self.coords_window)
        frame2.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(frame2)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        
        # Add a frame for the Done button
        frame3 = tk.Frame(self.coords_window)
        frame3.pack(fill=tk.X, pady=5)
        
        # Add the Done button
        done_button = tk.Button(frame3, text="Done", command=self.on_closing)
        done_button.pack(side=tk.RIGHT, padx=5, pady=5)

        self.update_listbox()

        self.coords_window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.coords_window.after(200, self.update_entry)

    def add_implicit(self):
        try:
            # Get the current selection from Excel
            selection = self.wb.selection
            selection_address = selection.address.replace('$', '')
            
            print(f"Selection address: {selection_address}")  # Debug info
            
            # Parse the selection to determine its type
            if ':' in selection_address:
                # Multiple cells selected
                start_cell, end_cell = selection_address.split(':')
                
                # Extract column and row information
                start_col = ''.join(filter(str.isalpha, start_cell))
                start_row = int(''.join(filter(str.isdigit, start_cell)))
                end_col = ''.join(filter(str.isalpha, end_cell))
                end_row = int(''.join(filter(str.isdigit, end_cell)))
                
                print(f"Start: {start_col}{start_row}, End: {end_col}{end_row}")  # Debug info
                
                # Determine selection type with better logic for merged cells
                if start_col == end_col:
                    # Vertical selection - look to the left
                    print("Detected as vertical selection")  # Debug info
                    self._handle_vertical_selection(start_col, start_row, end_row)
                elif start_row == end_row:
                    # Horizontal selection - look above
                    print("Detected as horizontal selection")  # Debug info
                    self._handle_horizontal_selection(start_col, end_col, start_row)
                else:
                    # Check if this might be a vertical selection with merged cells
                    # If the selection spans multiple columns but only one column has data
                    if self._is_effectively_vertical_selection(start_col, end_col, start_row, end_row):
                        print("Detected as effectively vertical selection (merged cells)")  # Debug info
                        # Use the leftmost column for vertical selection
                        self._handle_vertical_selection(start_col, start_row, end_row)
                    else:
                        # 2D selection - concatenate left + "_" + above
                        print("Detected as 2D selection")  # Debug info
                        self._handle_2d_selection(start_col, end_col, start_row, end_row)
            else:
                # Single cell selected - use original logic (look to the left)
                key = selection_address
                print("Detected as single cell selection")  # Debug info
                self._add_single_cell(key)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process selection: {str(e)}")
            print(f"Detailed error: {e}")  # Debug info
            import traceback
            traceback.print_exc()
    
    def _add_single_cell(self, key):
        """Handle single cell selection - look to the left"""
        try:
            col = ''.join(filter(str.isalpha, key))
            row = int(''.join(filter(str.isdigit, key)))
            
            # Convert column letter to number for easier manipulation
            col_num = self.column_letter_to_number(col)
            # Look leftward until we find a non-empty cell or reach column A
            for i in range(col_num - 1, 0, -1):
                check_col = self.column_number_to_letter(i)
                check_cell = f"{check_col}{row}"
                cell_value = self.get_cell_value(check_cell)
                if cell_value is not None and str(cell_value).strip() != "":
                    self.coords_dict[key] = cell_value
                    self.update_listbox()
                    return
            
            messagebox.showwarning("Warning", f"No non-empty cells found to the left of {key} in row {row}.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get value from cell to the left: {str(e)}")
    
    def _handle_vertical_selection(self, col, start_row, end_row):
        """Handle vertical selection - look to the left for each cell (transpose of horizontal)"""
        try:
            col_num = self.column_letter_to_number(col)
            
            # Process each cell in the vertical selection individually
            for row in range(start_row, end_row + 1):
                cell_key = f"{col}{row}"
                
                # Check if this cell is part of a merged range and not the top-left cell
                if self.is_merged_cell_not_top_left(cell_key):
                    continue  # Skip this cell as it's part of a merged range
                
                # Look leftward to find the non-empty cell for this specific row
                left_value = None
                for i in range(col_num - 1, 0, -1):
                    check_col = self.column_number_to_letter(i)
                    check_cell = f"{check_col}{row}"
                    cell_value = self.get_cell_value(check_cell)
                    
                    if cell_value is not None and str(cell_value).strip() != "":
                        left_value = cell_value
                        break
                
                if left_value is not None:
                    self.coords_dict[cell_key] = left_value
                else:
                    messagebox.showwarning("Warning", f"No non-empty cells found to the left of {cell_key}.")
            
            self.update_listbox()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to handle vertical selection: {str(e)}")
    
    def _handle_horizontal_selection(self, start_col, end_col, row):
        """Handle horizontal selection - look above for each cell"""
        try:
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
            # Process each cell in the horizontal selection individually
            for col_num in range(start_col_num, end_col_num + 1):
                col_letter = self.column_number_to_letter(col_num)
                cell_key = f"{col_letter}{row}"
                
                # Check if this cell is part of a merged range and not the top-left cell
                if self.is_merged_cell_not_top_left(cell_key):
                    continue  # Skip this cell as it's part of a merged range
                
                # Look above to find the non-empty cell for this specific column
                above_value = None
                for check_row in range(row - 1, 0, -1):
                    check_cell = f"{col_letter}{check_row}"
                    cell_value = self.get_cell_value(check_cell)
                    
                    if cell_value is not None and str(cell_value).strip() != "":
                        above_value = cell_value
                        break
                
                if above_value is not None:
                    self.coords_dict[cell_key] = above_value
                else:
                    messagebox.showwarning("Warning", f"No non-empty cells found above {cell_key}.")
            
            self.update_listbox()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to handle horizontal selection: {str(e)}")
    
    def _handle_2d_selection(self, start_col, end_col, start_row, end_row):
        """Handle 2D selection - concatenate left value + '_' + above value for each cell"""
        try:
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
            # Process each cell in the 2D selection individually
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = self.column_number_to_letter(col_num)
                    cell_key = f"{col_letter}{row}"
                    
                    # Check if this cell is part of a merged range and not the top-left cell
                    if self.is_merged_cell_not_top_left(cell_key):
                        continue  # Skip this cell as it's part of a merged range
                    
                    # Find the left value for this specific cell (outside the selection)
                    left_value = None
                    if col_num > 1:  # Make sure we're not at column A
                        # Look for the closest non-empty cell to the left in the same row
                        for i in range(col_num - 1, 0, -1):
                            check_col = self.column_number_to_letter(i)
                            check_cell = f"{check_col}{row}"
                            
                            # Skip if this cell is within the selected region
                            if self._is_cell_in_selection(check_col, row, start_col, end_col, start_row, end_row):
                                continue
                            
                            cell_value = self.get_cell_value(check_cell)
                            if cell_value is not None and str(cell_value).strip() != "":
                                left_value = cell_value
                                print(f"Found left header for {cell_key}: {left_value} from {check_cell}")
                                break
                    
                    # Find the above value for this specific cell (outside the selection)
                    above_value = None
                    # Look for the closest non-empty cell above in the same column
                    for check_row in range(row - 1, 0, -1):
                        check_cell = f"{col_letter}{check_row}"
                        
                        # Skip if this cell is within the selected region
                        if self._is_cell_in_selection(col_letter, check_row, start_col, end_col, start_row, end_row):
                            continue
                        
                        cell_value = self.get_cell_value(check_cell)
                        if cell_value is not None and str(cell_value).strip() != "":
                            above_value = cell_value
                            print(f"Found above header for {cell_key}: {above_value} from {check_cell}")
                            break
                    
                    # Create concatenated value for this specific cell
                    # Always create an entry if we have at least one header value
                    if left_value and above_value:
                        concatenated_value = f"{left_value}_{above_value}"
                        print(f"Created concatenated entry for {cell_key}: {concatenated_value}")
                    elif left_value:
                        concatenated_value = str(left_value)
                        print(f"Created left-only entry for {cell_key}: {concatenated_value}")
                    elif above_value:
                        concatenated_value = str(above_value)
                        print(f"Created above-only entry for {cell_key}: {concatenated_value}")
                    else:
                        # Even if no headers found, still create an entry for the coordinate
                        # This allows empty cells to be mapped to their location
                        concatenated_value = f"EMPTY_{cell_key}"
                        print(f"No headers found for {cell_key}, creating empty entry")
                    
                    self.coords_dict[cell_key] = concatenated_value
            
            self.update_listbox()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to handle 2D selection: {str(e)}")
    
    def get_cell_value(self, cell_address):
        """Get the value from a cell, handling merged cells"""
        try:
            cell = self.wb.sheets.active.range(cell_address)
            
            # Check if this cell is part of a merged range
            try:
                if hasattr(cell, 'merge_area') and cell.merge_area.address != cell_address:
                    # This cell is merged, get the value from the top-left cell of the merged area
                    merged_range = cell.merge_area
                    # Get the top-left cell address from the merged range
                    top_left_address = merged_range.address.split(':')[0].replace('$', '')
                    # Get the value from the top-left cell
                    top_left_cell = self.wb.sheets.active.range(top_left_address)
                    return top_left_cell.value
                else:
                    # This is a regular cell
                    return cell.value
            except Exception as merge_error:
                # If we can't access merge_area, just get the cell value directly
                print(f"Warning: Could not check merge area for {cell_address}: {merge_error}")
                return cell.value
                
        except Exception as e:
            print(f"Error getting value from {cell_address}: {e}")
            return None
    
    def is_merged_cell_not_top_left(self, cell_address):
        """Check if a cell is part of a merged range but not the top-left cell"""
        try:
            cell = self.wb.sheets.active.range(cell_address)
            
            # Check if this cell is part of a merged range
            try:
                if hasattr(cell, 'merge_area') and cell.merge_area.address != cell_address:
                    # This cell is part of a merged range
                    merged_range = cell.merge_area
                    # Get the top-left cell address from the merged range
                    top_left_address = merged_range.address.split(':')[0].replace('$', '')
                    # Return True if this cell is NOT the top-left cell
                    return cell_address != top_left_address
                else:
                    # This is not a merged cell or it's the top-left cell of its merge area
                    return False
            except Exception as merge_error:
                # If we can't access merge_area, assume it's not a merged cell
                print(f"Warning: Could not check merge area for {cell_address}: {merge_error}")
                return False
                
        except Exception as e:
            print(f"Error checking merged cell status for {cell_address}: {e}")
            return False
    
    def _is_effectively_vertical_selection(self, start_col, end_col, start_row, end_row):
        """Check if a selection that spans multiple columns is effectively a vertical selection due to merged cells"""
        try:
            # If it's only one column, it's definitely vertical
            if start_col == end_col:
                return True
            
            # Check if the selection is very narrow (1-2 columns wide) and tall
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            width = end_col_num - start_col_num + 1
            height = end_row - start_row + 1
            
            print(f"Selection dimensions: {width} cols x {height} rows")
            
            # If it's narrow and tall, it's likely a vertical selection with merged cells
            if width <= 2 and height > width:
                print(f"Selection is narrow ({width} cols) and tall ({height} rows) - likely vertical with merged cells")
                return True
            
            # For wider selections (3+ columns), check if it's actually vertical with merged cells
            if width >= 3:
                # Check if this is a vertical selection where each row has only one logical value
                # (either in leftmost column or in a merged cell spanning the row)
                rows_with_single_value = 0
                total_rows = height
                total_values_found = 0
                
                for row in range(start_row, end_row + 1):
                    values_in_row = 0
                    
                    # Check all columns in this row
                    for col_num in range(start_col_num, end_col_num + 1):
                        col_letter = self.column_number_to_letter(col_num)
                        check_cell = f"{col_letter}{row}"
                        
                        # Skip merged cells that aren't top-left
                        if self.is_merged_cell_not_top_left(check_cell):
                            continue
                            
                        check_value = self.get_cell_value(check_cell)
                        if check_value is not None and str(check_value).strip() != "":
                            values_in_row += 1
                            total_values_found += 1
                            print(f"Found value in row {row}: {check_cell} = '{check_value}'")
                    
                    # If this row has exactly one value, it's behaving like a vertical selection
                    if values_in_row <= 1:
                        rows_with_single_value += 1
                        print(f"Row {row} has {values_in_row} values - consistent with vertical selection")
                    else:
                        print(f"Row {row} has {values_in_row} values - suggests 2D data")
                
                # If we found no values at all, check if there are distinct column headers
                if total_values_found == 0:
                    # Check if there are different headers above each column (indicating 2D structure)
                    distinct_headers = set()
                    header_row = start_row - 1
                    
                    for col_num in range(start_col_num, end_col_num + 1):
                        col_letter = self.column_number_to_letter(col_num)
                        header_cell = f"{col_letter}{header_row}"
                        header_value = self.get_cell_value(header_cell)
                        if header_value and str(header_value).strip():
                            distinct_headers.add(str(header_value).strip())
                            print(f"Found column header: {header_cell} = '{header_value}'")
                    
                    # If we have multiple distinct column headers, it's a 2D structure
                    if len(distinct_headers) > 1:
                        print(f"Found {len(distinct_headers)} distinct column headers - treating as 2D selection")
                        return False
                    else:
                        print(f"Found {len(distinct_headers)} distinct column headers - treating as vertical selection")
                        return True
                
                # If most rows have only one value, treat as vertical selection
                if rows_with_single_value >= (total_rows * 0.8):  # 80% or more rows have single value
                    print(f"Wide selection with single values per row ({rows_with_single_value}/{total_rows}) - treating as vertical with merged cells")
                    return True
                else:
                    print(f"Wide selection ({width} cols) - treating as 2D to get both headers")
                    return False
            
            # Additional check: see if most cells in the selection are empty or part of merged ranges
            non_empty_cells = 0
            merged_cells = 0
            total_cells = 0
            
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = self.column_number_to_letter(col_num)
                    cell_key = f"{col_letter}{row}"
                    total_cells += 1
                    
                    # Check if it's a merged cell that's not the top-left
                    if self.is_merged_cell_not_top_left(cell_key):
                        merged_cells += 1
                        print(f"Found merged cell (not top-left): {cell_key}")
                        continue
                    
                    # Check if the cell has a value
                    cell_value = self.get_cell_value(cell_key)
                    if cell_value is not None and str(cell_value).strip() != "":
                        non_empty_cells += 1
                        print(f"Found non-empty cell: {cell_key} = '{cell_value}'")
                    else:
                        print(f"Found empty cell: {cell_key}")
            
            print(f"Selection analysis: {non_empty_cells} non-empty, {merged_cells} merged, {total_cells} total")
            
            # Be more conservative - only treat as vertical if almost all cells are empty/merged
            if total_cells > 0 and (non_empty_cells / total_cells) < 0.1:
                print(f"Almost all cells are empty/merged ({non_empty_cells}/{total_cells}) - treating as vertical")
                return True
            
            return False
            
        except Exception as e:
            print(f"Error checking if selection is effectively vertical: {e}")
            return False
    
    def _is_cell_in_selection(self, col, row, start_col, end_col, start_row, end_row):
        """Check if a cell is within the selected region"""
        try:
            # Convert column letters to numbers for comparison
            col_num = self.column_letter_to_number(col)
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
            # Check if the cell is within the selection bounds
            return (start_col_num <= col_num <= end_col_num and 
                    start_row <= row <= end_row)
                    
        except Exception as e:
            print(f"Error checking if cell is in selection: {e}")
            return False
    
    def column_letter_to_number(self, column_letter):
        """Convert Excel column letter to number (A=1, B=2, etc.)"""
        result = 0
        for char in column_letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    def column_number_to_letter(self, column_number):
        """Convert Excel column number to letter (1=A, 2=B, etc.)"""
        result = ""
        while column_number > 0:
            column_number -= 1
            result = chr(column_number % 26 + ord('A')) + result
            column_number //= 26
        return result

    def add_explicit(self):
        # This is the original add_entry method
        key = self.coord_entry.get()
        value = simpledialog.askstring("Input", f"Enter the value for '{key}' (leave blank to use cell above)")
        if key:
            if value is None:  # User pressed Cancel
                return
            if value == "":  # User didn't enter a value
                try:
                    # Get the cell above
                    col = ''.join(filter(str.isalpha, key))
                    row = int(''.join(filter(str.isdigit, key)))
                    cell_above = f"{col}{row - 1}"

                    # Get the value from the cell above
                    value = self.wb.sheets.active.range(cell_above).value

                    if value is None:
                        messagebox.showwarning("Warning", f"Cell {cell_above} is empty. Please enter a value manually.")
                        return
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to get value from cell above: {str(e)}")
                    return

            self.coords_dict[key] = value
            self.update_listbox()

    def remove_entry(self):
        try:
            selected_item = self.listbox.curselection()
            key = self.listbox.get(selected_item)
            key = key.split(": ", 1)[0]
            del self.coords_dict[key]
            self.update_listbox()
        except IndexError:
            pass

    def clear_all(self):
        self.coords_dict.clear()
        self.update_listbox()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for key, value in self.coords_dict.items():
            self.listbox.insert(tk.END, f"{key}: {value}")

    def update_entry(self):
        try:
            # Check if workbook is still open
            if not self._is_workbook_open():
                print("Excel workbook was closed, closing coordinate mapper...")
                self.on_closing()
                return
                
            current_selection = self.wb.selection.address
            current_selection = current_selection.split(':')[0]
            current_selection = current_selection.replace('$', '')
            self.coord_entry.delete(0, tk.END)
            self.coord_entry.insert(0, current_selection)
        except Exception as e:
            print(f"Error updating entry: {e}")
            # If we can't access the workbook, it might be closed
            if not self._is_workbook_open():
                print("Excel workbook appears to be closed, closing coordinate mapper...")
                self.on_closing()
                return
        self.coords_window.after(200, self.update_entry)

    def _is_workbook_open(self):
        """Check if the Excel workbook is still open and accessible"""
        try:
            if not self.wb:
                return False
            
            # Try to access the workbook's name or any property
            # This will raise an exception if the workbook is closed
            _ = self.wb.name
            return True
        except Exception:
            # Workbook is closed or no longer accessible
            return False

    def on_closing(self):
        if self.wb:
            try:
                self.wb.close()
            except Exception as e:
                print(f"Error closing workbook: {e}")
        self.coords_window.destroy()

    def get_result(self):
        return self.coords_dict


def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    xlsxpath=r"C:\Users\dcaoili\OneDrive - Samuel Engineering\Documents\ON-OFF VALVES - WKM DynaSeal Ball copy.xlsx"
    generator = CoordsToFieldsGenerator(root, xlsxpath, {})

    # Example usage
    generator.generate()

    # Wait for the window to close before exiting
    root.wait_window(generator.coords_window)

    result = generator.get_result()
    print("Final coordinates dictionary:")
    print(result)


if __name__ == "__main__":
    main()
