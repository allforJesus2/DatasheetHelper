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
        
        frame3 = tk.Frame(self.coords_window)
        frame3.pack(fill=tk.X, pady=5)
        
        done_button = tk.Button(frame3, text="Done", command=self.on_closing)
        done_button.pack(side=tk.RIGHT, padx=5, pady=5)

        self.update_listbox()
        self.coords_window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.coords_window.after(200, self.update_entry)

    def add_implicit(self):
        try:
            selection = self.wb.selection
            selection_address = selection.address.replace('$', '')
            
            print(f"Selection address: {selection_address}")
            
            if ':' in selection_address:
                start_cell, end_cell = selection_address.split(':')
                start_col, start_row = self._parse_cell_address(start_cell)
                end_col, end_row = self._parse_cell_address(end_cell)
                
                print(f"Start: {start_col}{start_row}, End: {end_col}{end_row}")
                
                if start_col == end_col:
                    print("Detected as vertical selection")
                    self._handle_directional_selection('vertical', start_col, start_row, end_row)
                elif start_row == end_row:
                    print("Detected as horizontal selection")
                    # Convert column letters to numbers for horizontal selection
                    start_col_num = self.column_letter_to_number(start_col)
                    end_col_num = self.column_letter_to_number(end_col)
                    self._handle_directional_selection('horizontal', start_row, start_col_num, end_col_num)
                else:
                    if self._is_effectively_vertical_selection(start_col, end_col, start_row, end_row):
                        print("Detected as effectively vertical selection (merged cells)")
                        self._handle_directional_selection('vertical', start_col, start_row, end_row)
                    else:
                        print("Detected as 2D selection")
                        self._handle_2d_selection(start_col, end_col, start_row, end_row)
            else:
                print("Detected as single cell selection")
                self._handle_single_cell(selection_address)
                
        except Exception as e:
            self._show_error("Failed to process selection", e)

    def _parse_cell_address(self, cell_address):
        """Parse cell address into column and row components"""
        col = ''.join(filter(str.isalpha, cell_address))
        row = int(''.join(filter(str.isdigit, cell_address)))
        return col, row

    def _handle_single_cell(self, cell_key):
        """Handle single cell selection - look to the left"""
        try:
            col, row = self._parse_cell_address(cell_key)
            col_num = self.column_letter_to_number(col)
            
            left_value = self._find_header_in_direction(cell_key, 'left')
            if left_value:
                self.coords_dict[cell_key] = left_value
                self.update_listbox()
            else:
                messagebox.showwarning("Warning", f"No non-empty cells found to the left of {cell_key} in row {row}.")
        except Exception as e:
            self._show_error("Failed to get value from cell to the left", e)

    def _handle_directional_selection(self, direction, fixed_coord, start_range, end_range):
        """Unified handler for both vertical and horizontal selections"""
        try:
            for pos in range(start_range, end_range + 1):
                if direction == 'vertical':
                    cell_key = f"{fixed_coord}{pos}"
                    header_value = self._find_header_in_direction(cell_key, 'left')
                else:
                    col_letter = self.column_number_to_letter(pos)
                    cell_key = f"{col_letter}{fixed_coord}"
                    header_value = self._find_header_in_direction(cell_key, 'up')
                
                if self.is_merged_cell_not_top_left(cell_key):
                    continue
                
                if header_value:
                    self.coords_dict[cell_key] = header_value
                else:
                    direction_text = "to the left" if direction == 'vertical' else "above"
                    messagebox.showwarning("Warning", f"No non-empty cells found {direction_text} of {cell_key}.")
            
            self.update_listbox()
            
        except Exception as e:
            direction_text = "vertical" if direction == 'vertical' else "horizontal"
            self._show_error(f"Failed to handle {direction_text} selection", e)

    def _find_header_in_direction(self, cell_key, direction):
        """Find header value in specified direction from a cell"""
        try:
            col, row = self._parse_cell_address(cell_key)
            
            if direction == 'left':
                col_num = self.column_letter_to_number(col)
                for i in range(col_num - 1, 0, -1):
                    check_col = self.column_number_to_letter(i)
                    check_cell = f"{check_col}{row}"
                    cell_value = self.get_cell_value(check_cell)
                    if cell_value and str(cell_value).strip():
                        return cell_value
            elif direction == 'up':
                for check_row in range(row - 1, 0, -1):
                    check_cell = f"{col}{check_row}"
                    cell_value = self.get_cell_value(check_cell)
                    if cell_value and str(cell_value).strip():
                        return cell_value
            
            return None
            
        except Exception as e:
            print(f"Error finding header in {direction} direction for {cell_key}: {e}")
            return None

    def _handle_2d_selection(self, start_col, end_col, start_row, end_row):
        """Handle 2D selection - concatenate left value + '_' + above value for each cell"""
        try:
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = self.column_number_to_letter(col_num)
                    cell_key = f"{col_letter}{row}"
                    
                    if self.is_merged_cell_not_top_left(cell_key):
                        continue
                    
                    left_value = self._find_left_header_for_2d(cell_key, start_col, end_col, start_row, end_row)
                    above_value = self._find_above_header_for_2d(cell_key, start_col, end_col, start_row, end_row)
                    
                    concatenated_value = self._create_concatenated_value(left_value, above_value, cell_key)
                    self.coords_dict[cell_key] = concatenated_value
            
            self.update_listbox()
            
        except Exception as e:
            self._show_error("Failed to handle 2D selection", e)

    def _find_left_header_for_2d(self, cell_key, start_col, end_col, start_row, end_row):
        """Find left header for 2D selection, excluding selected region"""
        col, row = self._parse_cell_address(cell_key)
        col_num = self.column_letter_to_number(col)
        
        if col_num <= 1:
            return None
            
        for i in range(col_num - 1, 0, -1):
            check_col = self.column_number_to_letter(i)
            check_cell = f"{check_col}{row}"
            
            if self._is_cell_in_selection(check_col, row, start_col, end_col, start_row, end_row):
                continue
            
            cell_value = self.get_cell_value(check_cell)
            if cell_value and str(cell_value).strip():
                print(f"Found left header for {cell_key}: {cell_value} from {check_cell}")
                return cell_value
        
        return None

    def _find_above_header_for_2d(self, cell_key, start_col, end_col, start_row, end_row):
        """Find above header for 2D selection, excluding selected region"""
        col, row = self._parse_cell_address(cell_key)
        
        for check_row in range(row - 1, 0, -1):
            check_cell = f"{col}{check_row}"
            
            if self._is_cell_in_selection(col, check_row, start_col, end_col, start_row, end_row):
                continue
            
            cell_value = self.get_cell_value(check_cell)
            if cell_value and str(cell_value).strip():
                print(f"Found above header for {cell_key}: {cell_value} from {check_cell}")
                return cell_value
        
        return None

    def _create_concatenated_value(self, left_value, above_value, cell_key):
        """Create concatenated value for 2D selection"""
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
            concatenated_value = f"EMPTY_{cell_key}"
            print(f"No headers found for {cell_key}, creating empty entry")
        
        return concatenated_value

    def get_cell_value(self, cell_address):
        """Get the value from a cell, handling merged cells"""
        try:
            cell = self.wb.sheets.active.range(cell_address)
            
            try:
                if hasattr(cell, 'merge_area') and cell.merge_area.address != cell_address:
                    merged_range = cell.merge_area
                    top_left_address = merged_range.address.split(':')[0].replace('$', '')
                    top_left_cell = self.wb.sheets.active.range(top_left_address)
                    return top_left_cell.value
                else:
                    return cell.value
            except Exception as merge_error:
                print(f"Warning: Could not check merge area for {cell_address}: {merge_error}")
                return cell.value
                
        except Exception as e:
            print(f"Error getting value from {cell_address}: {e}")
            return None

    def is_merged_cell_not_top_left(self, cell_address):
        """Check if a cell is part of a merged range but not the top-left cell"""
        try:
            cell = self.wb.sheets.active.range(cell_address)
            
            try:
                if hasattr(cell, 'merge_area') and cell.merge_area.address != cell_address:
                    merged_range = cell.merge_area
                    top_left_address = merged_range.address.split(':')[0].replace('$', '')
                    return cell_address != top_left_address
                else:
                    return False
            except Exception as merge_error:
                print(f"Warning: Could not check merge area for {cell_address}: {merge_error}")
                return False
                
        except Exception as e:
            print(f"Error checking merged cell status for {cell_address}: {e}")
            return False

    def _is_effectively_vertical_selection(self, start_col, end_col, start_row, end_row):
        """Check if a selection that spans multiple columns is effectively a vertical selection due to merged cells"""
        try:
            if start_col == end_col:
                return True
            
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            width = end_col_num - start_col_num + 1
            height = end_row - start_row + 1
            
            print(f"Selection dimensions: {width} cols x {height} rows")
            
            if width <= 2 and height > width:
                print(f"Selection is narrow ({width} cols) and tall ({height} rows) - likely vertical with merged cells")
                return True
            
            if width >= 3:
                rows_with_single_value = 0
                total_rows = height
                total_values_found = 0
                
                for row in range(start_row, end_row + 1):
                    values_in_row = 0
                    
                    for col_num in range(start_col_num, end_col_num + 1):
                        col_letter = self.column_number_to_letter(col_num)
                        check_cell = f"{col_letter}{row}"
                        
                        if self.is_merged_cell_not_top_left(check_cell):
                            continue
                            
                        check_value = self.get_cell_value(check_cell)
                        if check_value and str(check_value).strip():
                            values_in_row += 1
                            total_values_found += 1
                            print(f"Found value in row {row}: {check_cell} = '{check_value}'")
                    
                    if values_in_row <= 1:
                        rows_with_single_value += 1
                        print(f"Row {row} has {values_in_row} values - consistent with vertical selection")
                    else:
                        print(f"Row {row} has {values_in_row} values - suggests 2D data")
                
                if total_values_found == 0:
                    distinct_headers = self._get_distinct_headers(start_col_num, end_col_num, start_row)
                    if len(distinct_headers) > 1:
                        print(f"Found {len(distinct_headers)} distinct column headers - treating as 2D selection")
                        return False
                    else:
                        print(f"Found {len(distinct_headers)} distinct column headers - treating as vertical selection")
                        return True
                
                if rows_with_single_value >= (total_rows * 0.8):
                    print(f"Wide selection with single values per row ({rows_with_single_value}/{total_rows}) - treating as vertical with merged cells")
                    return True
                else:
                    print(f"Wide selection ({width} cols) - treating as 2D to get both headers")
                    return False
            
            non_empty_cells = 0
            merged_cells = 0
            total_cells = 0
            
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = self.column_number_to_letter(col_num)
                    cell_key = f"{col_letter}{row}"
                    total_cells += 1
                    
                    if self.is_merged_cell_not_top_left(cell_key):
                        merged_cells += 1
                        print(f"Found merged cell (not top-left): {cell_key}")
                        continue
                    
                    cell_value = self.get_cell_value(cell_key)
                    if cell_value and str(cell_value).strip():
                        non_empty_cells += 1
                        print(f"Found non-empty cell: {cell_key} = '{cell_value}'")
                    else:
                        print(f"Found empty cell: {cell_key}")
            
            print(f"Selection analysis: {non_empty_cells} non-empty, {merged_cells} merged, {total_cells} total")
            
            if total_cells > 0 and (non_empty_cells / total_cells) < 0.1:
                print(f"Almost all cells are empty/merged ({non_empty_cells}/{total_cells}) - treating as vertical")
                return True
            
            return False
            
        except Exception as e:
            print(f"Error checking if selection is effectively vertical: {e}")
            return False

    def _get_distinct_headers(self, start_col_num, end_col_num, start_row):
        """Get distinct column headers above the selection"""
        distinct_headers = set()
        header_row = start_row - 1
        
        for col_num in range(start_col_num, end_col_num + 1):
            col_letter = self.column_number_to_letter(col_num)
            header_cell = f"{col_letter}{header_row}"
            header_value = self.get_cell_value(header_cell)
            if header_value and str(header_value).strip():
                distinct_headers.add(str(header_value).strip())
                print(f"Found column header: {header_cell} = '{header_value}'")
        
        return distinct_headers

    def _is_cell_in_selection(self, col, row, start_col, end_col, start_row, end_row):
        """Check if a cell is within the selected region"""
        try:
            col_num = self.column_letter_to_number(col)
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
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

    def _show_error(self, message, exception):
        """Centralized error handling"""
        messagebox.showerror("Error", f"{message}: {str(exception)}")
        print(f"Detailed error: {exception}")
        import traceback
        traceback.print_exc()

    def add_explicit(self):
        key = self.coord_entry.get()
        value = simpledialog.askstring("Input", f"Enter the value for '{key}' (leave blank to use cell above)")
        if key:
            if value is None:
                return
            if value == "":
                try:
                    col, row = self._parse_cell_address(key)
                    cell_above = f"{col}{row - 1}"
                    value = self.wb.sheets.active.range(cell_above).value

                    if value is None:
                        messagebox.showwarning("Warning", f"Cell {cell_above} is empty. Please enter a value manually.")
                        return
                except Exception as e:
                    self._show_error("Failed to get value from cell above", e)
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
    root.withdraw()
    xlsxpath=r"C:\Users\dcaoili\OneDrive - Samuel Engineering\Documents\ON-OFF VALVES - WKM DynaSeal Ball copy.xlsx"
    generator = CoordsToFieldsGenerator(root, xlsxpath, {})

    generator.generate()
    root.wait_window(generator.coords_window)

    result = generator.get_result()
    print("Final coordinates dictionary:")
    print(result)


if __name__ == "__main__":
    main()
