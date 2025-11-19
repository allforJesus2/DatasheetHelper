import tkinter as tk
from tkinter import simpledialog, messagebox
import xlwings as xw
from main_functions import center_window_over_parent


class CoordsToFieldsGenerator:
    MAX_DEPTH = 5  # Default maximum header depth for auto mode
    
    def __init__(self, parent, xlsx_path, initial_coords_dict):
        self.parent = parent
        self.xlsx_path = xlsx_path
        self.coords_dict = initial_coords_dict or {}
        self.wb = None
        self.coords_window = None
        self.region_cache = None  # Will store internal representation of data region

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

        add_button = tk.Button(frame1, text="Auto Add", command=self.add_implicit)
        add_button.pack(side=tk.LEFT)

        explicit_add_button = tk.Button(frame1, text="Add Explicit", command=self.add_explicit)
        explicit_add_button.pack(side=tk.LEFT)

        manual_add_button = tk.Button(frame1, text="Manual Add", command=self.add_manual)
        manual_add_button.pack(side=tk.LEFT)

        remove_button = tk.Button(frame1, text="Remove Entry", command=self.remove_entry)
        remove_button.pack(side=tk.LEFT)

        clear_button = tk.Button(frame1, text="Clear All", command=self.clear_all)
        clear_button.pack(side=tk.LEFT)

        # Options frame for checkbox
        options_frame = tk.Frame(self.coords_window)
        options_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.deduplicate_blanks_var = tk.IntVar(value=0)  # Default unchecked
        deduplicate_blanks_checkbox = tk.Checkbutton(
            options_frame,
            text="Deduplicate blank cells (keep searching until no blanks)",
            variable=self.deduplicate_blanks_var
        )
        deduplicate_blanks_checkbox.pack(side=tk.LEFT)
        
        # Manual header levels frame
        manual_levels_frame = tk.Frame(self.coords_window)
        manual_levels_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.manual_levels_var = tk.IntVar(value=0)  # Default unchecked (for auto mode)
        
        # Left header levels entry
        tk.Label(manual_levels_frame, text="Left levels:").pack(side=tk.LEFT, padx=(0, 2))
        self.left_levels_entry = tk.Entry(manual_levels_frame, width=5)
        self.left_levels_entry.insert(0, "1")
        self.left_levels_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # Top header levels entry
        tk.Label(manual_levels_frame, text="Top levels:").pack(side=tk.LEFT, padx=(0, 2))
        self.top_levels_entry = tk.Entry(manual_levels_frame, width=5)
        self.top_levels_entry.insert(0, "0")
        self.top_levels_entry.pack(side=tk.LEFT)
        
        # Hint label
        tk.Label(manual_levels_frame, text="(0 = skip direction)", font=("", 8), fg="gray").pack(side=tk.LEFT, padx=(5, 0))
        
        # Custom prefix frame
        prefix_frame = tk.Frame(self.coords_window)
        prefix_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(prefix_frame, text="Custom Prefix:").pack(side=tk.LEFT, padx=(0, 5))
        self.custom_prefix_entry = tk.Entry(prefix_frame)
        self.custom_prefix_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Hint label for custom prefix
        tk.Label(prefix_frame, text="(Optional prefix for captured field names)", font=("", 8), fg="gray").pack(side=tk.LEFT)

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
        
        # Center the window over the parent window
        center_window_over_parent(self.coords_window)
        
        self.coords_window.after(200, self.update_entry)

    def add_manual(self):
        """Add coordinates using manual mode (extracts both left & top headers)"""
        # Temporarily enable manual mode
        self.manual_levels_var.set(1)
        try:
            # Call the same logic as add_implicit but with manual mode enabled
            self.add_implicit()
        finally:
            # Reset to auto mode after operation
            self.manual_levels_var.set(0)
    
    def _build_region_cache(self, start_col, end_col, start_row, end_row):
        """Build internal representation of data region including surrounding header areas
        
        PERFORMANCE: This is the key optimization! Instead of making hundreds/thousands of 
        individual COM calls to Excel (very slow), we:
        1. Read the entire region + header area in ONE bulk operation
        2. Store values and merge info in Python data structures (dict/list)
        3. All subsequent operations use this cache (100-1000x faster)
        
        For a 20x20 selection with 3 header levels, this reduces ~400 Excel calls to just 1!
        
        Args:
            start_col, end_col: Column letters for selection bounds
            start_row, end_row: Row numbers for selection bounds
        """
        start_col_num = self.column_letter_to_number(start_col)
        end_col_num = self.column_letter_to_number(end_col)
        
        # Determine header depths based on mode
        manual_mode = self.manual_levels_var.get() == 1
        
        if manual_mode:
            # Use user-specified depths from input boxes
            try:
                # Get left depth - treat blank as 0
                left_entry = self.left_levels_entry.get().strip()
                left_depth = max(1, int(left_entry)) if left_entry else 0
                
                # Get top depth - treat blank as 0
                top_entry = self.top_levels_entry.get().strip()
                top_depth = max(1, int(top_entry)) if top_entry else 0
            except (ValueError, AttributeError):
                left_depth = 0
                top_depth = 0
        else:
            # Use class default for auto mode
            left_depth = self.MAX_DEPTH
            top_depth = self.MAX_DEPTH
        
        print(f"Cache depths: left={left_depth}, top={top_depth}")
        
        # Expand region to include potential headers (x=left, y=top)
        cache_start_col = max(1, start_col_num - left_depth)
        cache_end_col = end_col_num
        cache_start_row = max(1, start_row - top_depth)
        cache_end_row = end_row
        
        # Build range address
        cache_start_addr = f"{self.column_number_to_letter(cache_start_col)}{cache_start_row}"
        cache_end_addr = f"{self.column_number_to_letter(cache_end_col)}{cache_end_row}"
        cache_range_addr = f"{cache_start_addr}:{cache_end_addr}"
        
        print(f"Building cache for region: {cache_range_addr}")
        
        # Read entire region in one call (much faster than individual cells)
        sheet = self.wb.sheets.active
        cache_range = sheet.range(cache_range_addr)
        
        # Store values in 2D array
        values = cache_range.value if cache_range.count > 1 else [[cache_range.value]]
        if not isinstance(values, list):
            values = [[values]]
        elif values and not isinstance(values[0], list):
            values = [values]
        
        # Store merge information
        merge_info = {}
        try:
            for cell in cache_range:
                cell_addr = cell.address.replace('$', '')
                try:
                    if hasattr(cell, 'merge_area') and cell.merge_area:
                        merge_area_addr = cell.merge_area.address.replace('$', '')
                        if ':' in merge_area_addr:
                            top_left = merge_area_addr.split(':')[0]
                            merge_info[cell_addr] = {
                                'is_merged': True,
                                'top_left': top_left,
                                'is_top_left': (cell_addr == top_left)
                            }
                        else:
                            merge_info[cell_addr] = {'is_merged': False}
                    else:
                        merge_info[cell_addr] = {'is_merged': False}
                except:
                    merge_info[cell_addr] = {'is_merged': False}
        except Exception as e:
            print(f"Warning: Could not fully process merge info: {e}")
        
        # Store cache
        self.region_cache = {
            'start_col': cache_start_col,
            'end_col': cache_end_col,
            'start_row': cache_start_row,
            'end_row': cache_end_row,
            'values': values,
            'merge_info': merge_info
        }
        print(f"Cache built: {len(values)} rows, {len(values[0]) if values else 0} cols, {len(merge_info)} merge entries")
    
    def add_implicit(self):
        try:
            # Clear any previous cache for fresh operation
            self.region_cache = None
            
            selection = self.wb.selection
            selection_address = selection.address.replace('$', '')
            
            print(f"Selection address: {selection_address}")
            
            if ':' in selection_address:
                start_cell, end_cell = selection_address.split(':')
                start_col, start_row = self._parse_cell_address(start_cell)
                end_col, end_row = self._parse_cell_address(end_cell)
                
                print(f"Start: {start_col}{start_row}, End: {end_col}{end_row}")
                
                # Build cache for the region (this is the key performance optimization!)
                self._build_region_cache(start_col, end_col, start_row, end_row)
                
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
                # Build cache for single cell too
                col, row = self._parse_cell_address(selection_address)
                self._build_region_cache(col, col, row, row)
                self._handle_single_cell(selection_address)
                
        except Exception as e:
            self._show_error("Failed to process selection", e)

    def _parse_cell_address(self, cell_address):
        """Parse cell address into column and row components"""
        col = ''.join(filter(str.isalpha, cell_address))
        row = int(''.join(filter(str.isdigit, cell_address)))
        return col, row
    
    def _extract_manual_headers(self, cell_key):
        """Extract headers for a cell using manual mode settings
        
        Gets headers from both left and top directions based on user-specified levels.
        
        Args:
            cell_key: The cell address to extract headers for
            
        Returns:
            Concatenated header string or None if no headers found
        """
        try:
            # Get left levels - treat blank as 0
            left_entry = self.left_levels_entry.get().strip()
            left_levels = max(0, int(left_entry)) if left_entry else 0
            
            # Get top levels - treat blank as 0
            top_entry = self.top_levels_entry.get().strip()
            top_levels = max(0, int(top_entry)) if top_entry else 0
        except ValueError:
            # If there's a ValueError (non-numeric input), default to 0 (skip direction)
            left_levels = 0
            top_levels = 0
        
        # Get left headers (only if left_levels > 0)
        left_parts = []
        if left_levels > 0:
            for level in range(left_levels):
                header = self._find_header_in_direction(cell_key, 'left', skip_count=level)
                if header:
                    left_parts.append(str(header).strip())
        
        # Get top headers (only if top_levels > 0)
        top_parts = []
        if top_levels > 0:
            for level in range(top_levels):
                header = self._find_header_in_direction(cell_key, 'up', skip_count=level)
                if header:
                    top_parts.append(str(header).strip())
        
        # Combine both directions
        left_value = "_".join(reversed(left_parts)) if left_parts else None
        top_value = "_".join(reversed(top_parts)) if top_parts else None
        
        # Create final value
        if left_value and top_value:
            final_value = f"{left_value}_{top_value}"
        elif left_value:
            final_value = left_value
        elif top_value:
            final_value = top_value
        else:
            return None
        
        # Apply custom prefix if provided
        custom_prefix = self.custom_prefix_entry.get().strip()
        if custom_prefix:
            return f"{custom_prefix}_{final_value}"
        else:
            return final_value

    def _handle_single_cell(self, cell_key):
        """Handle single cell selection - extracts headers from both directions in manual mode"""
        try:
            col, row = self._parse_cell_address(cell_key)
            manual_mode = self.manual_levels_var.get() == 1
            
            if manual_mode:
                final_value = self._extract_manual_headers(cell_key)
            else:
                # Auto mode: just get first level from left
                final_value = self._find_header_in_direction(cell_key, 'left')
            
            if final_value:
                self.coords_dict[cell_key] = final_value
                self.update_listbox()
            else:
                messagebox.showwarning("Warning", f"No non-empty cells found to the left of {cell_key} in row {row}.")
        except Exception as e:
            self._show_error("Failed to get headers for cell", e)

    def _handle_directional_selection(self, direction, fixed_coord, start_range, end_range):
        """Unified handler for both vertical and horizontal selections
        
        In manual mode: extracts headers from BOTH directions regardless of selection type
        In auto mode: uses deduplication logic for the primary direction
        """
        try:
            manual_mode = self.manual_levels_var.get() == 1
            
            # Collect all cell keys first (excluding merged cells that are not top-left)
            cell_keys = []
            for pos in range(start_range, end_range + 1):
                if direction == 'vertical':
                    cell_key = f"{fixed_coord}{pos}"
                else:
                    col_letter = self.column_number_to_letter(pos)
                    cell_key = f"{col_letter}{fixed_coord}"
                
                if not self.is_merged_cell_not_top_left(cell_key):
                    cell_keys.append(cell_key)
            
            if manual_mode:
                # Manual mode: extract from specified directions for each cell
                for cell_key in cell_keys:
                    final_value = self._extract_manual_headers(cell_key)
                    if final_value:
                        self.coords_dict[cell_key] = final_value
            else:
                # Auto mode: get headers with iterative deduplication (original behavior)
                headers_dict = self._get_headers_with_deduplication(cell_keys, direction)
                
                # Add to coords_dict
                for cell_key, header_value in headers_dict.items():
                    if header_value:
                        self.coords_dict[cell_key] = header_value
                    else:
                        direction_text = "to the left" if direction == 'vertical' else "above"
                        messagebox.showwarning("Warning", f"No non-empty cells found {direction_text} of {cell_key}.")
            
            self.update_listbox()
            
        except Exception as e:
            direction_text = "vertical" if direction == 'vertical' else "horizontal"
            self._show_error(f"Failed to handle {direction_text} selection", e)

    def _find_header_in_direction(self, cell_key, direction, skip_count=0):
        """Find header value in specified direction from a cell
        
        Args:
            cell_key: The cell to search from
            direction: 'left' or 'up'
            skip_count: Number of non-empty cells to skip before returning value
        """
        try:
            col, row = self._parse_cell_address(cell_key)
            found_count = 0
            
            if direction == 'left':
                col_num = self.column_letter_to_number(col)
                for i in range(col_num - 1, 0, -1):
                    check_col = self.column_number_to_letter(i)
                    check_cell = f"{check_col}{row}"
                    cell_value = self.get_cell_value(check_cell)
                    if cell_value and str(cell_value).strip():
                        if found_count == skip_count:
                            return cell_value
                        found_count += 1
            elif direction == 'up':
                for check_row in range(row - 1, 0, -1):
                    check_cell = f"{col}{check_row}"
                    cell_value = self.get_cell_value(check_cell)
                    if cell_value and str(cell_value).strip():
                        if found_count == skip_count:
                            return cell_value
                        found_count += 1
            
            return None
            
        except Exception as e:
            print(f"Error finding header in {direction} direction for {cell_key}: {e}")
            return None
    
    def _get_headers_with_deduplication(self, cell_keys, direction, max_depth=None):
        """Get headers for all cells with iterative deduplication
        
        Args:
            cell_keys: List of cell addresses
            direction: 'vertical' (search left) or 'horizontal' (search up)
            max_depth: Maximum number of levels to search back (defaults to self.MAX_DEPTH)
            
        Returns:
            Dictionary mapping cell_key to concatenated header string
        """
        if max_depth is None:
            max_depth = self.MAX_DEPTH
            
        search_direction = 'left' if direction == 'vertical' else 'up'
        deduplicate_blanks = self.deduplicate_blanks_var.get() == 1
        manual_mode = self.manual_levels_var.get() == 1
        
        # Determine how many levels to extract
        if manual_mode:
            # Use manual settings
            try:
                if direction == 'vertical':
                    # Vertical selection uses left headers - treat blank as 0
                    left_entry = self.left_levels_entry.get().strip()
                    specified_levels = int(left_entry) if left_entry else 0
                else:
                    # Horizontal selection uses top headers - treat blank as 0
                    top_entry = self.top_levels_entry.get().strip()
                    specified_levels = int(top_entry) if top_entry else 0
                
                # Clamp to valid range (0 means skip, so no minimum)
                specified_levels = max(0, min(specified_levels, max_depth))
                print(f"Manual mode: extracting {specified_levels} levels for {direction} selection")
            except ValueError:
                print("Invalid manual level value, defaulting to 0")
                specified_levels = 0
        
        # Build headers level by level
        headers_by_level = []  # List of lists, where headers_by_level[0] is the first level back
        
        for level in range(max_depth):
            level_headers = []
            for cell_key in cell_keys:
                header = self._find_header_in_direction(cell_key, search_direction, skip_count=level)
                # Convert to string, treating None and empty as ""
                header_str = str(header).strip() if header is not None else ""
                level_headers.append(header_str)
            
            headers_by_level.append(level_headers)
            print(f"Level {level} headers: {level_headers}")
            
            if manual_mode:
                # In manual mode, continue until we reach the specified number of levels
                if level + 1 >= specified_levels:
                    print(f"Reached manual level limit ({specified_levels}), stopping")
                    break
            else:
                # Auto mode: check if we have unique headers at this level
                if self._has_duplicates(level_headers, count_blanks=deduplicate_blanks):
                    print(f"Duplicates found at level {level}, continuing to next level...")
                    continue
                else:
                    print(f"No duplicates at level {level}, stopping deduplication")
                    break
        
        # Concatenate all levels to create final headers
        final_headers = {}
        for i, cell_key in enumerate(cell_keys):
            header_parts = []
            # Collect all levels from outermost (deepest) to innermost (first level)
            for level in range(len(headers_by_level) - 1, -1, -1):
                part = headers_by_level[level][i]
                if part:  # Only add non-empty parts
                    header_parts.append(part)
            
            # Concatenate with underscore
            if header_parts:
                final_header = "_".join(header_parts)
            else:
                final_header = ""
            
            # Apply custom prefix if provided
            custom_prefix = self.custom_prefix_entry.get().strip()
            if custom_prefix and final_header:
                final_header = f"{custom_prefix}_{final_header}"
            
            final_headers[cell_key] = final_header
            print(f"Final header for {cell_key}: {final_header}")
        
        return final_headers
    
    def _has_duplicates(self, values, count_blanks=True):
        """Check if a list has any duplicate values
        
        Args:
            values: List of string values
            count_blanks: If True, blank cells are treated as duplicates with each other.
                         If False, blank cells are ignored in duplicate detection.
            
        Returns:
            True if duplicates exist, False otherwise
        """
        # Count occurrences of each value
        value_counts = {}
        for val in values:
            # If count_blanks is False and value is blank, skip it
            if not count_blanks and val == "":
                continue
            value_counts[val] = value_counts.get(val, 0) + 1
        
        # Check if any value appears more than once
        for count in value_counts.values():
            if count > 1:
                return True
        
        return False

    def _handle_2d_selection(self, start_col, end_col, start_row, end_row):
        """Handle 2D selection - concatenate left headers + '_' + above headers for each cell"""
        try:
            start_col_num = self.column_letter_to_number(start_col)
            end_col_num = self.column_letter_to_number(end_col)
            
            manual_mode = self.manual_levels_var.get() == 1
            
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = self.column_number_to_letter(col_num)
                    cell_key = f"{col_letter}{row}"
                    
                    if self.is_merged_cell_not_top_left(cell_key):
                        continue
                    
                    if manual_mode:
                        # Manual mode: use unified method
                        final_value = self._extract_manual_headers(cell_key)
                        if final_value:
                            self.coords_dict[cell_key] = final_value
                    else:
                        # Auto mode: use 2D-specific logic that excludes selected region
                        deduplicate_blanks = self.deduplicate_blanks_var.get() == 1
                        
                        left_value = self._get_multilevel_header_for_2d(
                            cell_key, 'left', start_col, end_col, start_row, end_row, 
                            None, deduplicate_blanks
                        )
                        
                        above_value = self._get_multilevel_header_for_2d(
                            cell_key, 'up', start_col, end_col, start_row, end_row, 
                            None, deduplicate_blanks
                        )
                        
                        concatenated_value = self._create_concatenated_value(left_value, above_value, cell_key)
                        self.coords_dict[cell_key] = concatenated_value
            
            self.update_listbox()
            
        except Exception as e:
            self._show_error("Failed to handle 2D selection", e)

    def _get_multilevel_header_for_2d(self, cell_key, direction, start_col, end_col, start_row, end_row, 
                                       specified_levels=None, deduplicate_blanks=True, max_depth=None):
        """Extract multi-level headers for 2D selection, excluding selected region
        
        Args:
            cell_key: The cell to get headers for
            direction: 'left' or 'up'
            start_col, end_col, start_row, end_row: Selection bounds
            specified_levels: If set, extract exactly this many levels (manual mode)
            deduplicate_blanks: Whether to treat blanks as duplicates (auto mode)
            max_depth: Maximum levels to search (defaults to self.MAX_DEPTH)
            
        Returns:
            Concatenated header string with all levels joined by underscore
        """
        if max_depth is None:
            max_depth = self.MAX_DEPTH
            
        header_parts = []
        
        # Determine how many levels to extract
        if specified_levels is not None:
            # Manual mode: extract exactly the specified number of levels
            levels_to_extract = min(specified_levels, max_depth)
        else:
            # Auto mode: extract until unique (for now, default to 1 for 2D)
            # In future, could implement deduplication logic for 2D as well
            levels_to_extract = 1
        
        for level in range(levels_to_extract):
            if direction == 'left':
                header = self._find_left_header_for_2d_with_skip(
                    cell_key, start_col, end_col, start_row, end_row, skip_count=level
                )
            else:  # 'up'
                header = self._find_above_header_for_2d_with_skip(
                    cell_key, start_col, end_col, start_row, end_row, skip_count=level
                )
            
            if header:
                header_parts.append(str(header).strip())
        
        # Return concatenated headers (outermost to innermost)
        if header_parts:
            # Reverse so outermost (deepest) comes first
            return "_".join(reversed(header_parts))
        return None
    
    def _find_left_header_for_2d_with_skip(self, cell_key, start_col, end_col, start_row, end_row, skip_count=0):
        """Find left header for 2D selection, excluding selected region, with skip capability"""
        col, row = self._parse_cell_address(cell_key)
        col_num = self.column_letter_to_number(col)
        
        if col_num <= 1:
            return None
        
        found_count = 0
        for i in range(col_num - 1, 0, -1):
            check_col = self.column_number_to_letter(i)
            check_cell = f"{check_col}{row}"
            
            if self._is_cell_in_selection(check_col, row, start_col, end_col, start_row, end_row):
                continue
            
            cell_value = self.get_cell_value(check_cell)
            if cell_value and str(cell_value).strip():
                if found_count == skip_count:
                    print(f"Found left header (level {skip_count}) for {cell_key}: {cell_value} from {check_cell}")
                    return cell_value
                found_count += 1
        
        return None

    def _find_above_header_for_2d_with_skip(self, cell_key, start_col, end_col, start_row, end_row, skip_count=0):
        """Find above header for 2D selection, excluding selected region, with skip capability"""
        col, row = self._parse_cell_address(cell_key)
        
        found_count = 0
        for check_row in range(row - 1, 0, -1):
            check_cell = f"{col}{check_row}"
            
            if self._is_cell_in_selection(col, check_row, start_col, end_col, start_row, end_row):
                continue
            
            cell_value = self.get_cell_value(check_cell)
            if cell_value and str(cell_value).strip():
                if found_count == skip_count:
                    print(f"Found above header (level {skip_count}) for {cell_key}: {cell_value} from {check_cell}")
                    return cell_value
                found_count += 1
        
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
        """Get the value from a cell, handling merged cells (uses cache when available)"""
        # Try to use cache first
        if self.region_cache:
            col, row = self._parse_cell_address(cell_address)
            col_num = self.column_letter_to_number(col)
            
            cache = self.region_cache
            # Check if cell is within cached region
            if (cache['start_col'] <= col_num <= cache['end_col'] and 
                cache['start_row'] <= row <= cache['end_row']):
                
                # Check if it's a merged cell
                if cell_address in cache['merge_info']:
                    merge_data = cache['merge_info'][cell_address]
                    if merge_data['is_merged'] and not merge_data['is_top_left']:
                        # Get value from top-left cell of merged range
                        top_left = merge_data['top_left']
                        tl_col, tl_row = self._parse_cell_address(top_left)
                        tl_col_num = self.column_letter_to_number(tl_col)
                        tl_row_idx = tl_row - cache['start_row']
                        tl_col_idx = tl_col_num - cache['start_col']
                        return cache['values'][tl_row_idx][tl_col_idx]
                
                # Get value from cache
                row_idx = row - cache['start_row']
                col_idx = col_num - cache['start_col']
                if (cache['values'] and 
                    0 <= row_idx < len(cache['values']) and 
                    cache['values'][0] and
                    0 <= col_idx < len(cache['values'][0])):
                    return cache['values'][row_idx][col_idx]
        
        # Fallback to direct Excel access if not in cache
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
        """Check if a cell is part of a merged range but not the top-left cell (uses cache when available)"""
        # Try to use cache first
        if self.region_cache and cell_address in self.region_cache['merge_info']:
            merge_data = self.region_cache['merge_info'][cell_address]
            if merge_data['is_merged']:
                return not merge_data['is_top_left']
            return False
        
        # Fallback to direct Excel access if not in cache
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
    xlsxpath=r"C:\Users\dcaoili\SPEC-ME-11154 Air Diaphragm Pumps Rev 0.xlsx"
    generator = CoordsToFieldsGenerator(root, xlsxpath, {})

    generator.generate()
    root.wait_window(generator.coords_window)

    result = generator.get_result()
    print("Final coordinates dictionary:")
    print(result)


if __name__ == "__main__":
    main()
