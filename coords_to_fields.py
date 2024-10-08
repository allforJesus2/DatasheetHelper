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

        self.update_listbox()

        self.coords_window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.coords_window.after(200, self.update_entry)

    def add_implicit(self):
        key = self.coord_entry.get()
        if key:
            try:
                # Get the cell above
                col = ''.join(filter(str.isalpha, key))
                row = int(''.join(filter(str.isdigit, key)))
                cell_above = f"{col}{row - 1}"

                # Get the value from the cell above
                value = self.wb.sheets.active.range(cell_above).value

                if value is None:
                    messagebox.showwarning("Warning", f"Cell {cell_above} is empty. Using empty string instead.")
                    value = ""

                self.coords_dict[key] = value
                self.update_listbox()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to get value from cell above: {str(e)}")

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
            current_selection = self.wb.selection.address
            current_selection = current_selection.split(':')[0]
            current_selection = current_selection.replace('$', '')
            self.coord_entry.delete(0, tk.END)
            self.coord_entry.insert(0, current_selection)
        except Exception as e:
            print(f"Error updating entry: {e}")
        self.coords_window.after(200, self.update_entry)

    def on_closing(self):
        if self.wb:
            self.wb.close()
        self.coords_window.destroy()

    def get_result(self):
        return self.coords_dict


def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    xlsxpath=r"\\stor-dn-01.se.local\projects\Projects\22058_BPX_Modeling\Engineering\Instrumentation\DataSheets\Archive\TT Temperature Transmitters NWH.xlsx"
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
