import xlwings as xw
import os
import shutil
import tempfile
from pathlib import Path


class ExcelManager:
    """
    A class to manage Excel workbook interactions with special handling for network paths
    """

    def __init__(self):
        self.app = None
        self.wb = None
        self.is_dirty = False
        self.original_path = None
        self.temp_path = None

    def _is_network_path(self, path):
        """Check if the path is a network path"""
        return path.startswith('\\\\') or ':' in path and not path.startswith(('C:', 'D:'))

    def _create_temp_copy(self, network_path):
        """Create a temporary local copy of the network file"""
        try:
            # Create temp file with same extension
            ext = os.path.splitext(network_path)[1]
            temp_fd, temp_path = tempfile.mkstemp(suffix=ext)
            os.close(temp_fd)  # Close the file descriptor

            # Copy network file to temp location
            shutil.copy2(network_path, temp_path)
            return temp_path
        except Exception as e:
            print(f"Error creating temp copy: {e}")
            return None

    def _save_back_to_network(self):
        """Save the temporary file back to the network location"""
        try:
            if self.temp_path and self.original_path:
                # First save the workbook in temp location
                self.wb.save()

                # Close the workbook
                self.wb.close()
                self.wb = None

                # Copy back to network
                shutil.copy2(self.temp_path, self.original_path)

                # Clean up temp file
                try:
                    os.remove(self.temp_path)
                except:
                    pass  # Ignore cleanup errors

                self.temp_path = None
                self.is_dirty = False
                return True
        except Exception as e:
            print(f"Error saving to network: {e}")
            return False

    def open_workbook(self, path):
        """Opens a workbook with special handling for network paths"""
        try:
            if not self.app:
                self.app = xw.App(visible=True)

            # Close any open workbook first
            if self.wb:
                self.close_workbook()

            self.original_path = path

            # Handle network path
            if self._is_network_path(path):
                self.temp_path = self._create_temp_copy(path)
                if self.temp_path:
                    self.wb = self.app.books.open(self.temp_path)
                else:
                    raise Exception("Failed to create temporary copy")
            else:
                self.wb = self.app.books.open(path)

            self.is_dirty = False
            return True
        except Exception as e:
            print(f"Error opening workbook: {e}")
            return False

    def save_workbook(self):
        """Saves the workbook with special handling for network paths"""
        try:
            if self.wb and self.is_dirty:
                if self._is_network_path(self.original_path):
                    return self._save_back_to_network()
                else:
                    self.wb.save()
                    self.is_dirty = False
                return True
        except Exception as e:
            print(f"Error saving workbook: {e}")
            return False

    def close_workbook(self):
        """Closes the workbook with network path handling"""
        try:
            if self.wb:
                if self.is_dirty:
                    self.save_workbook()
                if not self._is_network_path(self.original_path):
                    self.wb.close()
                self.wb = None

                # Clean up temp file if it exists
                if self.temp_path and os.path.exists(self.temp_path):
                    try:
                        os.remove(self.temp_path)
                    except:
                        pass
                self.temp_path = None
                return True
        except Exception as e:
            print(f"Error closing workbook: {e}")
            return False

    def cleanup(self):
        """Cleans up all Excel resources"""
        try:
            if self.wb:
                self.close_workbook()
            if self.app:
                self.app.quit()
                self.app = None

            # Final cleanup of temp file if it still exists
            if self.temp_path and os.path.exists(self.temp_path):
                try:
                    os.remove(self.temp_path)
                except:
                    pass
        except Exception as e:
            print(f"Error during cleanup: {e}")

    def mark_as_modified(self):
        """Marks the workbook as having unsaved changes"""
        self.is_dirty = True