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
        self.wb = None
        self.app = None
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
            # Force close any existing connections
            self.cleanup()
            
            # Create new Excel instance without a blank workbook
            self.app = xw.App(visible=True, add_book=False)
            self.app.display_alerts = False
            
            # Open workbook
            self.wb = self.app.books.open(path)
            self.is_dirty = False
            
        except Exception as e:
            self.cleanup()
            raise Exception(f"Failed to open workbook: {e}")

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

    def release_connection(self):
        """Release xlwings connection but keep workbook open in Excel"""
        try:
            # Don't close the workbook - just release our connection to it
            # The workbook will remain open in Excel for the user to save
            
            # Clear our references to release the xlwings connection
            self.wb = None 
            self.app = None
            self.is_dirty = False
            
            # Force garbage collection to ensure xlwings objects are cleaned up
            import gc
            collected = gc.collect()
            print(f"Garbage collector collected {collected} objects")
            return True
        except Exception as e:
            print(f"Error releasing connection: {e}")
            return False

    def cleanup(self):
        """Clean up Excel resources"""
        try:
            if self.wb:
                self.wb.close()
            if self.app:
                self.app.quit()
        except:
            pass
        finally:
            self.wb = None 
            self.app = None
            self.is_dirty = False

    def mark_as_modified(self):
        """Marks the workbook as having unsaved changes"""
        self.is_dirty = True