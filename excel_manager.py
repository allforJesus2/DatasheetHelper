import xlwings as xw
import os
import shutil
import tempfile
from pathlib import Path


class ExcelManager:
    """
    A class to manage Excel workbook interactions with special handling for network paths
    Supports multiple open workbooks, one per destination path
    """

    def __init__(self):
        self.wb = None  # Current active workbook
        self.app = None  # Current active Excel app
        self.is_dirty = False
        self.original_path = None  # Path of current active workbook
        self.temp_path = None
        
        # Dictionary to track multiple open workbooks: {normalized_path: {'wb': workbook, 'app': app, 'is_dirty': bool, 'original_path': str, 'temp_path': str}}
        self.open_workbooks = {}

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

                # Remove from cache if it exists (network workbooks are closed after save)
                normalized_path = self._normalize_path(self.original_path)
                if normalized_path in self.open_workbooks:
                    del self.open_workbooks[normalized_path]

                self.temp_path = None
                self.is_dirty = False
                return True
        except Exception as e:
            print(f"Error saving to network: {e}")
            return False

    def _normalize_path(self, path):
        """Normalize path for consistent dictionary keys"""
        return os.path.normpath(os.path.abspath(path)).lower()
    
    def open_workbook(self, path):
        """Opens a workbook with special handling for network paths. 
        If workbook is already open, switches to it instead of closing."""
        print(f"DEBUG: ExcelManager.open_workbook starting for path: {path}")
        try:
            normalized_path = self._normalize_path(path)
            
            # Check if this workbook is already open
            if normalized_path in self.open_workbooks:
                workbook_info = self.open_workbooks[normalized_path]
                # Verify the workbook is still valid
                try:
                    _ = workbook_info['wb'].name  # Test if workbook is still accessible
                    print(f"DEBUG: Workbook already open for path: {path}, switching to it")
                    # Switch to this workbook
                    self.wb = workbook_info['wb']
                    self.app = workbook_info['app']
                    self.is_dirty = workbook_info['is_dirty']
                    self.original_path = workbook_info['original_path']
                    self.temp_path = workbook_info['temp_path']
                    print("DEBUG: Switched to existing workbook")
                    return
                except Exception as e:
                    print(f"DEBUG: Previously open workbook is no longer valid: {e}, removing from cache")
                    # Workbook was closed externally, remove from cache
                    del self.open_workbooks[normalized_path]
            
            # Save current workbook state before switching (if any)
            if self.wb and self.original_path:
                current_normalized = self._normalize_path(self.original_path)
                if current_normalized not in self.open_workbooks:
                    # Save current workbook state to cache
                    self.open_workbooks[current_normalized] = {
                        'wb': self.wb,
                        'app': self.app,
                        'is_dirty': self.is_dirty,
                        'original_path': self.original_path,
                        'temp_path': self.temp_path
                    }
                    print(f"DEBUG: Cached current workbook state for: {self.original_path}")
            
            # Create new Excel instance without a blank workbook
            print("DEBUG: ExcelManager creating new Excel app...")
            self.app = xw.App(visible=True, add_book=False)
            self.app.display_alerts = False
            print("DEBUG: ExcelManager Excel app created")
            
            # Open workbook
            print("DEBUG: ExcelManager opening workbook...")
            self.wb = self.app.books.open(path)
            self.is_dirty = False
            self.original_path = path
            
            # Store in cache
            self.open_workbooks[normalized_path] = {
                'wb': self.wb,
                'app': self.app,
                'is_dirty': self.is_dirty,
                'original_path': self.original_path,
                'temp_path': self.temp_path
            }
            print("DEBUG: ExcelManager workbook opened successfully and cached")
            
        except Exception as e:
            print(f"DEBUG: ExcelManager error opening workbook: {e}")
            # Don't cleanup here - might have other workbooks open
            raise Exception(f"Failed to open workbook: {e}")

    def save_workbook(self):
        """Saves the workbook with special handling for network paths"""
        try:
            if self.wb and self.is_dirty:
                if self._is_network_path(self.original_path):
                    result = self._save_back_to_network()
                    # Update cache after save
                    if self.original_path:
                        normalized_path = self._normalize_path(self.original_path)
                        if normalized_path in self.open_workbooks:
                            self.open_workbooks[normalized_path]['is_dirty'] = False
                    return result
                else:
                    self.wb.save()
                    self.is_dirty = False
                    # Update cache
                    if self.original_path:
                        normalized_path = self._normalize_path(self.original_path)
                        if normalized_path in self.open_workbooks:
                            self.open_workbooks[normalized_path]['is_dirty'] = False
                return True
        except Exception as e:
            print(f"Error saving workbook: {e}")
            return False

    def close_workbook(self, path=None):
        """Closes a specific workbook (or current if path not specified) with network path handling.
        Removes it from the cache but keeps other workbooks open."""
        try:
            if path:
                # Close a specific workbook by path
                normalized_path = self._normalize_path(path)
                if normalized_path in self.open_workbooks:
                    workbook_info = self.open_workbooks[normalized_path]
                    wb = workbook_info['wb']
                    original_path = workbook_info['original_path']
                    temp_path = workbook_info['temp_path']
                    
                    # Don't save automatically - let user save manually in Excel
                    # Close the workbook
                    if not self._is_network_path(original_path):
                        try:
                            wb.close()
                        except:
                            pass
                    
                    # Clean up temp file if it exists
                    if temp_path and os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except:
                            pass
                    
                    # Remove from cache
                    del self.open_workbooks[normalized_path]
                    
                    # If this was the current workbook, clear current state
                    if self.wb == wb:
                        self.wb = None
                        self.app = None
                        self.is_dirty = False
                        self.original_path = None
                        self.temp_path = None
                    
                    return True
            else:
                # Close current workbook
                if self.wb:
                    # Don't save automatically - let user save manually in Excel
                    if self.original_path and not self._is_network_path(self.original_path):
                        self.wb.close()
                    
                    # Remove from cache
                    if self.original_path:
                        normalized_path = self._normalize_path(self.original_path)
                        if normalized_path in self.open_workbooks:
                            del self.open_workbooks[normalized_path]
                    
                    # Clean up temp file if it exists
                    if self.temp_path and os.path.exists(self.temp_path):
                        try:
                            os.remove(self.temp_path)
                        except:
                            pass
                    
                    self.wb = None
                    self.app = None
                    self.is_dirty = False
                    self.original_path = None
                    self.temp_path = None
                    return True
            return False
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

    def cleanup(self, close_all=False):
        """Clean up Excel resources. 
        If close_all=True, closes all open workbooks. 
        Otherwise, just clears current references without closing workbooks."""
        print("DEBUG: ExcelManager cleanup starting...")
        if close_all:
            # Close all open workbooks
            paths_to_close = list(self.open_workbooks.keys())
            for normalized_path in paths_to_close:
                workbook_info = self.open_workbooks[normalized_path]
                original_path = workbook_info['original_path']
                try:
                    # Don't save automatically - let user save manually in Excel
                    if not self._is_network_path(original_path):
                        workbook_info['wb'].close()
                    
                    # Clean up temp file
                    if workbook_info['temp_path'] and os.path.exists(workbook_info['temp_path']):
                        try:
                            os.remove(workbook_info['temp_path'])
                        except:
                            pass
                except Exception as e:
                    print(f"DEBUG: Error closing workbook {original_path}: {e}")
            
            # Quit all Excel apps (they should all be closed now)
            apps_to_quit = set()
            for workbook_info in self.open_workbooks.values():
                if workbook_info['app']:
                    apps_to_quit.add(workbook_info['app'])
            
            for app in apps_to_quit:
                try:
                    app.quit()
                except:
                    pass
            
            self.open_workbooks.clear()
        
        # Clear current references
        try:
            self.wb = None 
            self.app = None
            self.is_dirty = False
            self.original_path = None
            self.temp_path = None
        except Exception as e:
            print(f"DEBUG: ExcelManager cleanup error: {e}")
            pass
        print("DEBUG: ExcelManager cleanup completed")

    def mark_as_modified(self):
        """Marks the workbook as having unsaved changes"""
        self.is_dirty = True
        # Also update the cached workbook info
        if self.original_path:
            normalized_path = self._normalize_path(self.original_path)
            if normalized_path in self.open_workbooks:
                self.open_workbooks[normalized_path]['is_dirty'] = True