#!/usr/bin/env python3
"""
Excel Manager - Handles Excel file operations for the dynamic datasheet system
"""

import os
import json
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class ExcelManager:
    """Manages Excel file operations for the dynamic system"""
    
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xls']
        
    def get_sheet_names(self, filepath: str) -> List[str]:
        """Get all sheet names from an Excel file"""
        try:
            if not os.path.exists(filepath):
                print(f"File not found: {filepath}")
                return []
                
            workbook = openpyxl.load_workbook(filepath, read_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
            return sheet_names
        except Exception as e:
            print(f"Error reading Excel file {filepath}: {e}")
            return []
            
    def read_excel_sheet(self, filepath: str, sheet_name: str = None) -> Optional[pd.DataFrame]:
        """Read a specific sheet from an Excel file"""
        try:
            if not os.path.exists(filepath):
                print(f"File not found: {filepath}")
                return None
                
            # Read the Excel file
            if sheet_name:
                df = pd.read_excel(filepath, sheet_name=sheet_name)
            else:
                df = pd.read_excel(filepath)
                
            return df
        except Exception as e:
            print(f"Error reading Excel sheet {sheet_name} from {filepath}: {e}")
            return None
            
    def read_multiple_sheets(self, filepath: str, sheet_names: List[str]) -> Dict[str, pd.DataFrame]:
        """Read multiple sheets from an Excel file"""
        try:
            if not os.path.exists(filepath):
                print(f"File not found: {filepath}")
                return {}
                
            # Read all specified sheets
            all_sheets = pd.read_excel(filepath, sheet_name=sheet_names)
            
            # Convert to dictionary if only one sheet
            if len(sheet_names) == 1:
                return {sheet_names[0]: all_sheets}
            else:
                return all_sheets
                
        except Exception as e:
            print(f"Error reading multiple sheets from {filepath}: {e}")
            return {}
            
    def generate_dictionary_from_xlsx(self, filepath: str, sheet_names: List[str] = None, 
                                    key_column: str = None) -> Dict[str, Dict[str, Any]]:
        """Generate a dictionary from Excel data"""
        try:
            if not os.path.exists(filepath):
                print(f"File not found: {filepath}")
                return {}
                
            # If no sheet names specified, get all sheets
            if not sheet_names:
                sheet_names = self.get_sheet_names(filepath)
                
            if not sheet_names:
                print(f"No sheets found in {filepath}")
                return {}
                
            # Read all sheets
            all_sheets = self.read_multiple_sheets(filepath, sheet_names)
            
            result_dict = {}
            
            for sheet_name, df in all_sheets.items():
                if df is None or df.empty:
                    continue
                    
                # Convert DataFrame to dictionary
                if key_column and key_column in df.columns:
                    # Use specified key column
                    df_dict = df.set_index(key_column).to_dict('index')
                    # Convert numpy types to Python types
                    for key, value in df_dict.items():
                        df_dict[key] = {k: self._convert_numpy_types(v) for k, v in value.items()}
                    result_dict.update(df_dict)
                else:
                    # Use first column as key
                    if len(df.columns) > 0:
                        key_col = df.columns[0]
                        df_dict = df.set_index(key_col).to_dict('index')
                        # Convert numpy types to Python types
                        for key, value in df_dict.items():
                            df_dict[key] = {k: self._convert_numpy_types(v) for k, v in value.items()}
                        result_dict.update(df_dict)
                        
            return result_dict
            
        except Exception as e:
            print(f"Error generating dictionary from {filepath}: {e}")
            return {}
            
    def _convert_numpy_types(self, value):
        """Convert numpy types to Python native types"""
        import numpy as np
        
        if pd.isna(value):
            return None
        elif isinstance(value, np.integer):
            return int(value)
        elif isinstance(value, np.floating):
            return float(value)
        elif isinstance(value, np.ndarray):
            return value.tolist()
        else:
            return value
            
    def write_dictionary_to_excel(self, data: Dict[str, Dict[str, Any]], filepath: str, 
                                sheet_name: str = "Data") -> bool:
        """Write a dictionary to an Excel file"""
        try:
            if not data:
                print("No data to write")
                return False
                
            # Convert dictionary to DataFrame
            df = pd.DataFrame.from_dict(data, orient='index')
            
            # Create workbook and write data
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = sheet_name
            
            # Write DataFrame to worksheet
            for r in dataframe_to_rows(df, index=True, header=True):
                worksheet.append(r)
                
            # Save workbook
            workbook.save(filepath)
            workbook.close()
            
            print(f"Data written to {filepath}")
            return True
            
        except Exception as e:
            print(f"Error writing to Excel file {filepath}: {e}")
            return False
            
    def write_coordinate_data_to_excel(self, coordinate_data: Dict[str, Dict[str, Any]], 
                                     filepath: str, sheet_name: str = "Coordinates") -> bool:
        """Write coordinate data to an Excel file"""
        try:
            if not coordinate_data:
                print("No coordinate data to write")
                return False
                
            # Convert coordinate data to DataFrame
            df = pd.DataFrame.from_dict(coordinate_data, orient='index')
            
            # Create workbook and write data
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = sheet_name
            
            # Write DataFrame to worksheet
            for r in dataframe_to_rows(df, index=True, header=True):
                worksheet.append(r)
                
            # Save workbook
            workbook.save(filepath)
            workbook.close()
            
            print(f"Coordinate data written to {filepath}")
            return True
            
        except Exception as e:
            print(f"Error writing coordinate data to Excel file {filepath}: {e}")
            return False
            
    def validate_excel_file(self, filepath: str) -> Tuple[bool, List[str]]:
        """Validate an Excel file and return issues"""
        issues = []
        
        if not os.path.exists(filepath):
            issues.append(f"File does not exist: {filepath}")
            return False, issues
            
        if not any(filepath.lower().endswith(ext) for ext in self.supported_extensions):
            issues.append(f"Unsupported file format. Supported: {', '.join(self.supported_extensions)}")
            return False, issues
            
        try:
            sheet_names = self.get_sheet_names(filepath)
            if not sheet_names:
                issues.append("No sheets found in Excel file")
                return False, issues
                
            # Check if sheets have data
            for sheet_name in sheet_names:
                df = self.read_excel_sheet(filepath, sheet_name)
                if df is None or df.empty:
                    issues.append(f"Sheet '{sheet_name}' is empty or could not be read")
                    
        except Exception as e:
            issues.append(f"Error validating Excel file: {e}")
            return False, issues
            
        return len(issues) == 0, issues
        
    def get_file_info(self, filepath: str) -> Dict[str, Any]:
        """Get information about an Excel file"""
        info = {
            'filepath': filepath,
            'exists': False,
            'size': 0,
            'sheets': [],
            'total_rows': 0,
            'total_columns': 0
        }
        
        if not os.path.exists(filepath):
            return info
            
        info['exists'] = True
        info['size'] = os.path.getsize(filepath)
        
        try:
            sheet_names = self.get_sheet_names(filepath)
            info['sheets'] = sheet_names
            
            # Get total rows and columns across all sheets
            total_rows = 0
            total_columns = 0
            
            for sheet_name in sheet_names:
                df = self.read_excel_sheet(filepath, sheet_name)
                if df is not None and not df.empty:
                    total_rows += len(df)
                    total_columns = max(total_columns, len(df.columns))
                    
            info['total_rows'] = total_rows
            info['total_columns'] = total_columns
            
        except Exception as e:
            info['error'] = str(e)
            
        return info 