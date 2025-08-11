# Dynamic Datasheet System

A fully dynamic and extensible datasheet management system that allows any data type to reference any other data type.

## 🚀 Key Features

### **Dynamic Relationship System**
- **Any data type can be the primary source** (not just TD)
- **Any data type can reference any other data type** (not just TD→PC)
- **Flexible relationships** with configurable transformations
- **No hardcoded dependencies**

### **Automatic UI Generation**
- **Dynamic GUI creation** based on data type configurations
- **Real-time updates** when adding new data types
- **Consistent interface** across all data types

### **Extensible Architecture**
- **Add new data types** without code changes
- **Configure relationships** through the GUI
- **Save/load configurations** as JSON files
- **System validation** to catch configuration issues

## 📁 Project Structure

```
dynamic_datasheet_system/
├── __init__.py                    # Package initialization
├── dynamic_relationship_system.py # Core relationship management
├── excel_manager.py              # Excel file operations
├── dynamic_datasheet_app.py      # Main GUI application
├── main.py                       # Application entry point
└── README.md                     # This file
```

## 🏗️ Architecture

### **Core Components**

1. **DynamicRelationshipSystem** - Manages data types and relationships
2. **ExcelManager** - Handles Excel file operations
3. **DynamicDatasheetApp** - Main GUI application

### **Key Concepts**

#### **Data Types**
Each data type has a configuration:
```python
{
    'name': 'Tag Dictionary',
    'short_name': 'TD',
    'description': 'Tag data from Excel files',
    'default_headers': ['TAG NUMBER', 'DESCRIPTION'],
    'coordinate_values': {'A1': 'TAG NUMBER', 'B1': 'DESCRIPTION'},
    'selected_sheets': None,
    'path_key': 'td_path'
}
```

#### **Relationships**
Relationships define how data types reference each other:
```python
{
    'source_key': 'TAG NUMBER',
    'target_key': 'Line No.',
    'transformation': 'x.split("-")[2] if "-" in x else x',
    'description': 'TD references PC via Line Number'
}
```

#### **Primary Data Type**
- **Configurable**: Any data type can be set as primary
- **Flexible**: Can be changed at runtime
- **Dynamic**: Relationships adapt automatically

## 🎯 Usage Examples

### **Traditional Setup (TD → PC)**
```python
# Set TD as primary
system.set_primary_data_type('td')

# Add relationship: TD references PC
system.add_relationship('td', 'pc', {
    'source_key': 'TAG NUMBER',
    'target_key': 'Line No.',
    'transformation': 'x.split("-")[2]'
})
```

### **Reverse Setup (PC → TD)**
```python
# Set PC as primary
system.set_primary_data_type('pc')

# Add relationship: PC references TD
system.add_relationship('pc', 'td', {
    'source_key': 'Line No.',
    'target_key': 'TAG NUMBER',
    'transformation': 'x.zfill(3)'  # Pad with zeros
})
```

### **Multiple Relationships**
```python
# TD can reference multiple data types
system.add_relationship('td', 'pc', {...})
system.add_relationship('td', 'equipment', {...})
system.add_relationship('td', 'materials', {...})
```

## 🖥️ GUI Features

### **Main Tab**
- **Data Sources**: Dynamic frames for each data type
- **Destination**: Single output file selection
- **Visual Separation**: Clear distinction between sources and destination

### **Relationships Tab**
- **Primary Data Type Selection**: Choose any data type as primary
- **Relationship Management**: View and configure relationships
- **Tree View**: Clear display of all relationships

### **Data Tab**
- **Data Type Selection**: Browse data for any type
- **Data Display**: View loaded data in formatted JSON

### **Coordinates Tab**
- **Coordinate Processing**: Process coordinates using relationships
- **Results Display**: View processed coordinate data
- **Export**: Save results to Excel

### **Settings Tab**
- **Add Data Types**: Extend the system with new data types
- **System Validation**: Check configuration for issues
- **Configuration Management**: Save/load system configurations

## 🔧 Installation & Setup

### **Requirements**
```bash
pip install pandas openpyxl tkinter
```

### **Running the Application**
```bash
# From the project root
python dynamic_datasheet_system/main.py
```

### **Using as a Package**
```python
from dynamic_datasheet_system import DynamicDatasheetApp
import tkinter as tk

root = tk.Tk()
app = DynamicDatasheetApp(root)
root.mainloop()
```

## 🔄 Migration from Original System

### **Key Differences**

| Original System | Dynamic System |
|----------------|----------------|
| Hardcoded TD/PC | Any data types |
| TD always primary | Configurable primary |
| Fixed relationships | Dynamic relationships |
| Static UI | Dynamic UI generation |
| Limited extensibility | Fully extensible |

### **Migration Steps**
1. **Load existing configuration** (if available)
2. **Set up data types** (TD, PC, etc.)
3. **Configure relationships** between data types
4. **Set primary data type** (default: TD)
5. **Load data** from Excel files
6. **Process coordinates** using dynamic relationships

## 🎨 Customization

### **Adding New Data Types**
```python
# Equipment data type
equipment_config = {
    'name': 'Equipment Data',
    'short_name': 'EQ',
    'description': 'Equipment information',
    'default_headers': ['Equipment ID', 'Type', 'Location'],
    'coordinate_values': {'E1': 'Equipment ID', 'F1': 'Type'},
    'selected_sheets': None,
    'path_key': 'equipment_path'
}
system.add_data_type('equipment', equipment_config)
```

### **Custom Transformations**
```python
# Simple transformation
'transformation': 'x.upper()'

# Complex transformation
'transformation': 'x.split("-")[1] if "-" in x else x[:3]'

# Conditional transformation
'transformation': 'x.zfill(4) if x.isdigit() else x'
```

## 🔍 System Validation

The system includes comprehensive validation:
- **Data type existence** checks
- **Relationship validity** verification
- **Coordinate mapping** validation
- **Configuration consistency** checks

## 📊 Data Flow

1. **Load Data**: Excel files → Data types
2. **Set Primary**: Choose primary data type
3. **Configure Relationships**: Define how types reference each other
4. **Process Coordinates**: Apply relationships to generate coordinate data
5. **Export Results**: Save to Excel file

## 🚀 Future Enhancements

- **Advanced transformation functions**
- **Visual relationship diagrams**
- **Data type templates**
- **Batch processing capabilities**
- **API integration**
- **Cloud storage support**

## 🤝 Contributing

This system is designed to be easily extensible. Key areas for contribution:
- **New data type templates**
- **Advanced transformation functions**
- **Enhanced GUI features**
- **Performance optimizations**
- **Documentation improvements**

## 📝 License

This project is part of the DatasheetHelper system and follows the same licensing terms. 