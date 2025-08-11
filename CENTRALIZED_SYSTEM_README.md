# Centralized Data Type System

## Overview

The DatasheetHelper application has been refactored to use a centralized data type system that eliminates redundancy between TD (Tag Dictionary) and PC (Process Conditions) implementations. This system allows for easy addition of new data types and provides a consistent interface for managing different types of data.

## Key Features

### 1. Centralized Configuration
All data types are now defined in one place in the `__init__` method:

```python
self.data_types = {
    'td': {
        'name': 'Tag Dictionary',
        'short_name': 'TD',
        'description': 'Tag data from Excel files',
        'default_headers': ['TAG NUMBER'],
        'coordinate_values': {},
        'selected_sheets': None,
        'path_key': 'tag_data_path'
    },
    'pc': {
        'name': 'Process Conditions', 
        'short_name': 'PC',
        'description': 'Process conditions data from Excel files',
        'default_headers': ['Line No.'],
        'coordinate_values': {},
        'selected_sheets': None,
        'path_key': 'process_conditions_path'
    }
}
```

### 2. Helper Methods
The system provides consistent helper methods for all data types:

- `get_data_type_config(data_type)` - Get configuration for a data type
- `get_data_type_name(data_type)` - Get full name of a data type
- `get_data_type_short_name(data_type)` - Get short name of a data type
- `get_data_type_headers(data_type)` - Get headers for a data type
- `set_data_type_headers(data_type, headers)` - Set headers for a data type
- `get_data_type_data(data_type)` - Get data dictionary for a data type
- `set_data_type_data(data_type, data)` - Set data dictionary for a data type
- `get_data_type_coordinate_values(data_type)` - Get coordinate values for a data type
- `set_data_type_coordinate_values(data_type, coordinate_values)` - Set coordinate values for a data type
- `get_data_type_selected_sheets(data_type)` - Get selected sheets for a data type
- `set_data_type_selected_sheets(data_type, selected_sheets)` - Set selected sheets for a data type
- `get_data_type_path(data_type)` - Get file path for a data type
- `set_data_type_path(data_type, path)` - Set file path for a data type
- `get_all_data_types()` - Get all available data types
- `add_data_type(data_type, config)` - Add a new data type configuration

### 3. Automatic UI Generation
The entire GUI now automatically generates UI elements for all data types:
- **Main GUI**: Automatically creates frames with Browse, Configure, Generate, and View buttons for each data type
- **Visual Organization**: Clear separation between "DATA SOURCES" (green label) and "DESTINATION" (blue label, bordered container)
- **Coordinates Tab**: Dynamically generates UI elements for all data types instead of hardcoding TD and PC frames
- **Real-time Updates**: When new data types are added, the GUI automatically refreshes to include them
- **Consistent Interface**: All data types get the same functionality and appearance
- **Single Destination**: Only one datasheet destination is shown, clearly separated from data sources

### 4. Easy Extension
Adding new data types is now simple and requires minimal code changes. The GUI automatically adapts to any number of data types.

## Benefits

1. **Reduced Redundancy**: No more duplicate code for TD and PC operations
2. **Consistency**: All data types use the same interface and patterns
3. **Maintainability**: Changes to data type handling only need to be made in one place
4. **Extensibility**: New data types can be added without modifying existing code
5. **Type Safety**: Centralized configuration reduces the chance of errors

## Usage Examples

### Adding a New Data Type

```python
# Define the configuration
equipment_config = {
    'name': 'Equipment Data',
    'short_name': 'EQ',
    'description': 'Equipment information from Excel files',
    'default_headers': ['Equipment ID'],
    'coordinate_values': {},
    'selected_sheets': None,
    'path_key': 'equipment_path'
}

# Add the data type
app.add_data_type('equipment', equipment_config)

# Add the path to parameters
app.parameters['equipment_path'] = ''
setattr(app, 'equipment_path', '')
```

### Working with Data Types

```python
# Get data for a specific type
td_data = app.get_data_type_data('td')
pc_data = app.get_data_type_data('pc')

# Set data for a specific type
app.set_data_type_data('td', new_td_data)

# Get coordinate values
td_coords = app.get_data_type_coordinate_values('td')

# Set coordinate values
app.set_data_type_coordinate_values('td', new_coordinate_values)

# Get configuration
config = app.get_data_type_config('td')
name = app.get_data_type_name('td')
short_name = app.get_data_type_short_name('td')
```

### Iterating Over All Data Types

```python
for data_type in app.get_all_data_types():
    data = app.get_data_type_data(data_type)
    config = app.get_data_type_config(data_type)
    print(f"Processing {config['name']}...")
```

## Testing

Run the test script to see the system in action:

```bash
python test_centralized_system.py
```

This will demonstrate:
- Initial data type configuration
- Adding new data types
- Data manipulation
- Coordinate value management

## Migration Notes

The refactoring maintains backward compatibility while providing the new centralized system. All existing functionality should work as before, but now with the benefits of the centralized approach.

## UI Improvements

The GUI has been enhanced with better visual organization:

### Visual Separation
- **Data Sources Section**: All input data types are grouped under a green "DATA SOURCES" label
- **Destination Section**: The single output destination is clearly marked with a blue "DESTINATION" label in a bordered container
- **Visual Separator**: A gray line separates the sources from the destination

### Benefits
- **Clear Workflow**: Users can easily distinguish between input sources and output destination
- **No Duplication**: Only one datasheet destination is shown, eliminating confusion
- **Scalable Design**: New data sources are automatically added to the sources section
- **Professional Appearance**: The bordered destination container makes the output clear

## Future Enhancements

With this centralized system, future enhancements could include:
- Data type validation
- Data type-specific processing rules
- Plugin system for custom data types
- Data type relationships and dependencies
- Custom UI themes and styling options 