# docx_properties

`docx_properties` is a Python library designed for reading, editing, and managing custom and core metadata properties in Microsoft Word `.docx` documents. It allows developers to programmatically interact with properties such as title, subject, author, and custom-defined metadata fields.

## Features

- **Custom Properties**:
  - Read, update, and filter custom properties in Word documents.
  - Add new custom properties or modify existing ones.
- **Core Properties**:
  - Access core document properties like title, subject, creator, and revision number.
- **Field Updates**:
  - Automatically update fields in Word documents (Windows only, requires `pywin32`).
- **Command-Line Interface (CLI)**:
  - Quickly manage document properties without writing Python code.

## Installation

Install the library using pip:

```bash
pip install docx_properties
```

For Windows users who require field updating functionality:
```bash
pip install pywin32
```

## Usage

### Python API

The library provides an easy-to-use API for interacting with Word document properties.

#### Example: Setting a Custom Property
You can use the `set_custom_property` method to add or update custom properties in a Word document.

```python
from docx_properties import DocxProperties

# Open the Word document
doc = DocxProperties("example.docx")

# Set or update a custom property
property_name = "ApprovalStatus"
property_value = "Approved"

# Add the custom property to the document
success = doc.set_custom_property(property_name, property_value)

if success:
    print(f"The custom property '{property_name}' was successfully set to '{property_value}'.")
else:
    print(f"Failed to set the custom property '{property_name}'.")
```

#### Example: Setting Multiple Properties
You can set multiple properties by calling `set_custom_property` multiple times.

```python
# Set multiple custom properties
doc.set_custom_property("ApprovalStatus", "Approved")
doc.set_custom_property("Version", 2.1)
doc.set_custom_property("ReviewedBy", "John Doe")
```

#### Example: Boolean and Numeric Properties
The library supports various data types for custom properties, such as strings, integers, floats, and booleans.

```python
# Boolean property
doc.set_custom_property("IsConfidential", True)

# Numeric property
doc.set_custom_property("DocumentID", 12345)
```

---

### CLI Usage

The library also includes a command-line interface (CLI) for quickly managing properties.

#### Setting a Custom Property via CLI
```bash
docxprop set-custom example.docx ApprovalStatus Approved
```

---

## Field Updates (Windows Only)

If you need to update fields in your Word document (e.g., updating a table of contents or other dynamic fields), use the `update_fields` method. This functionality requires `pywin32` and is only available on Windows.

```python
# Update fields in the document
if doc.update_fields():
    print("Fields were successfully updated.")
else:
    print("Field updating failed. Ensure pywin32 is installed.")
```

---

## Advanced Features

### Filtering Properties
You can filter custom properties by type and value.

```python
# Get all boolean custom properties with the value False
false_properties = doc.filter_custom_properties(filter_type="boolean", filter_value=False)

for name, prop in false_properties.items():
    print(f"Property: {name}, Value: {prop.value}")
```

### Reading Core Properties
Access core metadata such as the document title, creator, and modification date.

```python
core_properties = doc.get_core_properties()
for name, prop in core_properties.items():
    print(f"Core Property: {name}, Value: {prop.value} (Type: {prop.type})")
```

---

## License

This project is licensed under the MIT License.