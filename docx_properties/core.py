import zipfile
from lxml import etree
import os
import tempfile
import shutil
from typing import Dict, Any, Optional, Union, Tuple

class PropertyValue:
    """
    Class that represents the value and type of a document property.
    """
    
    def __init__(self, value: Any, type_name: str):
        """
        Initializes a PropertyValue object.
        
        Args:
            value: The value of the property
            type_name: The type of the property as a string (bool, str, int, float, datetime)
        """
        self._value = value
        self._type = type_name
    
    @property
    def value(self) -> Any:
        """Returns the value of the property."""
        return self._value
        
    @property
    def type(self) -> str:
        """Returns the type of the property as a string."""
        return self._type
    
    def __repr__(self) -> str:
        """String representation for debugging purposes."""
        return f"PropertyValue(value={repr(self._value)}, type={repr(self._type)})"
    
    def __str__(self) -> str:
        """String representation of the property."""
        return str(self._value)

class DocxProperties:
    """
    Class for reading and manipulating properties in Word documents.
    Supports both custom and core properties.
    """
    
    def __init__(self, docx_path: str):
        """
        Initializes a DocxProperties object for a Word file.
        
        Args:
            docx_path: Path to the Word file (.docx)
        """
        self.docx_path = os.path.abspath(docx_path)
        if not os.path.exists(self.docx_path):
            raise FileNotFoundError(f"The file {self.docx_path} was not found.")
    
    def get_properties(self) -> Dict[str, PropertyValue]:
        """
        Reads all custom properties from the Word document.
        DEPRECATED: Use get_custom_properties() instead.
        
        Returns:
            Dictionary with all custom properties
        """
        return self.get_custom_properties()
    
    def get_custom_properties(self) -> Dict[str, PropertyValue]:
        """
        Reads all custom properties from the Word document.
        
        Returns:
            Dictionary with all custom properties as PropertyValue objects
        """
        try:
            # Open .docx file as ZIP archive
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                # Check if the custom properties file exists
                custom_props_path = 'docProps/custom.xml'
                
                # If no custom properties are present, return empty dict
                if custom_props_path not in zip_ref.namelist():
                    return {}
                
                # Read XML file with custom properties
                with zip_ref.open(custom_props_path) as f:
                    xml_content = f.read()
                    
                # Parse XML with lxml
                root = etree.fromstring(xml_content)
                
                # Define namespaces
                namespaces = {
                    'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
                    'cp': 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
                }
                
                # Extract all properties
                properties = {}
                for prop in root.xpath('//cp:property', namespaces=namespaces):
                    name = prop.get('name')
                    
                    # The first child element contains the value and type
                    for child in prop:
                        tag = etree.QName(child).localname
                        if tag == 'lpwstr':  # String value
                            properties[name] = PropertyValue(child.text, "string")
                        elif tag == 'i4':  # Integer value
                            properties[name] = PropertyValue(int(child.text), "integer")
                        elif tag == 'bool':  # Boolean value
                            properties[name] = PropertyValue(child.text.lower() == 'true', "boolean")
                        elif tag == 'filetime':  # DateTime value
                            properties[name] = PropertyValue(child.text, "datetime")
                        elif tag == 'r8':  # Double value
                            properties[name] = PropertyValue(float(child.text), "float")
                        else:
                            properties[name] = PropertyValue(child.text, tag)
                        break
                        
                return properties
                
        except zipfile.BadZipFile:
            raise ValueError(f"{self.docx_path} is not a valid Word file")
        except Exception as e:
            raise Exception(f"Error extracting custom properties: {str(e)}")
    
    def get_core_properties(self) -> Dict[str, PropertyValue]:
        """
        Reads all core properties from the Word document.
        
        Returns:
            Dictionary with all core properties as PropertyValue objects
        """
        try:
            # Open .docx file as ZIP archive
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                # Check if the core properties file exists
                core_props_path = 'docProps/core.xml'
                
                # If no core properties are present, return empty dict
                if core_props_path not in zip_ref.namelist():
                    return {}
                
                # Read XML file with core properties
                with zip_ref.open(core_props_path) as f:
                    xml_content = f.read()
                    
                # Parse XML with lxml
                root = etree.fromstring(xml_content)
                
                # Define namespaces
                namespaces = {
                    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                    'dc': 'http://purl.org/dc/elements/1.1/',
                    'dcterms': 'http://purl.org/dc/terms/',
                    'dcmitype': 'http://purl.org/dc/dcmitype/',
                    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                }
                
                # Extract properties from different namespaces
                properties = {}
                
                # DC elements (Dublin Core) - all as string type
                for element_name in ['title', 'creator', 'subject', 'description', 'language']:
                    elements = root.xpath(f'//dc:{element_name}', namespaces=namespaces)
                    if elements and elements[0].text:
                        properties[element_name] = PropertyValue(elements[0].text, "string")
                
                # DCTERMS elements
                for element_name in ['created', 'modified']:
                    elements = root.xpath(f'//dcterms:{element_name}', namespaces=namespaces)
                    if elements and elements[0].text:
                        properties[element_name] = PropertyValue(elements[0].text, "datetime")
                
                # CP elements (Core Properties)
                for element_name in ['lastModifiedBy', 'category', 'contentStatus', 'keywords']:
                    elements = root.xpath(f'//cp:{element_name}', namespaces=namespaces)
                    if elements and elements[0].text:
                        properties[element_name] = PropertyValue(elements[0].text, "string")
                
                # Revision is usually an Integer
                elements = root.xpath(f'//cp:revision', namespaces=namespaces)
                if elements and elements[0].text:
                    try:
                        properties['revision'] = PropertyValue(int(elements[0].text), "integer")
                    except ValueError:
                        properties['revision'] = PropertyValue(elements[0].text, "string")
                
                return properties
                
        except zipfile.BadZipFile:
            raise ValueError(f"{self.docx_path} is not a valid Word file")
        except Exception as e:
            raise Exception(f"Error extracting core properties: {str(e)}")
    
    def get_all_properties(self) -> Dict[str, PropertyValue]:
        """
        Reads all properties (core and custom) from the Word document.
        
        Returns:
            Dictionary with all properties as PropertyValue objects
        """
        all_properties = {}
        
        # Add core properties
        core_properties = self.get_core_properties()
        for key, value in core_properties.items():
            all_properties[f"core:{key}"] = value
        
        # Add custom properties
        custom_properties = self.get_custom_properties()
        for key, value in custom_properties.items():
            all_properties[f"custom:{key}"] = value
            
        return all_properties
    
    def filter_properties(self, property_dict, filter_type=None, filter_value=None):
        """
        Filters properties by type and/or value.
        
        Args:
            property_dict: Dictionary with PropertyValue objects
            filter_type: Optional, filters by this type (e.g. 'boolean', 'string')
            filter_value: Optional, filters by this value
            
        Returns:
            Dictionary with filtered properties
        """
        result = {}
        
        for name, prop in property_dict.items():
            # Apply type filter if present
            if filter_type is not None and prop.type != filter_type:
                continue
                
            # Apply value filter if present
            if filter_value is not None and prop.value != filter_value:
                continue
                
            # If all filters passed, include in result
            result[name] = prop
            
        return result

    def filter_custom_properties(self, filter_type=None, filter_value=None):
        """
        Filters custom properties by type and/or value.
        
        Args:
            filter_type: Optional, filters by this type (e.g. 'boolean', 'string')
            filter_value: Optional, filters by this value
            
        Returns:
            Dictionary with filtered custom properties
        """
        custom_props = self.get_custom_properties()
        return self.filter_properties(custom_props, filter_type, filter_value)

    def filter_core_properties(self, filter_type=None, filter_value=None):
        """
        Filters core properties by type and/or value.
        
        Args:
            filter_type: Optional, filters by this type (e.g. 'datetime', 'string')
            filter_value: Optional, filters by this value
            
        Returns:
            Dictionary with filtered core properties
        """
        core_props = self.get_core_properties()
        return self.filter_properties(core_props, filter_type, filter_value)

    def filter_all_properties(self, filter_type=None, filter_value=None):
        """
        Filters all properties by type and/or value.
        
        Args:
            filter_type: Optional, filters by this type (e.g. 'boolean', 'string')
            filter_value: Optional, filters by this value
            
        Returns:
            Dictionary with filtered properties
        """
        all_props = self.get_all_properties()
        return self.filter_properties(all_props, filter_type, filter_value)
    
    def get_property(self, property_name: str) -> Optional[PropertyValue]:
        """
        Reads a custom property from the Word document.
        DEPRECATED: Use get_custom_property() instead.
        
        Args:
            property_name: Name of the property
            
        Returns:
            PropertyValue object of the property or None if the property doesn't exist
        """
        return self.get_custom_property(property_name)
    
    def get_custom_property(self, property_name: str) -> Optional[PropertyValue]:
        """
        Reads a specific custom property from the Word document.
        
        Args:
            property_name: Name of the custom property
            
        Returns:
            PropertyValue object of the property or None if the property doesn't exist
        """
        properties = self.get_custom_properties()
        return properties.get(property_name)
    
    def get_core_property(self, property_name: str) -> Optional[PropertyValue]:
        """
        Reads a specific core property from the Word document.
        
        Args:
            property_name: Name of the core property
            
        Returns:
            PropertyValue object of the property or None if the property doesn't exist
        """
        properties = self.get_core_properties()
        return properties.get(property_name)
    
    def set_property(self, property_name: str, property_value: Any, update_fields: bool = True) -> bool:
        """
        Sets a custom property in the Word document.
        DEPRECATED: Use set_custom_property() instead.
        
        Args:
            property_name: Name of the property
            property_value: Value for the property
            update_fields: Whether to update fields in the document (requires pywin32)
            
        Returns:
            True on success, False on error
        """
        return self.set_custom_property(property_name, property_value, update_fields)
    
    def set_custom_property(self, property_name: str, property_value: Any, update_fields: bool = True) -> bool:
        """
        Sets a custom property in the Word document.
        
        Args:
            property_name: Name of the custom property
            property_value: Value for the property
            update_fields: Whether to update fields in the document (requires pywin32)
            
        Returns:
            True on success, False on error
        """
        # If property_value is a PropertyValue object, extract only the value
        if isinstance(property_value, PropertyValue):
            property_value = property_value.value
            
        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_dir, "temp.docx")
        
        try:
            # Extract Word document as ZIP archive to temporary directory
            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Path to custom.xml in the temporary directory
            custom_props_path = os.path.join(temp_dir, 'docProps', 'custom.xml')
            custom_props_dir = os.path.dirname(custom_props_path)
            
            # Create directory if it doesn't exist
            if not os.path.exists(custom_props_dir):
                os.makedirs(custom_props_dir)
            
            # Check if custom.xml exists
            if os.path.exists(custom_props_path):
                tree = etree.parse(custom_props_path)
                root = tree.getroot()
            else:
                root = etree.Element("{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}Properties")
                root.set("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
            
            # Define namespaces
            namespaces = {
                'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
                'cp': 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
            }
            
            # Look for the property and maximum PID
            max_pid = 1
            property_element = None
            
            for prop in root.xpath('//cp:property', namespaces=namespaces):
                pid = int(prop.get('pid', '0'))
                max_pid = max(max_pid, pid)
                
                if prop.get('name') == property_name:
                    property_element = prop
            
            # Create property if it doesn't exist
            if property_element is None:
                property_element = etree.SubElement(root, "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property")
                property_element.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
                property_element.set("pid", str(max_pid + 1))
                property_element.set("name", property_name)
            
            # Remove all child elements
            for child in property_element:
                property_element.remove(child)
            
            # Set the correct type for the value
            if isinstance(property_value, bool):
                value_element = etree.SubElement(property_element, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}bool")
                value_element.text = "true" if property_value else "false"
            elif isinstance(property_value, int):
                value_element = etree.SubElement(property_element, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}i4")
                value_element.text = str(property_value)
            elif isinstance(property_value, float):
                value_element = etree.SubElement(property_element, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}r8")
                value_element.text = str(property_value)
            else:
                value_element = etree.SubElement(property_element, "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr")
                value_element.text = str(property_value)
            
            # Save XML
            tree = etree.ElementTree(root)
            tree.write(custom_props_path, xml_declaration=True, encoding="UTF-8")
            
            # Create new Word file
            with zipfile.ZipFile(temp_docx, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                for folder_name, _, filenames in os.walk(temp_dir):
                    for filename in filenames:
                        if filename == "temp.docx":
                            continue
                        file_path = os.path.join(folder_name, filename)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        new_zip.write(file_path, arc_name)
            
            # Replace the original file
            shutil.move(temp_docx, self.docx_path)
            
            # Update fields if desired
            if update_fields:
                self.update_fields()
                
            return True
            
        except Exception as e:
            return False
        finally:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
    
    def update_fields(self) -> bool:
        """
        Updates all fields in the Word document.
        
        Returns:
            True on success, False on error or when pywin32 is not available
        """
        try:
            import win32com.client
            
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                doc = word.Documents.Open(self.docx_path)
                word.Selection.WholeStory()
                word.Selection.Fields.Update()
                doc.Save()
                doc.Close()
                return True
            finally:
                word.Quit()
        except ImportError:
            return False
        except Exception:
            return False