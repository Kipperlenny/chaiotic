"""Utility functions for XML handling in documents."""

from typing import Dict, Any, Optional
import xml.etree.ElementTree as ET

try:
    import lxml.etree as LET
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False

def create_xml_element(tag: str, attrib: Optional[Dict[str, str]] = None, 
                      text: Optional[str] = None, nsmap: Optional[Dict[str, str]] = None) -> Any:
    """Create an XML element with proper namespace handling.
    
    Args:
        tag: Element tag name (can include namespace prefix)
        attrib: Optional attributes dictionary
        text: Optional element text
        nsmap: Optional namespace mapping
        
    Returns:
        XML Element object
    """
    if LXML_AVAILABLE:
        if ':' in tag and nsmap:
            ns, tag_name = tag.split(':')
            tag = f"{{{nsmap[ns]}}}{tag_name}"
        elem = LET.Element(tag, attrib or {}, nsmap=nsmap)
    else:
        elem = ET.Element(tag, attrib or {})
    
    if text:
        elem.text = text
    return elem

def create_metadata_element(tag: str, text: str, author: str, date: str, 
                          nsmap: Optional[Dict[str, str]] = None) -> Any:
    """Create a metadata XML element with author and date.
    
    Args:
        tag: Element tag name
        text: Element content text
        author: Author name for metadata
        date: Date string for metadata
        nsmap: Optional namespace mapping
        
    Returns:
        XML Element object with metadata
    """
    elem = create_xml_element(tag, nsmap=nsmap)
    
    # Add change-info element according to ODT spec
    change_info = create_xml_element('office:change-info', nsmap=nsmap)
    
    # Add metadata elements within change-info
    creator = create_xml_element('dc:creator', text=author, nsmap=nsmap)
    date_elem = create_xml_element('dc:date', text=date, nsmap=nsmap)
    
    change_info.append(creator)
    change_info.append(date_elem)
    elem.append(change_info)
    
    # Add content as paragraph if text is provided
    if text:
        content = create_xml_element('text:p', text=text, nsmap=nsmap)
        elem.append(content)
    
    return elem

def create_tracked_change_region(change_id: str, change_type: str, text: str,
                               author: str, date: str, nsmap: Dict[str, str]) -> Any:
    """Create a tracked change region with either deletion or insertion.
    
    Args:
        change_id: Unique identifier for the change
        change_type: Type of change ('deletion' or 'insertion')
        text: Text content for the change
        author: Author of the change
        date: Date of the change
        nsmap: Namespace mapping
        
    Returns:
        XML Element containing the tracked change
    """
    # Create change region container with both text:id and xml:id attributes
    region = create_xml_element('text:changed-region', 
                              attrib={
                                  'text:id': change_id,
                                  'xml:id': change_id  # Also add the xml:id attribute
                              }, 
                              nsmap=nsmap)
    
    # Create the appropriate change element based on type (only one per region)
    if change_type == 'deletion':
        change_elem = create_xml_element('text:deletion', nsmap=nsmap)
    else:  # insertion
        change_elem = create_xml_element('text:insertion', nsmap=nsmap)
    
    # Add change-info element
    change_info = create_xml_element('office:change-info', nsmap=nsmap)
    creator = create_xml_element('dc:creator', text=author, nsmap=nsmap)
    date_elem = create_xml_element('dc:date', text=date, nsmap=nsmap)
    change_info.append(creator)
    change_info.append(date_elem)
    
    # Add the change_info to the change element
    change_elem.append(change_info)
    
    # Add text content as a separate paragraph after the change-info, as child of the change element
    # This is the correct interpretation of the ODT spec
    if text:
        content = create_xml_element('text:p', text=text, nsmap=nsmap)
        change_elem.append(content)
    
    # Add the change element to the region
    region.append(change_elem)
    
    return region

def parse_xml_file(file_path: str, parser: Any = None) -> Any:
    """Parse an XML file with proper error handling.
    
    Args:
        file_path: Path to the XML file
        parser: Optional custom parser
        
    Returns:
        Parsed XML tree
    """
    try:
        if LXML_AVAILABLE and not parser:
            parser = LET.XMLParser(remove_blank_text=True)
            return LET.parse(file_path, parser)
        else:
            return ET.parse(file_path, parser)
    except Exception as e:
        print(f"Error parsing XML file {file_path}: {e}")
        raise

def write_xml_file(tree: Any, file_path: str, encoding: str = 'UTF-8', 
                   xml_declaration: bool = True, pretty_print: bool = True) -> None:
    """Write an XML tree to file with proper formatting.
    
    Args:
        tree: XML tree to write
        file_path: Output file path
        encoding: File encoding
        xml_declaration: Whether to include XML declaration
        pretty_print: Whether to format output with proper indentation
    """
    try:
        if LXML_AVAILABLE:
            tree.write(file_path, encoding=encoding, 
                      xml_declaration=xml_declaration,
                      pretty_print=pretty_print)
        else:
            tree.write(file_path, encoding=encoding,
                      xml_declaration=xml_declaration)
    except Exception as e:
        print(f"Error writing XML file {file_path}: {e}")
        raise