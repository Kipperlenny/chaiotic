"""Module for extracting structured content from documents."""

import os
import re

from .document_reader import ODF_AVAILABLE

def extract_structured_content(file_path):
    """Extract structured content from a document, preserving paragraph metadata.
    
    Args:
        file_path: Path to the document file.
        
    Returns:
        List of dictionaries containing structured content, where each dictionary has:
        - 'id': A unique identifier for the element
        - 'type': The element type (paragraph, heading, list, table)
        - 'content': The text content of the element
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.odt':
        content = extract_structured_odt(file_path)
    elif file_ext == '.txt':
        content = extract_structured_text(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_ext}")
    
    # Debug output
    print(f"\n===== EXTRACTED DOCUMENT STRUCTURE ({len(content)} elements) =====")
    for i, item in enumerate(content[:3], 1):  # Print first 3 items for preview
        print(f"\nItem {i}/{len(content)}: ID={item.get('id')}, Type={item.get('type')}")
        print(f"Content: {item.get('content')[:100]}{'...' if len(item.get('content', '')) > 100 else ''}")
        
        # Check if we have XML content
        if 'xml_content' in item:
            xml_preview = item['xml_content'][:150]
            print(f"XML: {xml_preview}{'...' if len(item['xml_content']) > 150 else ''}")
        
        # Check for nested elements
        if 'metadata' in item and 'nested_elements' in item['metadata']:
            nested = item['metadata']['nested_elements']
            print(f"Nested elements: {len(nested)}")
            if nested:
                print(f"  First nested: {nested[0].get('type')}, Text: {nested[0].get('text')[:50]}...")
    
    if len(content) > 3:
        print(f"\n... and {len(content) - 3} more items\n")
    
    print("=" * 60)
    
    return content

def extract_structured_odt(file_path):
    """Extract structured content from an ODT file."""
    try:
        if not ODF_AVAILABLE:
            print("odfpy library not available. Using basic XML parsing for structured ODT.")
            return extract_structured_odt_xml(file_path)
        
        from odf.opendocument import load
        from odf.text import P, H, Span
        from odf.table import Table, TableRow, TableCell
        import odf.teletype as teletype
        
        doc = load(file_path)
        structured_content = []
        
        def get_text_content(element):
            """Extract text from an ODT element with better handling of nested structures."""
            parts = []
            # Add debugging to track complex nesting
            element_type = element.__class__.__name__
            
            try:
                # For Span elements and similar nested structures
                if hasattr(element, 'childNodes'):
                    for child in element.childNodes:
                        if hasattr(child, 'data'):
                            parts.append(child.data)
                        elif hasattr(child, 'childNodes'):
                            # Handle nested elements
                            parts.append(get_text_content(child))
                    
                # Some elements might just have a text attribute
                elif hasattr(element, 'text'):
                    if element.text:
                        parts.append(element.text)
                    
                    # Handle children using ElementTree-style API
                    for child in element:
                        parts.append(get_text_content(child))
                        if child.tail:
                            parts.append(child.tail)
                
                # Handle direct text content
                elif hasattr(element, 'data'):
                    parts.append(element.data)
            except Exception as e:
                print(f"Warning: Error extracting text from {element_type}: {e}")
                
            return ''.join(parts)
        
        def get_element_xml(element):
            """Get the XML representation of the element."""
            try:
                # Try to directly get the XML
                return element.toXml()
            except Exception as e:
                print(f"Warning: Could not get XML directly: {e}")
                # Fallback approach if direct conversion fails
                try:
                    import lxml.etree as LET
                    from io import StringIO
                    
                    # Create a simplified XML representation
                    output = StringIO()
                    output.write(f"<{element.tagName}")
                    
                    # Add attributes if any
                    if hasattr(element, 'attributes') and element.attributes:
                        for attr_name, attr_value in element.attributes.items():
                            output.write(f' {attr_name}="{attr_value}"')
                    
                    output.write(">")
                    
                    # Add content if any
                    if hasattr(element, 'childNodes') and element.childNodes:
                        for child in element.childNodes:
                            if hasattr(child, 'toXml'):
                                output.write(child.toXml())
                            elif hasattr(child, 'data'):
                                output.write(child.data)
                    
                    output.write(f"</{element.tagName}>")
                    return output.getvalue()
                except Exception as e2:
                    print(f"Warning: Could not create XML representation: {e2}")
                    return f"<{element.tagName}>{get_text_content(element)}</{element.tagName}>"
        
        # Process paragraphs and headings with better tracking of nested elements
        idx = 0
        for element in doc.text.childNodes:
            try:
                if isinstance(element, (P, H)):
                    text = get_text_content(element)
                    xml_content = get_element_xml(element)
                    
                    # Get element style if available
                    style_name = None
                    if hasattr(element, 'attributes') and element.attributes:
                        style_name = element.attributes.get('text:style-name')
                    
                    element_info = {
                        'type': 'heading' if isinstance(element, H) else 'paragraph',
                        'text': text,
                        'xml': xml_content,
                        'style': style_name,
                        'nested': []
                    }
                    
                    # Check for nested spans and other elements
                    if hasattr(element, 'childNodes'):
                        for child in element.childNodes:
                            if isinstance(child, Span):
                                span_text = get_text_content(child)
                                span_xml = get_element_xml(child)
                                if span_text.strip():
                                    element_info['nested'].append({
                                        'type': 'span',
                                        'text': span_text,
                                        'xml': span_xml
                                    })
                    
                    if text.strip():
                        idx += 1
                        structured_content.append({
                            'id': f'p{idx}',
                            'type': element_info['type'],
                            'content': text,
                            'xml_content': xml_content,
                            'metadata': {
                                'nested_elements': element_info['nested'],
                                'style_name': style_name
                            }
                        })
                elif isinstance(element, Table):
                    for r_idx, row in enumerate(element.getElementsByType(TableRow)):
                        for c_idx, cell in enumerate(row.getElementsByType(TableCell)):
                            text = get_text_content(cell)
                            xml_content = get_element_xml(cell)
                            if text.strip():
                                idx += 1
                                structured_content.append({
                                    'id': f't{len(structured_content)+1}r{r_idx+1}c{c_idx+1}',
                                    'type': 'table-cell',
                                    'content': text,
                                    'xml_content': xml_content
                                })
            except Exception as e:
                print(f"Warning: Error processing element {type(element)}: {e}")
                continue
        
        return structured_content
        
    except Exception as e:
        print(f"Error extracting ODT content: {e}")
        return extract_structured_odt_xml(file_path)

def extract_structured_odt_xml(file_path):
    """Extract structured content from an ODT file using XML parsing."""
    import zipfile
    import xml.etree.ElementTree as ET
    
    try:
        # Try to use lxml for better XML handling
        import lxml.etree as LET
        parser = LET
    except ImportError:
        parser = ET
    
    structured_content = []
    
    try:
        with zipfile.ZipFile(file_path) as odt:
            # Read content.xml
            with odt.open('content.xml') as content:
                tree = parser.parse(content)
                root = tree.getroot()
                
                # Define namespaces
                ns = {
                    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
                    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
                }
                
                # Process paragraphs and headings
                idx = 0
                for element in root.findall('.//text:p', ns) + root.findall('.//text:h', ns):
                    text = ''.join(element.itertext()).strip()
                    if text:
                        idx += 1
                        structured_content.append({
                            'id': f'p{idx}',
                            'type': 'heading' if element.tag.endswith('}h') else 'paragraph',
                            'content': text
                        })
                
                # Process tables
                for t_idx, table in enumerate(root.findall('.//table:table', ns)):
                    for r_idx, row in enumerate(table.findall('.//table:table-row', ns)):
                        for c_idx, cell in enumerate(row.findall('.//table:table-cell', ns)):
                            text = ''.join(cell.itertext()).strip()
                            if text:
                                structured_content.append({
                                    'id': f't{t_idx+1}r{r_idx+1}c{c_idx+1}',
                                    'type': 'table-cell',
                                    'content': text
                                })
        
        return structured_content
        
    except Exception as e:
        print(f"Error extracting ODT XML content: {e}")
        return []

def extract_structured_text(file_path):
    """Extract structured content from a plain text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # Try with a different encoding if UTF-8 fails
        with open(file_path, 'r', encoding='latin-1') as f:
            content = f.read()
    
    return extract_structured_text_from_content(content)

def extract_structured_text_from_content(content):
    """Extract structured content from plain text content."""
    structured_content = []
    
    # Split into paragraphs
    paragraphs = re.split(r'\n\s*\n', content)
    
    for i, para in enumerate(paragraphs):
        para = para.strip()
        if not para:
            continue
        
        # Try to identify headings by looking for short single lines with capital letters
        lines = para.split('\n')
        if len(lines) == 1 and len(para) < 100 and para.isupper():
            structured_content.append({
                'id': f'p{i+1}',
                'type': 'heading-1',
                'content': para
            })
        elif len(lines) == 1 and len(para) < 100 and para[0].isupper() and not para.endswith('.'):
            structured_content.append({
                'id': f'p{i+1}',
                'type': 'heading-2',
                'content': para
            })
        else:
            # Regular paragraph
            structured_content.append({
                'id': f'p{i+1}',
                'type': 'paragraph',
                'content': para
            })
    
    return structured_content