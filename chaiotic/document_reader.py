"""Module for reading document files of various formats."""

import os
import re
import zipfile
from lxml import etree

# Check for odf library availability and get element classes
ODF_AVAILABLE = False
try:
    from odf.opendocument import OpenDocument
    from odf.text import P, H, Span
    from odf.element import Element as OdfElement
    from odf import teletype
    ODF_AVAILABLE = True
except ImportError:
    ODF_AVAILABLE = False

def read_document(file_path):
    """Read document content and return structured format."""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.odt':
        content, structured_content = read_odt(file_path)
        # No need for conversion since read_odt now returns proper format
        return content, structured_content
    else:
        content = read_text_file(file_path)
        return content, None

def read_odt(file_path):
    """Read an ODT file and return its content as text."""
    try:
        from odf.opendocument import load
        
        # Load the ODT file
        print(f"Reading ODT file: {file_path}")
        doc = load(file_path)
        
        # Extract all paragraphs and headings
        paragraphs = []
        structured_content = []
        
        # Extract headings and paragraphs
        for i, element in enumerate(doc.text.childNodes):
            # Check if element is an ODT text element
            if isinstance(element, OdfElement) and element.qname[1] in ('p', 'h'):
                content = teletype.extractText(element)
                if content.strip():
                    paragraphs.append(content)
                    structured_content.append({
                        'id': f'p{i+1}',
                        'type': 'heading' if element.qname[1] == 'h' else 'paragraph',
                        'content': content
                    })
        
        # Join paragraphs with newlines
        content = '\n\n'.join(paragraphs)
        
        print(f"Successfully read ODT content ({len(content)} characters)")
        return content, structured_content
        
    except ImportError:
        print("odfpy library not found, using alternative extraction method")
        return read_odt_xml(file_path)

def read_odt_xml(file_path):
    """Read an ODT document by parsing its XML directly."""
    try:
        content = []
        structured_content = []
        
        with zipfile.ZipFile(file_path) as odt:
            with odt.open('content.xml') as f:
                tree = etree.parse(f)
                root = tree.getroot()
                
                # Define namespaces
                ns = {
                    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
                }
                
                # Get all paragraphs and headings
                elements = root.xpath('//text:p | //text:h', namespaces=ns)
                
                for i, elem in enumerate(elements):
                    text = elem.text or ''
                    # Add all text from child elements
                    for child in elem.xpath('.//text()', namespaces=ns):
                        text += child
                    
                    if text.strip():
                        content.append(text)
                        elem_type = 'heading' if elem.tag.endswith('h') else 'paragraph'
                        structured_content.append({
                            'id': f'p{i+1}',
                            'type': elem_type,
                            'content': text
                        })
        
        return '\n\n'.join(content), structured_content
        
    except Exception as e:
        print(f"Error reading ODT XML: {e}")
        return read_odt_as_zip(file_path)

def read_odt_as_zip(file_path):
    """Read an ODT as a simple ZIP and extract text from content.xml."""
    try:
        content = []
        structured_content = []
        
        with zipfile.ZipFile(file_path) as odt:
            # Check if content.xml exists
            if 'content.xml' not in odt.namelist():
                raise ValueError("Invalid ODT file: content.xml not found")
                
            # Extract content.xml as text
            content_xml = odt.read('content.xml').decode('utf-8')
            
            # Simple regex-based extraction for paragraphs
            # This is a last resort method when proper XML parsing fails
            paragraphs = re.findall(r'<text:p[^>]*>(.*?)</text:p>', content_xml, re.DOTALL)
            headings = re.findall(r'<text:h[^>]*>(.*?)</text:h>', content_xml, re.DOTALL)
            
            # Clean up XML tags from content
            for i, p in enumerate(paragraphs):
                # Remove XML tags
                clean_text = re.sub(r'<[^>]+>', ' ', p).strip()
                # Normalize whitespace
                clean_text = re.sub(r'\s+', ' ', clean_text).strip()
                
                if clean_text:
                    content.append(clean_text)
                    structured_content.append({
                        'id': f'p{i+1}',
                        'type': 'paragraph',
                        'content': clean_text
                    })
            
            # Process headings similarly
            for i, h in enumerate(headings):
                clean_text = re.sub(r'<[^>]+>', ' ', h).strip()
                clean_text = re.sub(r'\s+', ' ', clean_text).strip()
                
                if clean_text:
                    content.append(clean_text)
                    structured_content.append({
                        'id': f'h{i+1}',
                        'type': 'heading',
                        'content': clean_text
                    })
        
        if not content:
            return "Failed to extract text from ODT file.", []
            
        return '\n\n'.join(content), structured_content
        
    except Exception as e:
        print(f"Error extracting content from ODT as ZIP: {e}")
        return f"Error reading document: {str(e)}", []

def read_text_file(file_path):
    """Read a plain text file."""
    try:
        encodings = ['utf-8', 'latin-1', 'cp1252', 'ascii']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                print(f"Successfully read text file with {encoding} encoding")
                return content
            except UnicodeDecodeError:
                continue
        
        # If all encodings fail, try binary and attempt to detect
        with open(file_path, 'rb') as f:
            binary_content = f.read()
            
        # Try to detect encoding from BOM
        if binary_content.startswith(b'\xef\xbb\xbf'):
            # UTF-8 with BOM
            content = binary_content[3:].decode('utf-8')
        elif binary_content.startswith(b'\xff\xfe'):
            # UTF-16 (LE)
            content = binary_content[2:].decode('utf-16-le')
        elif binary_content.startswith(b'\xfe\xff'):
            # UTF-16 (BE)
            content = binary_content[2:].decode('utf-16-be')
        else:
            # Last resort: try UTF-8 ignoring errors
            content = binary_content.decode('utf-8', errors='replace')
        
        return content
        
    except Exception as e:
        print(f"Error reading text file: {e}")
        return f"Error reading document: {str(e)}"