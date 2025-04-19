"""Module for reading document files of various formats."""

import os
import re
import zipfile
import io
from lxml import etree
from typing import Tuple, List, Dict, Any, Optional

# Check for docx library availability
DOCX_AVAILABLE = False
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Check for odf library availability
ODF_AVAILABLE = False
try:
    from odf.opendocument import load as odf_load
    from odf import text as odf_text
    ODF_AVAILABLE = True
except ImportError:
    ODF_AVAILABLE = False

def read_document(file_path):
    """Read a document file and return its content as text.
    
    Args:
        file_path: Path to the document file
        
    Returns:
        Tuple of (content, document_object, is_docx)
    """
    # Check file extension
    file_ext = os.path.splitext(file_path)[1].lower()
    
    # Handle different file types
    if file_ext == '.docx':
        return read_docx(file_path)
    elif file_ext == '.odt':
        # Make sure to return a tuple with all the expected values
        content, doc_obj = read_odt(file_path)
        return content, doc_obj, False
    elif file_ext in ['.txt', '.md', '.rtf']:
        content = read_text_file(file_path)
        return content, None, False
    else:
        print(f"Unsupported file type: {file_ext}")
        return None, None, False

def read_docx(file_path):
    """Read content from a DOCX file and return structured representation."""
    try:
        # Add print for debugging
        print(f"Reading DOCX file: {file_path}")
        
        doc = Document(file_path)
        full_text = []
        structured_content = []
        element_id = 0
        
        # Extract each paragraph
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text:
                full_text.append(text)
                element_id += 1
                
                # Add to structured content with metadata
                structured_content.append({
                    'id': f'p{element_id}',
                    'type': 'paragraph',
                    'content': text,
                    'style': paragraph.style.name
                })
                
        # Also extract tables if present
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text = paragraph.text.strip()
                        if text:
                            full_text.append(text)
                            element_id += 1
                            
                            # Add to structured content with metadata
                            structured_content.append({
                                'id': f'p{element_id}',
                                'type': 'table_cell',
                                'content': text,
                                'style': paragraph.style.name
                            })
        
        full_content = '\n\n'.join(full_text)
        
        # Add debug prints to check what we're returning
        print(f"Extracted {len(structured_content)} structured elements from document")
        structured_found = len(structured_content) > 0
        print(f"Structured content found: {structured_found}")
        
        return full_content, doc, structured_content
        
    except Exception as e:
        print(f"Error reading DOCX file: {e}")
        return None, None, None

def read_docx_xml(file_path):
    """Read a DOCX document by parsing its XML directly."""
    # ...existing code...

def read_docx_as_zip(file_path):
    """Read a DOCX as a simple ZIP and extract text from document.xml."""
    # ...existing code...

def read_odt(file_path):
    """Read an ODT file and return its content as text.
    
    Args:
        file_path: Path to the ODT file
        
    Returns:
        Tuple of (content, odt_document_object)
    """
    # Try using odfpy if available
    try:
        from odf import text, teletype
        from odf.opendocument import load
        
        # Load the ODT file
        print(f"Reading ODT file: {file_path}")
        doc = load(file_path)
        
        # Extract all paragraphs and headings
        paragraphs = []
        
        # Extract headings (they are under text:h)
        headings = doc.getElementsByType(text.H)
        for heading in headings:
            paragraphs.append(teletype.extractText(heading))
        
        # Extract paragraphs (they are under text:p)
        paras = doc.getElementsByType(text.P)
        for para in paras:
            paragraphs.append(teletype.extractText(para))
        
        # Join paragraphs with newlines
        content = '\n\n'.join(paragraphs)
        
        print(f"Successfully read ODT content ({len(content)} characters)")
        return content, doc
        
    except ImportError:
        print("odfpy library not found, using alternative extraction method")
        
        # Alternative method: Extract content.xml from ODT (it's a ZIP file) and parse
        try:
            import zipfile
            import xml.etree.ElementTree as ET
            
            # Register namespaces for ElementTree
            namespaces = {
                'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
            }
            
            # Extract text content from content.xml in the ODT file
            with zipfile.ZipFile(file_path, 'r') as odt_file:
                content_xml = odt_file.read('content.xml')
                
                # Parse XML
                root = ET.fromstring(content_xml)
                
                # Extract paragraphs and headings
                paragraphs = []
                
                # Find all paragraph and heading elements
                text_elements = root.findall('.//text:p', namespaces) + root.findall('.//text:h', namespaces)
                
                # Extract text from each element
                for element in text_elements:
                    # Get all text content
                    text = ''.join(element.itertext())
                    if text.strip():  # Only add non-empty paragraphs
                        paragraphs.append(text)
                
                # Join paragraphs with newlines
                content = '\n\n'.join(paragraphs)
                
                print(f"Successfully extracted content from ODT ({len(content)} characters)")
                return content, odt_file
                
        except Exception as e:
            print(f"Error reading ODT file: {e}")
            import traceback
            traceback.print_exc()
            return "Error reading ODT file.", None

def read_odt_xml(file_path):
    """Read an ODT document by parsing its XML directly."""
    # ...existing code...

def read_odt_as_zip(file_path):
    """Read an ODT as a simple ZIP and extract text from content.xml."""
    # ...existing code...

def read_text_file(file_path):
    """Read a plain text file."""
    # ...existing code...