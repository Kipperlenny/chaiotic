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
    """Read document content from supported file types (DOCX, ODT)."""
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.docx':
        return read_docx(file_path)
    elif file_ext == '.odt':
        return read_odt(file_path)
    else:
        print(f"Unsupported file format: {file_ext}")
        return None, None, None

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
    """Read an ODT document."""
    # ...existing code...

def read_odt_xml(file_path):
    """Read an ODT document by parsing its XML directly."""
    # ...existing code...

def read_odt_as_zip(file_path):
    """Read an ODT as a simple ZIP and extract text from content.xml."""
    # ...existing code...

def read_text_file(file_path):
    """Read a plain text file."""
    # ...existing code...