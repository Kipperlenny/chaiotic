"""Module for extracting structured content from documents."""

import os
import re
import zipfile
from lxml import etree
from typing import List, Dict, Any

from .document_reader import DOCX_AVAILABLE, ODF_AVAILABLE, read_document

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
    
    if file_ext == '.docx':
        return extract_structured_docx(file_path)
    elif file_ext == '.odt':
        return extract_structured_odt(file_path)
    elif file_ext == '.txt':
        return extract_structured_text(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_ext}")

def extract_structured_docx(file_path):
    """Extract structured content from a DOCX file."""
    try:
        if not DOCX_AVAILABLE:
            print("python-docx library not available. Using basic XML parsing for structured DOCX.")
            print("For better results, install: pip install python-docx")
            return extract_structured_docx_xml(file_path)
        
        from docx import Document
        doc = Document(file_path)
        structured_content = []
        
        # Process paragraphs
        for i, para in enumerate(doc.paragraphs):
            if not para.text.strip():
                continue  # Skip empty paragraphs
                
            # Determine paragraph type based on style
            para_type = "paragraph"
            if para.style.name.startswith('Heading'):
                para_type = f"heading-{para.style.name.split(' ')[-1]}"
            elif para.style.name == 'List Paragraph':
                para_type = "list-item"
                
            structured_content.append({
                'id': f'p{i+1}',
                'type': para_type,
                'content': para.text
            })
        
        # Process tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    # Combine all paragraphs in a cell
                    cell_text = '\n'.join(p.text for p in cell.paragraphs if p.text.strip())
                    if cell_text.strip():
                        structured_content.append({
                            'id': f't{t_idx+1}r{r_idx+1}c{c_idx+1}',
                            'type': 'table-cell',
                            'content': cell_text
                        })
        
        return structured_content
    except Exception as e:
        print(f"Error extracting structured content from DOCX: {e}")
        # Fall back to XML parsing
        try:
            return extract_structured_docx_xml(file_path)
        except Exception as e2:
            print(f"XML fallback failed: {e2}")
            # Last resort - use simple text
            content, _, _ = read_document(file_path)
            return extract_structured_text_from_content(content)

def extract_structured_docx_xml(file_path):
    """Extract structured content from a DOCX file using XML parsing."""
    # ...existing code...

def extract_structured_odt(file_path):
    """Extract structured content from an ODT file."""
    # ...existing code...

def extract_structured_odt_xml(file_path):
    """Extract structured content from an ODT file using XML parsing."""
    # ...existing code...

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