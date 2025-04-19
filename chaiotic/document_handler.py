"""Document handler module for reading and writing documents."""

import os
import re
import zipfile
import tempfile
import io
import difflib
from lxml import etree
from typing import Tuple, List, Dict, Any, Optional
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement, qn
from docx.enum.text import WD_COLOR_INDEX

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
    """Read a DOCX document by parsing its XML directly.
    
    This is used as a fallback when the python-docx library is not available.
    
    Args:
        file_path: Path to the DOCX file.
        
    Returns:
        Tuple of (content, None, True)
    """
    # Define XML namespaces used in DOCX
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    try:
        # DOCX files are ZIP archives
        with zipfile.ZipFile(file_path) as docx_zip:
            # The document content is in word/document.xml
            if 'word/document.xml' not in docx_zip.namelist():
                raise ValueError("Invalid DOCX format: word/document.xml not found")
                
            with docx_zip.open('word/document.xml') as f:
                xml_content = f.read()
                
            # Parse XML
            root = etree.fromstring(xml_content)
            
            # Extract paragraphs
            paragraphs = []
            for para in root.xpath('//w:p', namespaces=namespaces):
                text_elements = para.xpath('.//w:t', namespaces=namespaces)
                texts = [elem.text if elem.text is not None else '' for elem in text_elements]
                paragraph_text = ''.join(texts)
                paragraphs.append(paragraph_text)
            
            content = '\n'.join(paragraphs)
            return content, None, True
    except Exception as e:
        print(f"Error parsing DOCX XML: {e}")
        raise

def read_docx_as_zip(file_path):
    """Read a DOCX as a simple ZIP and extract text from document.xml.
    
    This is a last resort when other methods fail.
    
    Args:
        file_path: Path to the DOCX file.
        
    Returns:
        Tuple of (content, None, True)
    """
    try:
        text_content = []
        with zipfile.ZipFile(file_path) as z:
            if 'word/document.xml' in z.namelist():
                with z.open('word/document.xml') as f:
                    content = f.read().decode('utf-8')
                    # Simple regex to extract text between <w:t> tags
                    text_parts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', content, re.DOTALL)
                    text_content = ' '.join(text_parts)
                    # Split into paragraphs based on common paragraph breaks
                    paragraphs = re.split(r'</w:p>\s*<w:p', text_content)
                    text_content = '\n'.join(paragraph.strip() for paragraph in paragraphs)
        
        return text_content, None, True
    except Exception as e:
        print(f"Error reading DOCX as ZIP: {e}")
        raise

def read_odt(file_path):
    """Read an ODT document.
    
    Args:
        file_path: Path to the ODT file.
        
    Returns:
        Tuple of (content, document_object, False)
    """
    try:
        if not ODF_AVAILABLE:
            print("odfpy library not available. Using basic XML parsing for ODT.")
            print("For better results, install: pip install odfpy")
            return read_odt_xml(file_path)
        
        doc = odf_load(file_path)
        full_text = []
        
        # Extract paragraphs
        for element in doc.getElementsByType(odf_text.P):
            paragraph_text = element.plainText()
            full_text.append(paragraph_text)
        
        content = '\n'.join(full_text)
        return content, doc, False
    except Exception as e:
        print(f"Error reading ODT document: {e}")
        # Fallback to XML parsing
        try:
            return read_odt_xml(file_path)
        except Exception as e2:
            print(f"Error reading ODT as XML: {e2}")
            try:
                return read_odt_as_zip(file_path)
            except Exception as e3:
                print(f"All ODT reading methods failed. Final error: {e3}")
                raise ValueError(f"Could not read ODT file: {file_path}")

def read_odt_xml(file_path):
    """Read an ODT document by parsing its XML directly.
    
    This is used as a fallback when the odfpy library is not available.
    
    Args:
        file_path: Path to the ODT file.
        
    Returns:
        Tuple of (content, None, False)
    """
    # Define XML namespaces used in ODT
    namespaces = {
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
    }
    
    try:
        # ODT files are ZIP archives
        with zipfile.ZipFile(file_path) as odt_zip:
            # The document content is in content.xml
            if 'content.xml' not in odt_zip.namelist():
                raise ValueError("Invalid ODT format: content.xml not found")
                
            with odt_zip.open('content.xml') as f:
                xml_content = f.read()
                
            # Parse XML
            root = etree.fromstring(xml_content)
            
            # Extract paragraphs
            paragraphs = []
            for para in root.xpath('//text:p', namespaces=namespaces):
                # Get all text content, including from spans
                paragraph_text = ''.join(para.xpath('.//text()', namespaces=namespaces))
                paragraphs.append(paragraph_text)
            
            content = '\n'.join(paragraphs)
            return content, None, False
    except Exception as e:
        print(f"Error parsing ODT XML: {e}")
        raise

def read_odt_as_zip(file_path):
    """Read an ODT as a simple ZIP and extract text from content.xml.
    
    This is a last resort when other methods fail.
    
    Args:
        file_path: Path to the ODT file.
        
    Returns:
        Tuple of (content, None, False)
    """
    try:
        text_content = []
        with zipfile.ZipFile(file_path) as z:
            if 'content.xml' in z.namelist():
                with z.open('content.xml') as f:
                    content = f.read().decode('utf-8')
                    # Simple regex to extract text between paragraph tags
                    text_parts = re.findall(r'<text:p[^>]*>(.*?)</text:p>', content, re.DOTALL)
                    # Remove XML tags to get plain text
                    for part in text_parts:
                        clean_text = re.sub(r'<[^>]+>', '', part)
                        text_content.append(clean_text)
        
        return '\n'.join(text_content), None, False
    except Exception as e:
        print(f"Error reading ODT as ZIP: {e}")
        raise

def read_text_file(file_path):
    """Read a plain text file.
    
    Args:
        file_path: Path to the text file.
        
    Returns:
        Tuple of (content, None, None)
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return content, None, None
    except UnicodeDecodeError:
        # Try with a different encoding if UTF-8 fails
        try:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()
            return content, None, None
        except Exception as e:
            print(f"Error reading text file: {e}")
            raise

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
    """Extract structured content from a DOCX file.
    
    Args:
        file_path: Path to the DOCX file.
        
    Returns:
        List of dictionaries containing structured content
    """
    try:
        if not DOCX_AVAILABLE:
            print("python-docx library not available. Using basic XML parsing for structured DOCX.")
            print("For better results, install: pip install python-docx")
            return extract_structured_docx_xml(file_path)
        
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
                'id': f'p{i}',
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
                            'id': f't{t_idx}r{r_idx}c{c_idx}',
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
    """Extract structured content from a DOCX file using XML parsing.
    
    Args:
        file_path: Path to the DOCX file.
        
    Returns:
        List of dictionaries containing structured content
    """
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }
    
    try:
        structured_content = []
        
        with zipfile.ZipFile(file_path) as docx_zip:
            if 'word/document.xml' not in docx_zip.namelist():
                raise ValueError("Invalid DOCX format: word/document.xml not found")
                
            with docx_zip.open('word/document.xml') as f:
                xml_content = f.read()
                
            # Parse XML
            root = etree.fromstring(xml_content)
            
            # First check if we have styles available to determine headings
            styles_map = {}
            style_map_created = False
            
            if 'word/styles.xml' in docx_zip.namelist():
                try:
                    with docx_zip.open('word/styles.xml') as f:
                        styles_xml = f.read()
                    styles_root = etree.fromstring(styles_xml)
                    
                    # Extract style information
                    for style in styles_root.xpath('//w:style', namespaces=namespaces):
                        style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
                        style_name = None
                        
                        # Get style name
                        name_elem = style.find('.//w:name', namespaces)
                        if name_elem is not None:
                            style_name = name_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        
                        if style_id and style_name:
                            styles_map[style_id] = style_name
                            style_map_created = True
                except Exception as e:
                    print(f"Warning: Could not parse styles.xml: {e}")
            
            # Process paragraphs
            for i, para in enumerate(root.xpath('//w:p', namespaces=namespaces)):
                # Extract text
                text_elements = para.xpath('.//w:t', namespaces=namespaces)
                texts = [elem.text if elem.text is not None else '' for elem in text_elements]
                paragraph_text = ''.join(texts)
                
                if not paragraph_text.strip():
                    continue  # Skip empty paragraphs
                
                # Determine paragraph type based on style
                para_type = "paragraph"
                
                # Try to identify heading styles if we have the style map
                if style_map_created:
                    p_style = para.find('.//w:pStyle', namespaces=namespaces)
                    if p_style is not None:
                        style_id = p_style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        if style_id in styles_map and styles_map[style_id].startswith('Heading'):
                            heading_level = styles_map[style_id].split()[-1]
                            if heading_level.isdigit():
                                para_type = f"heading-{heading_level}"
                
                structured_content.append({
                    'id': f'p{i}',
                    'type': para_type,
                    'content': paragraph_text
                })
        
        return structured_content
    except Exception as e:
        print(f"Error extracting structured content from DOCX XML: {e}")
        # Fallback to simple text structure
        content, _, _ = read_document(file_path)
        return extract_structured_text_from_content(content)

def extract_structured_odt(file_path):
    """Extract structured content from an ODT file.
    
    Args:
        file_path: Path to the ODT file.
        
    Returns:
        List of dictionaries containing structured content
    """
    try:
        if not ODF_AVAILABLE:
            print("odfpy library not available. Using basic XML parsing for structured ODT.")
            print("For better results, install: pip install odfpy")
            return extract_structured_odt_xml(file_path)
        
        doc = odf_load(file_path)
        structured_content = []
        
        # Process all elements
        element_index = 0
        
        # Extract text elements
        for element in doc.getElementsByType(odf_text.P):
            # Skip empty paragraphs
            if not element.plainText().strip():
                continue
            
            # Try to determine element type
            style_name = element.getAttribute('stylename')
            para_type = "paragraph"
            
            # Check if it's a heading
            if style_name and style_name.startswith('Heading'):
                heading_parts = style_name.split()
                if len(heading_parts) > 1 and heading_parts[1].isdigit():
                    para_type = f"heading-{heading_parts[1]}"
            
            structured_content.append({
                'id': f'p{element_index}',
                'type': para_type,
                'content': element.plainText()
            })
            
            element_index += 1
        
        return structured_content
    except Exception as e:
        print(f"Error extracting structured content from ODT: {e}")
        # Fallback to XML parsing
        try:
            return extract_structured_odt_xml(file_path)
        except Exception as e2:
            print(f"XML fallback failed: {e2}")
            # Last resort - use simple text
            content, _, _ = read_document(file_path)
            return extract_structured_text_from_content(content)

def extract_structured_odt_xml(file_path):
    """Extract structured content from an ODT file using XML parsing.
    
    Args:
        file_path: Path to the ODT file.
        
    Returns:
        List of dictionaries containing structured content
    """
    namespaces = {
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'
    }
    
    try:
        structured_content = []
        
        with zipfile.ZipFile(file_path) as odt_zip:
            if 'content.xml' not in odt_zip.namelist():
                raise ValueError("Invalid ODT format: content.xml not found")
                
            with odt_zip.open('content.xml') as f:
                xml_content = f.read()
                
            # Parse XML
            root = etree.fromstring(xml_content)
            
            # Process paragraphs
            for i, para in enumerate(root.xpath('//text:p', namespaces=namespaces)):
                # Extract text (all text nodes within the paragraph)
                paragraph_text = ''.join(para.xpath('.//text()', namespaces=namespaces))
                
                if not paragraph_text.strip():
                    continue  # Skip empty paragraphs
                
                # Determine paragraph type
                para_type = "paragraph"
                style_name = para.get('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name')
                
                if style_name and style_name.startswith('Heading'):
                    heading_parts = style_name.split()
                    if len(heading_parts) > 1 and heading_parts[1].isdigit():
                        para_type = f"heading-{heading_parts[1]}"
                
                structured_content.append({
                    'id': f'p{i}',
                    'type': para_type,
                    'content': paragraph_text
                })
        
        return structured_content
    except Exception as e:
        print(f"Error extracting structured content from ODT XML: {e}")
        # Fallback to simple text structure
        content, _, _ = read_document(file_path)
        return extract_structured_text_from_content(content)

def extract_structured_text(file_path):
    """Extract structured content from a plain text file.
    
    Args:
        file_path: Path to the text file.
        
    Returns:
        List of dictionaries containing structured content
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # Try with a different encoding if UTF-8 fails
        with open(file_path, 'r', encoding='latin-1') as f:
            content = f.read()
    
    return extract_structured_text_from_content(content)

def extract_structured_text_from_content(content):
    """Extract structured content from plain text content.
    
    Args:
        content: The text content to structure.
        
    Returns:
        List of dictionaries containing structured content
    """
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
                'id': f'p{i}',
                'type': 'heading-1',
                'content': para
            })
        elif len(lines) == 1 and len(para) < 100 and para[0].isupper() and not para.endswith('.'):
            structured_content.append({
                'id': f'p{i}',
                'type': 'heading-2',
                'content': para
            })
        else:
            # Regular paragraph
            structured_content.append({
                'id': f'p{i}',
                'type': 'paragraph',
                'content': para
            })
    
    return structured_content

def create_sample_document(file_path, content=None):
    """Create a sample document for testing.
    
    Args:
        file_path: Path where to save the document.
        content: Optional content to include in the document.
        
    Returns:
        Path to the created document.
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if not content:
        content = (
            "Überschrift 1\n\n"
            "Dies ist ein Absatz mit einigen falsch geschriebenen Wörtern. "
            "Zum Beispeil hier und drot und villeicht auch anderswo.\n\n"
            "Überschrift 2\n\n"
            "Dieser Abschnit enthält auch Grammatikfehler, die der korrektur bedürfen. "
            "Die Katze hast auf dem Sofa gesessen, obwohl sie dass nicht tun sollte.\n\n"
            "Manchmal fehlen auch Kommas oder es gibt zu viele davon, oder die Satzstellung ist falsch verwendet worden."
        )
    
    if file_ext == '.docx':
        return create_sample_docx(file_path, content)
    elif file_ext == '.odt':
        return create_sample_odt(file_path, content)
    elif file_ext == '.txt':
        return create_sample_txt(file_path, content)
    else:
        raise ValueError(f"Unsupported file type: {file_ext}")

def create_sample_docx(file_path, content):
    """Create a sample DOCX document.
    
    Args:
        file_path: Path where to save the document.
        content: Content to include in the document.
        
    Returns:
        Path to the created document.
    """
    if not DOCX_AVAILABLE:
        print("python-docx library not available. Creating a plain text file instead.")
        return create_sample_txt(os.path.splitext(file_path)[0] + '.txt', content)
    
    try:
        doc = Document()
        
        # Split into paragraphs
        paragraphs = re.split(r'\n\s*\n', content)
        
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
            
            # Check if it looks like a heading (short, single line)
            lines = para.split('\n')
            if len(lines) == 1 and len(para) < 100 and not para.endswith('.'):
                # Probable heading - add as Heading 1
                doc.add_heading(para, level=1)
            else:
                # Regular paragraph
                doc.add_paragraph(para)
        
        doc.save(file_path)
        print(f"Sample DOCX created at {file_path}")
        return file_path
    except Exception as e:
        print(f"Error creating sample DOCX: {e}")
        # Fallback to text file
        return create_sample_txt(os.path.splitext(file_path)[0] + '.txt', content)

def create_sample_odt(file_path, content):
    """Create a sample ODT document.
    
    Args:
        file_path: Path where to save the document.
        content: Content to include in the document.
        
    Returns:
        Path to the created document.
    """
    if not ODF_AVAILABLE:
        print("odfpy library not available. Creating a plain text file instead.")
        return create_sample_txt(os.path.splitext(file_path)[0] + '.txt', content)
    
    try:
        from odf.opendocument import OpenDocumentText
        from odf.style import Style, TextProperties, ParagraphProperties
        from odf.text import H, P
        
        doc = OpenDocumentText()
        
        # Create heading style
        heading_style = Style(name="Heading", family="paragraph")
        heading_style.addElement(TextProperties(fontweight="bold", fontsize="16pt"))
        doc.styles.addElement(heading_style)
        
        # Split into paragraphs
        paragraphs = re.split(r'\n\s*\n', content)
        
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
            
            # Check if it looks like a heading (short, single line)
            lines = para.split('\n')
            if len(lines) == 1 and len(para) < 100 and not para.endswith('.'):
                # Probable heading - add as Heading
                heading = H(stylename=heading_style, outlinelevel=1)
                heading.addText(para)
                doc.text.addElement(heading)
            else:
                # Regular paragraph
                p = P()
                p.addText(para)
                doc.text.addElement(p)
        
        doc.save(file_path)
        print(f"Sample ODT created at {file_path}")
        return file_path
    except Exception as e:
        print(f"Error creating sample ODT: {e}")
        # Fallback to text file
        return create_sample_txt(os.path.splitext(file_path)[0] + '.txt', content)

def create_sample_txt(file_path, content):
    """Create a sample text document.
    
    Args:
        file_path: Path where to save the document.
        content: Content to include in the document.
        
    Returns:
        Path to the created document.
    """
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Sample text file created at {file_path}")
        return file_path
    except Exception as e:
        print(f"Error creating sample text file: {e}")
        raise

def save_document(file_path: str, corrections: List[Dict[Any, Any]], original_doc=None, is_docx: bool = True) -> str:
    """
    Save the document with corrections as Word comments instead of direct modifications.
    This function preserves the original document and adds review comments.
    
    Args:
        file_path: Path to the original document
        corrections: List of corrections to apply
        original_doc: Original document object if available
        is_docx: Whether the document is DOCX (otherwise ODT)
        
    Returns:
        Path to the saved document
    """
    # Generate output filename
    file_base, file_ext = os.path.splitext(file_path)
    output_path = f"{file_base}_reviewed{file_ext}"
    
    print(f"Saving document with review comments to {output_path}")
    print(f"Number of corrections to apply: {len(corrections)}")
    
    # Validate corrections to ensure they are all properly formatted dictionaries
    valid_corrections = []
    for correction in corrections:
        if not isinstance(correction, dict):
            print(f"Warning: Skipping non-dictionary correction: {correction}")
            continue
        if 'original' not in correction or 'corrected' not in correction:
            print(f"Warning: Skipping correction missing required fields: {correction}")
            continue
        valid_corrections.append(correction)
    
    print(f"After validation: {len(valid_corrections)} valid corrections")
    corrections = valid_corrections
    
    if is_docx:
        # If we don't have the original document, load it now
        if original_doc is None:
            original_doc = Document(file_path)
        
        # Create a copy of the original document (don't create new - preserve all formatting)
        doc = Document(file_path)
        
        # Create a mapping of corrections by paragraph content
        correction_by_paragraph = {}
        for correction in corrections:
            if 'id' in correction and correction['id'].startswith('p'):
                para_id = correction['id']
                if para_id not in correction_by_paragraph:
                    correction_by_paragraph[para_id] = []
                correction_by_paragraph[para_id].append(correction)
        
        # Also create a backup map using paragraph text
        correction_by_text = {}
        for correction in corrections:
            if 'original' in correction:
                original_text = correction['original']
                if original_text not in correction_by_text:
                    correction_by_text[original_text] = []
                correction_by_text[original_text].append(correction)
        
        # Process each paragraph from the original document
        for i, para in enumerate(doc.paragraphs):
            para_id = f"p{i+1}"  # Create a paragraph ID similar to what we use in structured content
            para_text = para.text
            
            # Skip empty paragraphs
            if not para_text.strip():
                continue
            
            # Check if we have corrections for this paragraph
            para_corrections = correction_by_paragraph.get(para_id, [])
            
            # If not, try to find by paragraph text
            if not para_corrections:
                # Look for any original text that might be in this paragraph
                for orig_text, corrs in correction_by_text.items():
                    if orig_text in para_text:
                        para_corrections.extend(corrs)
            
            # If no corrections for this paragraph, skip it
            if not para_corrections:
                continue
            
            # Debug this paragraph
            print(f"Processing paragraph {para_id}: '{para_text}'")
            print(f"Found {len(para_corrections)} corrections for this paragraph")
            
            # This paragraph has corrections - need to add comments
            # First, identify character positions of corrections
            sorted_corrections = []
            for correction in para_corrections:
                original = correction.get('original', '')
                corrected = correction.get('corrected', '')
                explanation = correction.get('explanation', '')
                
                # Find all occurrences of the original text in the paragraph
                start_pos = 0
                while True:
                    pos = para_text.find(original, start_pos)
                    if pos == -1:
                        # If exact match not found, try fuzzy matching
                        if len(original) > 5:  # Only for longer text to avoid false positives
                            # Try to find the closest match
                            best_ratio = 0
                            best_pos = -1
                            for j in range(len(para_text) - len(original) + 1):
                                substring = para_text[j:j+len(original)]
                                ratio = difflib.SequenceMatcher(None, original, substring).ratio()
                                if ratio > 0.8 and ratio > best_ratio:  # 80% similarity threshold
                                    best_ratio = ratio
                                    best_pos = j
                            
                            if best_pos != -1:
                                print(f"Found approximate match for '{original}' at position {best_pos} (similarity: {best_ratio:.2f})")
                                pos = best_pos
                            else:
                                break
                        else:
                            break
                            
                    sorted_corrections.append({
                        'start': pos,
                        'end': pos + len(original),
                        'original': original,
                        'corrected': corrected,
                        'explanation': explanation
                    })
                    start_pos = pos + len(original)
            
            # Sort corrections by start position (ascending)
            sorted_corrections.sort(key=lambda x: x['start'])
            
            # Check if we found positions for all corrections
            if len(sorted_corrections) != len(para_corrections):
                print(f"Warning: Could not locate all corrections in paragraph {para_id}")
                for corr in para_corrections:
                    if not any(sc['original'] == corr.get('original', '') for sc in sorted_corrections):
                        print(f"  Could not locate: '{corr.get('original', '')}' -> '{corr.get('corrected', '')}'")
            
            # If we didn't find any correction positions, skip this paragraph
            if not sorted_corrections:
                continue
            
            # Now add Word comments for each correction
            for corr in sorted_corrections:
                start = corr['start']
                end = corr['end']
                original = corr['original']
                corrected = corr['corrected']
                explanation = corr['explanation']
                
                # Create comment content
                comment_text = f"Änderungsvorschlag: '{original}' → '{corrected}'"
                if explanation:
                    comment_text += f"\n\nBegründung: {explanation}"
                
                # Add comment to the paragraph at the specified position
                try:
                    # We need to split the paragraph if it doesn't have appropriate runs
                    if len(para.runs) == 0:
                        # Create a run with the entire paragraph text
                        para.text = ""
                        para.add_run(para_text)
                    
                    # Find or create appropriate runs for adding the comment
                    target_run = None
                    current_pos = 0
                    
                    # Find which run contains the start position
                    for run in para.runs:
                        if current_pos + len(run.text) > start:
                            # This run contains our starting position
                            # We may need to split this run
                            if current_pos < start:
                                # Split: text before our target, target text, text after
                                rel_start = start - current_pos
                                rel_end = min(end - current_pos, len(run.text))
                                
                                # Store original text and formatting
                                original_text = run.text
                                
                                # Split the run
                                run.text = original_text[:rel_start]
                                target_run = para.add_run(original_text[rel_start:rel_end])
                                copy_run_formatting(run, target_run)
                                
                                # If there's text after the end, add another run
                                if rel_end < len(original_text):
                                    after_run = para.add_run(original_text[rel_end:])
                                    copy_run_formatting(run, after_run)
                            else:
                                # The start position is at the beginning of this run
                                target_run = run
                            
                            break
                        current_pos += len(run.text)
                    
                    # If we couldn't find an appropriate run, skip this correction
                    if target_run is None:
                        print(f"Warning: Could not identify run for comment at position {start}-{end}")
                        continue
                    
                    # Add the comment to the target run
                    add_comment(para, target_run, comment_text, "Grammatik-Prüfung")
                    
                except Exception as e:
                    print(f"Error adding comment: {e}")
                    continue
        
        # Save the document
        try:
            doc.save(output_path)
            print(f"Successfully saved document with review comments to {output_path}")
            return output_path
        except Exception as e:
            print(f"Error saving document: {str(e)}")
            # Try saving to a different location
            fallback_path = os.path.join(os.path.dirname(file_path), "reviewed_output.docx")
            try:
                doc.save(fallback_path)
                print(f"Saved to alternate location: {fallback_path}")
                return fallback_path
            except Exception as e2:
                print(f"Failed to save document: {str(e2)}")
                raise
    else:
        # Handle ODT files - fallback to text highlighting since ODT comments are not supported
        from chaiotic.utils import save_document as utils_save_document
        print("Note: Word comments are not supported for ODT files. Using text highlighting instead.")
        _, _, output_path = utils_save_document(file_path, corrections, original_doc, is_docx=False)
        return output_path

def add_comment(paragraph, run, comment_text, author="Grammatik-Prüfung"):
    """Add a Word comment to a specific run in a paragraph.
    
    Args:
        paragraph: The paragraph containing the run
        run: The run to comment on
        comment_text: The text of the comment
        author: The author name for the comment
    """
    # Get the document from the paragraph
    doc = paragraph.part.document
    
    # Create a new comment
    comments_part = doc.part.package.get_part('word/comments.xml')
    if comments_part is None:
        # If comments part doesn't exist, we need to create it
        create_comments_part(doc)
        comments_part = doc.part.package.get_part('word/comments.xml')
    
    # Get next comment ID
    comment_id = get_next_comment_id(doc)
    
    # Add the comment reference to the run
    comment_reference = create_comment_reference(comment_id)
    run._r.append(comment_reference)
    
    # Add the actual comment
    comment = create_comment(comment_id, comment_text, author)
    comments_root = comments_part.element.getroot()
    comments_root.append(comment)

def create_comments_part(document):
    """Create the comments part if it doesn't exist."""
    from docx.opc.constants import CONTENT_TYPE as CT
    
    # Create the comments part
    document.part.package.add_part('/word/comments.xml', CT.COMMENTS)
    
    # Create minimal comments XML structure
    comments_xml = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    </w:comments>
    """
    comments_part = document.part.package.get_part('/word/comments.xml')
    comments_part.blob = comments_xml.encode('utf-8')
    
    # Update content types
    content_types = document.part.package.content_types
    if not content_types.has_override('/word/comments.xml'):
        content_types.add_override('/word/comments.xml', CT.COMMENTS)
    
    # Update document rels
    document.part.rels.add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '/word/comments.xml')

def get_next_comment_id(document):
    """Get the next available comment ID."""
    try:
        comments_part = document.part.package.get_part('/word/comments.xml')
        if comments_part is None:
            return 0
            
        comments = parse_xml(comments_part.blob)
        comment_ids = [int(comment.get(qn('w:id'))) for comment in comments.xpath('.//w:comment', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})]
        return max(comment_ids) + 1 if comment_ids else 0
    except:
        return 0

def create_comment_reference(comment_id):
    """Create XML for a comment reference."""
    comment_reference = OxmlElement('w:commentReference')
    comment_reference.set(qn('w:id'), str(comment_id))
    return comment_reference

def create_comment(comment_id, text, author, initials=None):
    """Create an XML comment element."""
    from datetime import datetime
    
    comment = OxmlElement('w:comment')
    comment.set(qn('w:id'), str(comment_id))
    comment.set(qn('w:author'), author)
    if initials:
        comment.set(qn('w:initials'), initials)
    comment.set(qn('w:date'), datetime.now().isoformat())
    
    # Add the comment text
    paragraph = OxmlElement('w:p')
    run = OxmlElement('w:r')
    text_element = OxmlElement('w:t')
    text_element.text = text
    run.append(text_element)
    paragraph.append(run)
    comment.append(paragraph)
    
    return comment

def copy_run_formatting(source_run, target_run):
    """Copy formatting from source run to target run."""
    # Copy basic formatting
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    target_run.font.strike = source_run.font.strike
    target_run.font.subscript = source_run.font.subscript
    target_run.font.superscript = source_run.font.superscript
    
    # Copy font properties
    if source_run.font.name:
        target_run.font.name = source_run.font.name
    if source_run.font.size:
        target_run.font.size = source_run.font.size
    if source_run.font.color.rgb:
        target_run.font.color.rgb = source_run.font.color.rgb
    
    # Copy style if present
    if hasattr(source_run, 'style') and source_run.style:
        target_run.style = source_run.style