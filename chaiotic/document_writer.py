"""Document writer module for saving documents with tracked changes."""

import os
import json
import tempfile
import datetime
import shutil
import re
from pathlib import Path

# Try to import document handling libraries, with fallbacks
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    
try:
    from odf import text, teletype
    from odf.opendocument import load
    ODT_AVAILABLE = True
except ImportError:
    ODT_AVAILABLE = False

def save_correction_outputs(file_path, corrections, original_doc, is_docx=True):
    """Save correction outputs in multiple formats.
    
    Args:
        file_path: Path to the original file
        corrections: Dictionary containing correction data
        original_doc: Original document object
        is_docx: Whether this is a DOCX file
        
    Returns:
        Tuple of (json_file_path, text_file_path, doc_file_path)
    """
    try:
        # Create base output path
        base_path = os.path.splitext(file_path)[0]
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save corrections as JSON
        json_file_path = f"{base_path}_corrections_{timestamp}.json"
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(corrections, json_file, ensure_ascii=False, indent=2)
        
        # Save corrected text as plain text
        if "corrected_text" in corrections:
            text_file_path = f"{base_path}_corrected_{timestamp}.txt"
            with open(text_file_path, 'w', encoding='utf-8') as text_file:
                text_file.write(corrections["corrected_text"])
        else:
            text_file_path = None
        
        # Save as document with tracked changes
        if original_doc:
            if is_docx and DOCX_AVAILABLE:
                doc_file_path = save_docx_with_changes(file_path, corrections, original_doc, timestamp)
            elif not is_docx and ODT_AVAILABLE:
                doc_file_path = save_odt_with_changes(file_path, corrections, original_doc, timestamp)
            else:
                doc_file_path = None
        else:
            # Try to create document from scratch if no original_doc is available
            if "changes" in corrections:
                if is_docx and DOCX_AVAILABLE:
                    doc_file_path = create_docx_with_changes(file_path, corrections, timestamp)
                elif not is_docx and ODT_AVAILABLE:
                    doc_file_path = create_odt_with_changes(file_path, corrections, timestamp)
                else:
                    doc_file_path = None
            else:
                doc_file_path = None
        
        return json_file_path, text_file_path, doc_file_path
    
    except Exception as e:
        print(f"Error saving correction outputs: {e}")
        return None, None, None

def save_docx_with_changes(file_path, corrections, original_doc, timestamp):
    """Save a DOCX document with tracked changes applied.
    
    Args:
        file_path: Path to the original file
        corrections: Dictionary containing correction data
        original_doc: Original document object
        timestamp: Timestamp string for the output file
        
    Returns:
        Path to the saved document
    """
    try:
        # Create output path
        base_path = os.path.splitext(file_path)[0]
        output_path = f"{base_path}_corrected_{timestamp}.docx"
        
        # If we don't have a Document object, try to load it
        if original_doc is None and DOCX_AVAILABLE:
            try:
                original_doc = Document(file_path)
            except Exception as e:
                print(f"Could not load original document: {e}")
                return None
        
        # If we still don't have a Document object, we can't proceed
        if original_doc is None:
            return None
        
        # Get the changes from the corrections
        changes = corrections.get("changes", [])
        if not changes:
            print("No changes to apply")
            return None
        
        # Create a copy of the original document
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_file:
            temp_path = temp_file.name
        
        original_doc.save(temp_path)
        
        # Load the copy
        modified_doc = Document(temp_path)
        
        # Get all paragraphs
        paragraphs = list(modified_doc.paragraphs)
        
        # Apply changes with tracked revisions
        for change in changes:
            original = change.get("original", "")
            corrected = change.get("corrected", "")
            
            if original == corrected:
                continue
                
            # Find the paragraph containing this text
            for para in paragraphs:
                if original in para.text:
                    # Simple text replacement - more sophisticated revision tracking would need Word API
                    para.text = para.text.replace(original, f"[{original} → {corrected}]")
                    break
        
        # Save the modified document
        modified_doc.save(output_path)
        
        # Clean up
        try:
            os.unlink(temp_path)
        except:
            pass
        
        return output_path
    
    except Exception as e:
        print(f"Error saving DOCX with changes: {e}")
        return None

def save_odt_with_changes(file_path, corrections, original_doc, timestamp):
    """Save an ODT document with tracked changes applied.
    
    Args:
        file_path: Path to the original file
        corrections: Dictionary containing correction data
        original_doc: Original document object
        timestamp: Timestamp string for the output file
        
    Returns:
        Path to the saved document
    """
    try:
        # Create output path
        base_path = os.path.splitext(file_path)[0]
        output_path = f"{base_path}_corrected_{timestamp}.odt"
        
        # If we don't have a Document object, try to load it
        if original_doc is None and ODT_AVAILABLE:
            try:
                original_doc = load(file_path)
            except Exception as e:
                print(f"Could not load original document: {e}")
                return None
        
        # If we still don't have a Document object, we can't proceed
        if original_doc is None:
            return None
        
        # Get the changes from the corrections
        changes = corrections.get("changes", [])
        if not changes:
            print("No changes to apply")
            return None
        
        # Create a copy of the original document
        with tempfile.NamedTemporaryFile(delete=False, suffix='.odt') as temp_file:
            temp_path = temp_file.name
        
        # Save a copy of the original document
        original_doc.save(temp_path)
        
        # Reload the copy
        modified_doc = load(temp_path)
        
        # Get all paragraphs
        paragraphs = modified_doc.getElementsByType(text.P)
        
        # Apply changes with annotations
        for change in changes:
            original = change.get("original", "")
            corrected = change.get("corrected", "")
            
            if original == corrected:
                continue
                
            # Find the paragraph containing this text
            for para in paragraphs:
                para_text = teletype.extractText(para)
                if original in para_text:
                    # Clear the paragraph
                    while len(para) > 0:
                        para.removeChild(para.childNodes[0])
                    
                    # Add the corrected text with annotation
                    new_text = para_text.replace(original, f"[{original} → {corrected}]")
                    para.addText(new_text)
                    break
        
        # Save the modified document
        modified_doc.save(output_path)
        
        # Clean up
        try:
            os.unlink(temp_path)
        except:
            pass
        
        return output_path
    
    except Exception as e:
        print(f"Error saving ODT with changes: {e}")
        return None

def create_docx_with_changes(file_path, corrections, timestamp):
    """Create a new DOCX document with changes highlighted.
    
    Args:
        file_path: Path to the original file
        corrections: Dictionary containing correction data
        timestamp: Timestamp string for the output file
        
    Returns:
        Path to the saved document
    """
    if not DOCX_AVAILABLE:
        return None
        
    try:
        # Create output path
        base_path = os.path.splitext(file_path)[0]
        output_path = f"{base_path}_corrected_{timestamp}.docx"
        
        # Create a new document
        doc = Document()
        
        # Add a title
        doc.add_heading('Grammar Corrections', 0)
        
        # Add a summary paragraph
        doc.add_paragraph(f'Corrections for {os.path.basename(file_path)} generated on {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        
        # Get the original and corrected text
        original_text = corrections.get("original_text", "")
        corrected_text = corrections.get("corrected_text", "")
        
        # Add section for original text
        doc.add_heading('Original Text', level=1)
        doc.add_paragraph(original_text)
        
        # Add section for corrected text
        doc.add_heading('Corrected Text', level=1)
        doc.add_paragraph(corrected_text)
        
        # Add section for individual changes
        doc.add_heading('Detailed Changes', level=1)
        
        # Get the changes from the corrections
        changes = corrections.get("changes", [])
        
        if changes:
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            
            # Add header row
            header_cells = table.rows[0].cells
            header_cells[0].text = 'Original'
            header_cells[1].text = 'Correction'
            header_cells[2].text = 'Explanation'
            
            # Add changes
            for change in changes:
                row_cells = table.add_row().cells
                row_cells[0].text = change.get("original", "")
                row_cells[1].text = change.get("corrected", "")
                row_cells[2].text = change.get("explanation", "")
        else:
            doc.add_paragraph('No specific changes found.')
        
        # Save the document
        doc.save(output_path)
        
        return output_path
    
    except Exception as e:
        print(f"Error creating DOCX with changes: {e}")
        return None

def create_odt_with_changes(file_path, corrections, timestamp):
    """Create a new ODT document with changes highlighted.
    
    Args:
        file_path: Path to the original file
        corrections: Dictionary containing correction data
        timestamp: Timestamp string for the output file
        
    Returns:
        Path to the saved document
    """
    if not ODT_AVAILABLE:
        return None
        
    try:
        # Create output path
        base_path = os.path.splitext(file_path)[0]
        output_path = f"{base_path}_corrected_{timestamp}.odt"
        
        # Create a new document
        from odf.opendocument import OpenDocumentText
        from odf.style import Style, TextProperties, ParagraphProperties, TableProperties, TableRowProperties, TableCellProperties
        from odf.text import H, P, Span
        from odf.table import Table, TableColumn, TableRow, TableCell
        
        doc = OpenDocumentText()
        
        # Create styles
        heading_style = Style(name="Heading", family="paragraph")
        heading_style.addElement(TextProperties(attributes={'fontsize':"24pt", 'fontweight':"bold"}))
        doc.styles.addElement(heading_style)
        
        heading1_style = Style(name="Heading 1", family="paragraph")
        heading1_style.addElement(TextProperties(attributes={'fontsize':"18pt", 'fontweight':"bold"}))
        doc.styles.addElement(heading1_style)
        
        table_style = Style(name="Table", family="table")
        table_style.addElement(TableProperties(attributes={'width':"100%", 'align':"center"}))
        doc.styles.addElement(table_style)
        
        table_cell_style = Style(name="TableCell", family="table-cell")
        table_cell_style.addElement(TableCellProperties(attributes={'padding':"0.1cm", 'border':"0.05pt solid #000000"}))
        doc.styles.addElement(table_cell_style)
        
        # Add a title
        h = H(outlinelevel=1, stylename=heading_style, text="Grammar Corrections")
        doc.text.addElement(h)
        
        # Add a summary paragraph
        p = P(text=f'Corrections for {os.path.basename(file_path)} generated on {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        doc.text.addElement(p)
        
        # Get the original and corrected text
        original_text = corrections.get("original_text", "")
        corrected_text = corrections.get("corrected_text", "")
        
        # Add section for original text
        h = H(outlinelevel=2, stylename=heading1_style, text="Original Text")
        doc.text.addElement(h)
        p = P(text=original_text)
        doc.text.addElement(p)
        
        # Add section for corrected text
        h = H(outlinelevel=2, stylename=heading1_style, text="Corrected Text")
        doc.text.addElement(h)
        p = P(text=corrected_text)
        doc.text.addElement(p)
        
        # Add section for individual changes
        h = H(outlinelevel=2, stylename=heading1_style, text="Detailed Changes")
        doc.text.addElement(h)
        
        # Get the changes from the corrections
        changes = corrections.get("changes", [])
        
        if changes:
            table = Table(stylename=table_style)
            
            # Define columns
            table.addElement(TableColumn())
            table.addElement(TableColumn())
            table.addElement(TableColumn())
            
            # Add header row
            tr = TableRow()
            
            tc = TableCell(stylename=table_cell_style)
            tc.addElement(P(text="Original"))
            tr.addElement(tc)
            
            tc = TableCell(stylename=table_cell_style)
            tc.addElement(P(text="Correction"))
            tr.addElement(tc)
            
            tc = TableCell(stylename=table_cell_style)
            tc.addElement(P(text="Explanation"))
            tr.addElement(tc)
            
            table.addElement(tr)
            
            # Add changes
            for change in changes:
                tr = TableRow()
                
                tc = TableCell(stylename=table_cell_style)
                tc.addElement(P(text=change.get("original", "")))
                tr.addElement(tc)
                
                tc = TableCell(stylename=table_cell_style)
                tc.addElement(P(text=change.get("corrected", "")))
                tr.addElement(tc)
                
                tc = TableCell(stylename=table_cell_style)
                tc.addElement(P(text=change.get("explanation", "")))
                tr.addElement(tc)
                
                table.addElement(tr)
            
            doc.text.addElement(table)
        else:
            p = P(text='No specific changes found.')
            doc.text.addElement(p)
        
        # Save the document
        doc.save(output_path)
        
        return output_path
    
    except Exception as e:
        print(f"Error creating ODT with changes: {e}")
        return None