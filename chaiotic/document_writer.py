"""Module for saving documents with corrections."""

import os
import difflib
from typing import List, Dict, Any, Optional, Tuple
import time
import json

from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement, qn
from lxml import etree

from utils.text_utils import FuzzyMatcher, generate_full_text_from_corrections
from docx_utils import add_comments_to_paragraph, add_comment, create_comments_part, get_next_comment_id, create_comment_reference, create_comment, copy_run_formatting, add_comment_to_paragraph, _get_new_comment_id, _get_or_create_comments_part

def save_document(file_path: str, corrections: List[Dict[Any, Any]], original_doc=None, is_docx: bool = True) -> str:
    """
    Save the document with corrections as Word comments.
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
    output_path = f"{file_base}_corrected{file_ext}"
    
    print(f"Saving document with corrections to {output_path}")
    print(f"Number of corrections to apply: {len(corrections)}")
    
    # Create an instance of our optimized fuzzy matcher
    matcher = FuzzyMatcher()
    
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
            from docx import Document
            original_doc = Document(file_path)
        
        # Apply corrections directly to a copy of the document
        corrected_doc = apply_corrections_to_document(original_doc, corrections)
        
        # Save the corrected document
        try:
            corrected_doc.save(output_path)
            print(f"Successfully saved corrected document to {output_path}")
            return output_path
        except Exception as e:
            print(f"Error saving document: {str(e)}")
            # Try saving to a different location
            fallback_path = os.path.join(os.path.dirname(file_path), "corrected_output.docx")
            try:
                corrected_doc.save(fallback_path)
                print(f"Saved to alternate location: {fallback_path}")
                return fallback_path
            except Exception as e2:
                print(f"Failed to save document: {str(e2)}")
                raise
    else:
        # Handle ODT files - fallback to basic text format
        print("Note: Word comments are not supported for ODT files.")
        output_path = file_base + "_corrected.txt"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("CORRECTION SUGGESTIONS:\n\n")
            for i, correction in enumerate(corrections, 1):
                if isinstance(correction, dict):
                    original = correction.get('original', 'N/A')
                    corrected = correction.get('corrected', 'N/A')
                    explanation = correction.get('explanation', 'No explanation provided')
                    f.write(f"{i}. ORIGINAL: {original}\n")
                    f.write(f"   CORRECTED: {corrected}\n")
                    f.write(f"   EXPLANATION: {explanation}\n\n")
        
        print(f"Saved corrections to text file: {output_path}")
        return output_path

def save_correction_outputs(file_path, corrections, original_doc, is_docx=True):
    """Save correction outputs to files.
    
    Args:
        file_path: Original document file path
        corrections: List of corrections or dictionary with corrections
        original_doc: Original document object
        is_docx: Whether the document is a DOCX file
        
    Returns:
        Tuple of (json_path, text_path, doc_path)
    """
    import os
    import json
    import shutil
    
    # Generate output file paths
    base_name = os.path.splitext(file_path)[0]
    json_path = f"{base_name}_corrections.json"
    text_path = f"{base_name}_corrections.txt"
    doc_path = f"{base_name}_corrected{'.docx' if is_docx else '.odt'}"
    
    # Process corrections into a consistent format
    if isinstance(corrections, dict):
        # Handle dictionary format (either with 'corrections' key or direct corrections)
        if 'corrections' in corrections:
            corrections_list = corrections['corrections']
            corrected_full_text = corrections.get('corrected_full_text', '')
        else:
            # It's a direct dictionary of corrections
            corrections_list = [corrections]
            corrected_full_text = ''
    elif isinstance(corrections, list):
        # Already a list of corrections
        corrections_list = corrections
        corrected_full_text = ''
    else:
        # Unknown format
        print(f"Warning: Unknown corrections format: {type(corrections)}")
        corrections_list = []
        corrected_full_text = ''
    
    # Ensure all items in corrections_list are dictionaries
    valid_corrections = []
    for item in corrections_list:
        if isinstance(item, dict) and 'original' in item and 'corrected' in item:
            valid_corrections.append(item)
        else:
            print(f"Warning: Skipping invalid correction: {item}")
    
    corrections_list = valid_corrections
    print(f"Number of valid corrections to process: {len(corrections_list)}")
    
    # Generate full text from corrections if not provided
    if not corrected_full_text:
        corrected_full_text = generate_full_text_from_corrections(corrections_list)
    
    # Save JSON file with corrections
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({
            'corrections': corrections_list,
            'corrected_full_text': corrected_full_text
        }, f, indent=2, ensure_ascii=False)
    
    # Save text file with corrected content
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(corrected_full_text)
    
    # Save corrected document file
    try:
        from chaiotic.document_writer import apply_corrections_to_document
        corrected_doc = apply_corrections_to_document(original_doc, corrections_list, is_docx)
        corrected_doc.save(doc_path)
    except Exception as e:
        print(f"Error saving corrected document: {e}")
        print("Falling back to basic text document creation...")
        try:
            # Simple fallback to create a basic document
            if is_docx:
                from docx import Document
                doc = Document()
                for line in corrected_full_text.split('\n'):
                    doc.add_paragraph(line)
                doc.save(doc_path)
            else:
                # For ODT, just copy the text file
                shutil.copy(text_path, doc_path)
        except Exception as e2:
            print(f"Error creating basic document: {e2}")
    
    return json_path, text_path, doc_path

# Debugging why corrections are not applied
# Added print statements to verify corrections are being processed and applied

def apply_corrections_to_document(original_doc, corrections, is_docx=True):
    """Apply corrections to a document as tracked changes suggestions.
    
    Args:
        original_doc: Original document object
        corrections: List of corrections to apply
        is_docx: Whether the document is a DOCX file
        
    Returns:
        Corrected document object with tracked changes
    """
    if is_docx:
        from docx import Document
        from docx_utils import add_tracked_change_with_comment
        from docx.shared import RGBColor
        
        # Create a new document for the corrected version to ensure we're not modifying the original
        # This is important as python-docx sometimes has issues with deepcopy
        corrected_doc = Document()
        
        # Copy document styles from the original document to the new document
        # This needs to be done before creating paragraphs with those styles
        try:
            # Copy the base styles
            for style_id in original_doc.styles:
                if style_id not in corrected_doc.styles:
                    try:
                        style = original_doc.styles[style_id]
                        # We can't directly copy styles between documents,
                        # but we can ensure the style exists in the target document
                        if style.type == 1:  # Paragraph style
                            if style.name not in corrected_doc.styles:
                                corrected_doc.styles.add_style(style.name, style.type)
                    except Exception as e:
                        print(f"Warning: Could not copy style {style_id}: {e}")
        except Exception as e:
            print(f"Warning: Error copying document styles: {e}")
        
        # Copy the document properties and sections
        for section in original_doc.sections:
            new_section = corrected_doc.add_section()
            try:
                new_section.page_height = section.page_height
                new_section.page_width = section.page_width
                new_section.left_margin = section.left_margin
                new_section.right_margin = section.right_margin
                new_section.top_margin = section.top_margin
                new_section.bottom_margin = section.bottom_margin
                new_section.header_distance = section.header_distance
                new_section.footer_distance = section.footer_distance
            except Exception as e:
                print(f"Warning: Could not copy section properties: {e}")
        
        # Create a mapping of corrections by paragraph ID
        correction_by_paragraph = {}
        for correction in corrections:
            if 'id' in correction and correction['id'].startswith('p'):
                para_id = correction['id']
                if para_id not in correction_by_paragraph:
                    correction_by_paragraph[para_id] = []
                correction_by_paragraph[para_id].append(correction)
        
        # Debug: Print corrections mapped to paragraphs
        print("Corrections mapped to paragraphs:", correction_by_paragraph)
        
        # Apply corrections to paragraphs
        for i, original_para in enumerate(original_doc.paragraphs):
            para_id = f"p{i+1}"
            
            # Create a new paragraph in the corrected document
            new_para = corrected_doc.add_paragraph()
            
            # Copy the paragraph style
            if hasattr(original_para, 'style') and original_para.style:
                try:
                    # First try to copy by style name
                    style_name = original_para.style.name
                    if style_name in corrected_doc.styles:
                        new_para.style = style_name
                    else:
                        # If the style doesn't exist, try to get the built-in style with same name
                        try:
                            corrected_doc.styles.add_style(style_name, 1)  # 1 = paragraph style
                            new_para.style = style_name
                        except:
                            # Fall back to Normal style
                            new_para.style = 'Normal'
                except Exception as e:
                    print(f"Warning: Could not copy style for paragraph {para_id}: {e}")
            
            # Check if we have corrections for this paragraph
            para_corrections = correction_by_paragraph.get(para_id, [])
            
            if para_corrections and original_para.text.strip():
                # We have corrections for this paragraph
                print(f"Applying corrections to paragraph {para_id}: {original_para.text}")
                print("Corrections:", para_corrections)
                
                # Get the original text
                original_text = original_para.text
                processed_text = ""
                remaining_text = original_text
                
                # Apply each correction
                for correction in para_corrections:
                    if 'original' in correction and 'corrected' in correction:
                        original_part = correction.get('original', '')
                        corrected_part = correction.get('corrected', '')
                        explanation = correction.get('explanation', '')
                        
                        if original_part in remaining_text:
                            # Split the text at the correction point
                            before, match, after = remaining_text.partition(original_part)
                            
                            # Add the text before the correction
                            if before:
                                processed_text += before
                                new_para.add_run(before)
                            
                            # Add the correction with strikethrough and red text
                            del_run = new_para.add_run(original_part)
                            del_run.font.strike = True
                            
                            ins_run = new_para.add_run(corrected_part)
                            ins_run.bold = True
                            # Use RGBColor directly
                            ins_run.font.color.rgb = RGBColor(255, 0, 0)  # Red color
                            
                            # Add a comment with the explanation
                            comment_run = new_para.add_run(f" [{explanation}]")
                            comment_run.italic = True
                            comment_run.font.color.rgb = RGBColor(128, 128, 128)  # Gray
                            
                            # Update the remaining text for the next correction
                            processed_text += corrected_part
                            remaining_text = after
                        else:
                            print(f"WARNING: Could not find text '{original_part}' in paragraph.")
                
                # Add any remaining text
                if remaining_text:
                    new_para.add_run(remaining_text)
                    processed_text += remaining_text
                
                # Make sure we processed the full paragraph text
                if not processed_text.strip() and original_text.strip():
                    # If nothing was processed, just add the original text
                    new_para.text = original_text
                    print(f"Warning: No corrections applied to paragraph {para_id}, using original text")
            else:
                # No corrections, just copy the paragraph text and formatting
                if original_para.runs:
                    # Copy runs to preserve formatting
                    for run in original_para.runs:
                        new_run = new_para.add_run(run.text)
                        # Copy run formatting
                        try:
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            if hasattr(run, 'font') and hasattr(run.font, 'color') and run.font.color:
                                new_run.font.color.rgb = run.font.color.rgb
                        except Exception as e:
                            print(f"Warning: Could not copy run formatting: {e}")
                else:
                    # If no runs, just copy the text
                    new_para.text = original_para.text
        
        # Copy tables
        for table in original_doc.tables:
            try:
                new_table = corrected_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                # Copy table style if available
                if hasattr(table, 'style') and table.style:
                    try:
                        new_table.style = table.style
                    except:
                        print("Warning: Could not copy table style")
                
                # Copy cell contents
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        try:
                            # Copy each paragraph in the cell
                            for para in cell.paragraphs:
                                cell_para = new_table.rows[i].cells[j].add_paragraph()
                                cell_para.text = para.text
                                # Try to copy paragraph style
                                if hasattr(para, 'style') and para.style:
                                    try:
                                        cell_para.style = para.style.name
                                    except:
                                        pass
                        except Exception as e:
                            # Fall back to simple text copy
                            new_table.rows[i].cells[j].text = cell.text
                            print(f"Warning: Error copying table cell content: {e}")
            except Exception as e:
                print(f"Warning: Could not copy table: {e}")
        
        return corrected_doc
    
    else:
        # For ODT files, we don't have direct editing capability
        return original_doc

from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
import uuid
from datetime import datetime

# Define namespace map similar to what we were trying to import with NS
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'comments': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
}

def ensure_comments_part_exists(doc):
    """Ensure the document has a comments part.
    
    Args:
        doc: The docx Document object
        
    Returns:
        The comments part
    """
    # Check if document has a main document part
    if not hasattr(doc, 'part'):
        raise ValueError("Document doesn't have a main document part")
    
    # Try to get existing comments part
    try:
        for rel in doc.part.rels.values():
            if rel.reltype == NS['comments']:
                return rel.target
    except:
        pass
    
    # Create a new comments part if it doesn't exist
    try:
        from docx.parts.comments import CommentsXml
        comments_content = CommentsXml.new()
        comments_part = doc.part.package.create_part('/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml', comments_content)
        
        # Add a relationship to the comments part
        rel_id = doc.part.relate_to(comments_part, NS['comments'], is_external=False)
        
        # Add comments reference to document settings
        settings = doc.settings
        if settings is None:
            settings = doc.add_settings()
        
        # Return the comments part
        return comments_part
    except Exception as e:
        print(f"Error creating comments part: {e}")
        raise

def add_comment_to_document(doc, paragraph, comment_text, author="Grammatik-Prüfung", initials="GP"):
    """Add a comment to a paragraph in a document.
    
    Args:
        doc: The docx Document object
        paragraph: The paragraph to add the comment to
        comment_text: The text of the comment
        author: The author of the comment
        initials: The initials of the author
        
    Returns:
        The comment ID
    """
    # Get or create comments part
    comments_part = ensure_comments_part_exists(doc)
    
    # Create a unique comment ID
    comment_id = str(len(comments_part.xpath('//w:comment')) + 1)
    
    # Create a comment element
    comment = OxmlElement('w:comment')
    comment.set(qn('w:id'), comment_id)
    comment.set(qn('w:author'), author)
    comment.set(qn('w:initials'), initials)
    comment.set(qn('w:date'), datetime.now().isoformat())
    
    # Add the comment text as a paragraph
    comment_para = OxmlElement('w:p')
    comment_r = OxmlElement('w:r')
    comment_text_element = OxmlElement('w:t')
    comment_text_element.text = comment_text
    comment_r.append(comment_text_element)
    comment_para.append(comment_r)
    comment.append(comment_para)
    
    # Add the comment to the comments part
    comments_element = comments_part.getroot()
    comments_element.append(comment)
    
    # Add a comment reference to the paragraph
    run = paragraph.add_run()
    comment_reference = OxmlElement('w:commentReference')
    comment_reference.set(qn('w:id'), comment_id)
    run._element.append(comment_reference)
    
    return comment_id

def add_tracked_change_with_comment(doc, paragraph, original_text, corrected_text, explanation):
    """Add a tracked change and a comment to a paragraph in a DOCX document.

    Args:
        doc: The docx Document object
        paragraph: The paragraph to modify.
        original_text: The original text to replace.
        corrected_text: The corrected text to insert.
        explanation: The explanation for the change, added as a comment.
    """
    # Setup tracked changes in document settings if not already set
    settings = doc.settings
    if settings is None:
        settings = doc.add_settings()
    
    track_revisions = settings._element.find(qn('w:trackRevisions'))
    if track_revisions is None:
        track_revisions = OxmlElement('w:trackRevisions')
        settings._element.append(track_revisions)
    
    # Get the paragraph's XML
    p_xml = paragraph._element
    
    # Find the run containing the original text
    for run in paragraph.runs:
        if original_text in run.text:
            # Split the run into three parts: before, the match, and after
            before, match, after = run.text.partition(original_text)
            
            # Clear the original run's text
            run.text = before
            
            # Add the <w:del> tag for the original text
            del_run = paragraph.add_run()
            del_element = OxmlElement('w:del')
            del_element.set(qn('w:author'), "Correction")
            del_element.set(qn('w:date'), datetime.now().isoformat())
            del_element.set(qn('w:id'), str(uuid.uuid4())[:8])
            del_run._element.append(del_element)
            
            del_text = OxmlElement('w:t')
            if match.startswith(' ') or match.endswith(' '):
                del_text.set(qn('xml:space'), 'preserve')
            del_text.text = match
            del_element.append(del_text)
            
            # Add the <w:ins> tag for the corrected text
            ins_run = paragraph.add_run()
            ins_element = OxmlElement('w:ins')
            ins_element.set(qn('w:author'), "Correction")
            ins_element.set(qn('w:date'), datetime.now().isoformat())
            ins_element.set(qn('w:id'), str(uuid.uuid4())[:8])
            ins_run._element.append(ins_element)
            
            ins_text = OxmlElement('w:t')
            if corrected_text.startswith(' ') or corrected_text.endswith(' '):
                ins_text.set(qn('xml:space'), 'preserve')
            ins_text.text = corrected_text
            ins_element.append(ins_text)
            
            # Add the 'after' part if it exists
            if after:
                after_run = paragraph.add_run(after)
                after_run.bold = run.bold
                after_run.italic = run.italic
                after_run.underline = run.underline
            
            # Add a comment with the explanation
            add_comment_to_document(doc, paragraph, explanation)
            
            return True
    
    return False