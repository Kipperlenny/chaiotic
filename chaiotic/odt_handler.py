"""Module for handling ODT document operations."""

import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List, Dict, Any, Union

# Common XML namespaces for ODT documents
ODT_NAMESPACES = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'xlink': 'http://www.w3.org/1999/xlink',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
    'manifest': 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'
}

def register_namespaces():
    """Register ODT namespaces for XML parsing."""
    for prefix, uri in ODT_NAMESPACES.items():
        ET.register_namespace(prefix, uri)

def apply_corrections_to_odt(file_path: str, corrections: List[Dict[str, Any]], 
                            output_path: str) -> str:
    """Apply corrections to an ODT document with tracked changes.
    
    Args:
        file_path: Path to the original ODT document
        corrections: List of correction dictionaries
        output_path: Path to save the corrected document
        
    Returns:
        Path to the saved document
    """
    # Register namespaces first
    register_namespaces()
    
    # Try to use lxml for better XML handling if available
    try:
        import lxml.etree as LET
        print("Using lxml for better XML handling")
        return apply_corrections_with_lxml(file_path, corrections, output_path)
    except ImportError:
        print("lxml not available, falling back to standard ElementTree")
        return apply_corrections_with_elementtree(file_path, corrections, output_path)
    except Exception as e:
        print(f"Error using lxml: {e}")
        return apply_corrections_with_elementtree(file_path, corrections, output_path)

def apply_corrections_with_lxml(file_path: str, corrections: List[Dict[str, Any]], 
                               output_path: str) -> str:
    """Apply corrections to an ODT document using lxml library."""
    import lxml.etree as LET
    
    try:
        # Extract ODT content
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extract all files
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Parse content.xml
            content_file = os.path.join(temp_dir, 'content.xml')
            
            # Make sure namespaces are properly registered
            parser = LET.XMLParser(remove_blank_text=True)
            tree = LET.parse(content_file, parser)
            root = tree.getroot()
            
            # Get or create tracked-changes element with proper namespace
            nsmap = root.nsmap.copy()  # Copy to avoid modifying the original
            
            # Ensure required namespaces are present
            if 'text' not in nsmap:
                nsmap['text'] = ODT_NAMESPACES['text']
            if 'dc' not in nsmap:
                nsmap['dc'] = ODT_NAMESPACES['dc']
            if 'office' not in nsmap:
                nsmap['office'] = ODT_NAMESPACES['office']
                
            # Remove any problematic namespaces
            if any(ns.startswith('officeooo') for ns in nsmap):
                cleaned_nsmap = {k: v for k, v in nsmap.items() 
                                if not k.startswith('officeooo')}
                nsmap = cleaned_nsmap
            
            # Find text:tracked-changes element or create it
            office_text = None
            for child in root.iter('{' + nsmap['office'] + '}text'):
                office_text = child
                break
            
            if office_text is None:
                print("ERROR: office:text element not found in the document")
                # Create a basic fallback document
                return create_fallback_document(corrections, output_path)
            
            # Look for tracked-changes element
            tracked_changes = None
            for child in office_text.iterchildren():
                if child.tag == '{' + nsmap['text'] + '}tracked-changes':
                    tracked_changes = child
                    break
            
            # Create tracked-changes element if not found
            if tracked_changes is None:
                tracked_changes = LET.Element('{' + nsmap['text'] + '}tracked-changes', nsmap=nsmap)
                # Important: insert tracked-changes as the FIRST child of office:text
                office_text.insert(0, tracked_changes)
                print("Created new tracked-changes element")
            
            # Process corrections
            change_id = 1  # Initialize a unique change ID counter
            for correction in corrections:
                original = correction.get('original', '')
                corrected = correction.get('corrected', '')
                
                if not original or original == corrected:
                    continue
                
                print(f"\n=== Processing correction: '{original}' -> '{corrected}' ===")
                
                # Determine exactly what changed (for partial text changes)
                if len(original) > 2 and len(corrected) > 2:
                    # Find the common prefix and suffix
                    # This allows us to track only the specific characters that changed
                    i = 0
                    while i < min(len(original), len(corrected)) and original[i] == corrected[i]:
                        i += 1
                    
                    j = 1
                    while j <= min(len(original), len(corrected)) and original[-j] == corrected[-j]:
                        j += 1
                    
                    if i > 0 and j > 1:
                        prefix = original[:i]
                        suffix = original[-j+1:] if j > 1 else ""
                        orig_middle = original[i:-j+1] if j > 1 else original[i:]
                        corr_middle = corrected[i:-j+1] if j > 1 else corrected[i:]
                        
                        if orig_middle != corr_middle and (orig_middle or corr_middle):
                            # Only track the specific part that changed
                            print(f"Tracking specific change: '{orig_middle}' -> '{corr_middle}'")
                            orig_to_track = orig_middle
                            corr_to_track = corr_middle
                        else:
                            # Fall back to tracking the whole thing
                            orig_to_track = original
                            corr_to_track = corrected
                    else:
                        # Fall back to tracking the whole thing
                        orig_to_track = original
                        corr_to_track = corrected
                else:
                    # For short text, just track the whole thing
                    orig_to_track = original
                    corr_to_track = corrected
                
                # Create two separate region IDs for deletion and insertion
                del_change_id_str = f"ctd{change_id}"
                ins_change_id_str = f"cti{change_id}"
                change_id += 1  # Increment for next change
                
                # 1. Create changed-region for DELETION
                del_region = LET.Element('{' + nsmap['text'] + '}changed-region', nsmap=nsmap)
                del_region.set('{' + nsmap['text'] + '}id', del_change_id_str)
                del_region.set('{http://www.w3.org/XML/1998/namespace}id', del_change_id_str)
                
                # Add deletion element to the deletion region
                deletion = LET.SubElement(del_region, '{' + nsmap['text'] + '}deletion')
                
                # Add change-info element to deletion
                change_info = LET.SubElement(deletion, '{' + nsmap['office'] + '}change-info')
                
                # Add creator and date to change-info
                creator = LET.SubElement(change_info, '{' + nsmap['dc'] + '}creator')
                creator.text = "Chaiotic Grammar Checker"
                date = LET.SubElement(change_info, '{' + nsmap['dc'] + '}date')
                date.text = datetime.now().isoformat()
                
                # Add explanation as paragraph in change-info if available
                if 'explanation' in correction:
                    explanation_p = LET.SubElement(change_info, '{' + nsmap['text'] + '}p')
                    explanation_p.text = correction.get('explanation', '')
                
                # Add deleted text content as paragraph
                deleted_text = LET.SubElement(deletion, '{' + nsmap['text'] + '}p')
                deleted_text.text = orig_to_track
                
                # Add the deletion region to tracked-changes
                tracked_changes.append(del_region)
                
                # 2. Create a SEPARATE changed-region for INSERTION
                ins_region = LET.Element('{' + nsmap['text'] + '}changed-region', nsmap=nsmap)
                ins_region.set('{' + nsmap['text'] + '}id', ins_change_id_str)
                ins_region.set('{http://www.w3.org/XML/1998/namespace}id', ins_change_id_str)
                
                # Add insertion element to the insertion region
                insertion = LET.SubElement(ins_region, '{' + nsmap['text'] + '}insertion')
                
                # Add change-info element to insertion
                change_info = LET.SubElement(insertion, '{' + nsmap['office'] + '}change-info')
                
                # Add creator and date to change-info
                creator = LET.SubElement(change_info, '{' + nsmap['dc'] + '}creator')
                creator.text = "Chaiotic Grammar Checker"
                date = LET.SubElement(change_info, '{' + nsmap['dc'] + '}date')
                date.text = datetime.now().isoformat()
                
                # Add explanation as paragraph in change-info if available
                if 'explanation' in correction:
                    explanation_p = LET.SubElement(change_info, '{' + nsmap['text'] + '}p')
                    explanation_p.text = correction.get('explanation', '')
                
                # Add inserted text content as paragraph
                inserted_text = LET.SubElement(insertion, '{' + nsmap['text'] + '}p')
                inserted_text.text = corr_to_track
                
                # Add the insertion region to tracked-changes
                tracked_changes.append(ins_region)
                
                # Find and update text in paragraphs
                find_and_mark_text_in_paragraphs(office_text, original, corrected, 
                                               orig_to_track, corr_to_track,
                                               del_change_id_str, ins_change_id_str, nsmap)
            
            # Remove any problematic attributes from the entire tree
            for element in root.iter():
                # Remove officeooo:rsid attributes which cause validation errors
                attrib_keys = list(element.attrib.keys())
                for attr in attrib_keys:
                    if 'officeooo:rsid' in attr or 'rsid' in attr:
                        del element.attrib[attr]
            
            # Validate all change-start elements have corresponding change-end elements
            # and all change IDs are properly defined
            change_ids = {elem.get('{' + nsmap['text'] + '}id') 
                         for elem in tracked_changes.findall('{' + nsmap['text'] + '}changed-region')}
            
            for elem in root.xpath('//*[local-name()="change-start"]'):
                change_id = elem.get('{' + nsmap['text'] + '}change-id')
                # Check if reference exists
                if change_id not in change_ids:
                    # If not, either fix or remove the reference
                    parent = elem.getparent()
                    if parent is not None:
                        parent.remove(elem)
                        
            for elem in root.xpath('//*[local-name()="change-end"]'):
                change_id = elem.get('{' + nsmap['text'] + '}change-id')
                # Check if reference exists
                if change_id not in change_ids:
                    # If not, either fix or remove the reference
                    parent = elem.getparent()
                    if parent is not None:
                        parent.remove(elem)
            
            # Write the modified XML back to the file
            tree.write(content_file, encoding='UTF-8', xml_declaration=True, pretty_print=True)
            
            # Create a new ODT file
            with zipfile.ZipFile(output_path, 'w') as new_odt:
                # Add mimetype first without compression
                mimetype_path = os.path.join(temp_dir, 'mimetype')
                if os.path.exists(mimetype_path):
                    new_odt.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
                
                # Add all other files
                for root_dir, _, files in os.walk(temp_dir):
                    for file in files:
                        if file != 'mimetype':
                            file_path_in_temp = os.path.join(root_dir, file)
                            arcname = os.path.relpath(file_path_in_temp, temp_dir)
                            new_odt.write(file_path_in_temp, arcname, compress_type=zipfile.ZIP_DEFLATED)
        
        return output_path
    
    except Exception as e:
        print(f"Error processing ODT with lxml: {e}")
        import traceback
        traceback.print_exc()
        return create_fallback_document(corrections, output_path)

def apply_corrections_with_elementtree(file_path: str, corrections: List[Dict[str, Any]], 
                                      output_path: str) -> str:
    """Apply corrections to an ODT document using standard ElementTree.
    
    Args:
        file_path: Path to the original ODT document
        corrections: List of correction dictionaries
        output_path: Path to save the corrected document
        
    Returns:
        Path to the saved document
    """
    # Register namespaces
    register_namespaces()
    
    # Extract ODT content
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract all files
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Parse content.xml
        content_file = os.path.join(temp_dir, 'content.xml')
        
        try:
            tree = ET.parse(content_file)
            root = tree.getroot()
            
            # Working with ElementTree is more challenging for complex XML manipulations
            # For simplicity, we'll add comments with corrections instead of tracked changes
            
            # Find all text paragraphs
            text_ns = ODT_NAMESPACES['text']
            office_ns = ODT_NAMESPACES['office']
            
            # Find the office:text element
            office_text = None
            for elem in root.iter(f"{{{office_ns}}}text"):
                office_text = elem
                break
            
            if office_text is None:
                print("Warning: office:text element not found")
                return create_fallback_document(corrections, output_path)
            
            # Create a comment at the beginning with corrections
            comment = ET.Comment("Grammar corrections by Chaiotic:")
            office_text.insert(0, comment)
            
            # For each paragraph, see if it contains text that needs correction
            for correction in corrections:
                original = correction.get('original', '')
                corrected = correction.get('corrected', '')
                
                if not original or original == corrected:
                    continue
                    
                # Add a comment about this correction
                comment_text = f"Correction: '{original}' -> '{corrected}'"
                comment = ET.Comment(comment_text)
                office_text.insert(0, comment)
                
                # Find paragraphs containing the text to correct
                for para in office_text.findall(f".//{{{text_ns}}}p"):
                    # Get all text content in this paragraph
                    para_text = "".join(para.itertext())
                    
                    if original in para_text:
                        # Add a note about the correction as an annotation
                        try:
                            # We cannot reliably use tracked changes with ElementTree
                            # so we'll just note corrections in comments
                            note_comment = ET.Comment(f"CORRECTION: '{original}' should be '{corrected}'")
                            para.insert(0, note_comment)
                        except Exception as e:
                            print(f"Error adding correction note: {e}")
            
            # Write the modified XML back to the file
            tree.write(content_file, encoding='UTF-8', xml_declaration=True)
            
            # Create a new ODT file
            with zipfile.ZipFile(output_path, 'w') as new_odt:
                # Add mimetype first without compression
                mimetype_path = os.path.join(temp_dir, 'mimetype')
                if os.path.exists(mimetype_path):
                    new_odt.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
                
                # Add all other files
                for root_dir, _, files in os.walk(temp_dir):
                    for file in files:
                        if file != 'mimetype':
                            file_path_in_temp = os.path.join(root_dir, file)
                            arcname = os.path.relpath(file_path_in_temp, temp_dir)
                            new_odt.write(file_path_in_temp, arcname, compress_type=zipfile.ZIP_DEFLATED)
            
            return output_path
            
        except Exception as e:
            print(f"Error processing ODT with ElementTree: {e}")
            import traceback
            traceback.print_exc()
            return create_fallback_document(corrections, output_path)

def create_fallback_document(corrections: List[Dict[str, Any]], output_path: str) -> str:
    """Create a simple text file with corrections when ODT processing fails.
    
    Args:
        corrections: List of correction dictionaries
        output_path: Original output path
        
    Returns:
        Path to the saved document
    """
    # Create a text file with corrections
    text_path = f"{output_path}.txt"
    try:
        with open(text_path, 'w', encoding='utf-8') as f:
            f.write("GRAMMAR CORRECTIONS:\n\n")
            for i, corr in enumerate(corrections, 1):
                if isinstance(corr, dict):
                    f.write(f"{i}. Original: {corr.get('original', '')}\n")
                    f.write(f"   Corrected: {corr.get('corrected', '')}\n")
                    if 'explanation' in corr:
                        f.write(f"   Explanation: {corr['explanation']}\n")
                elif isinstance(corr, str):
                    f.write(f"{i}. Correction: {corr}\n")
                f.write("\n")
    except Exception as e:
        print(f"Error creating fallback text file: {e}")
        return None
        
    return text_path

def create_structured_odt(content: Union[str, List[Dict[str, Any]]], 
                         output_path: str) -> str:
    """Create a new ODT document with structured content.
    
    Args:
        content: Text content or structured content list
        output_path: Path to save the document
        
    Returns:
        Path to the saved document
    """
    try:
        # Try to use odfpy library if available
        try:
            from odf.opendocument import OpenDocumentText
            from odf.style import Style, TextProperties, ParagraphProperties
            from odf.text import H, P, Span
            
            textdoc = OpenDocumentText()
            
            # Add styles
            heading_style = Style(name="Heading", family="paragraph")
            heading_style.addElement(TextProperties(fontsize="16pt", fontweight="bold"))
            textdoc.styles.addElement(heading_style)
            
            # Process content
            if isinstance(content, str):
                # Simple text content
                paragraphs = content.split('\n\n')
                for para_text in paragraphs:
                    if not para_text.strip():
                        continue
                    p = P(text=para_text)
                    textdoc.text.addElement(p)
            else:
                # Structured content
                for item in content:
                    if isinstance(item, dict):
                        item_type = item.get('type', 'paragraph')
                        item_text = item.get('text', '')
                        item_level = item.get('level', 1)
                        
                        if not item_text.strip():
                            continue
                            
                        if item_type == 'heading':
                            h = H(outlinelevel=item_level, text=item_text)
                            textdoc.text.addElement(h)
                        else:
                            p = P(text=item_text)
                            textdoc.text.addElement(p)
            
            # Save the document
            textdoc.save(output_path)
            return output_path
            
        except ImportError:
            print("odfpy library not found, using alternative method")
            return create_basic_odt(content, output_path)
    except Exception as e:
        print(f"Error creating structured ODT: {e}")
        # Create a simple text file as fallback
        text_path = f"{output_path}.txt"
        try:
            with open(text_path, 'w', encoding='utf-8') as f:
                if isinstance(content, str):
                    f.write(content)
                else:
                    for item in content:
                        if isinstance(item, dict):
                            f.write(item.get('text', '') + '\n\n')
                        else:
                            f.write(str(item) + '\n\n')
        except Exception as text_err:
            print(f"Error creating text fallback: {text_err}")
            return None
            
        return text_path

def create_basic_odt(content: Union[str, List[Dict[str, Any]]], output_path: str) -> str:
    """Create a basic ODT document without specialized libraries.
    
    Args:
        content: Text content or structured content list
        output_path: Path to save the document
        
    Returns:
        Path to the saved document
    """
    # Register namespaces
    register_namespaces()
    
    # Convert content to string if it's structured
    if not isinstance(content, str):
        text_content = ""
        for item in content:
            if isinstance(item, dict):
                item_type = item.get('type', 'paragraph')
                item_text = item.get('text', '')
                
                if item_type == 'heading':
                    text_content += f"# {item_text}\n\n"
                else:
                    text_content += f"{item_text}\n\n"
            else:
                text_content += f"{item}\n\n"
        content = text_content
    
    # Create a simple ODT file structure
    with tempfile.TemporaryDirectory() as temp_dir:
        # Create META-INF directory and mimetype
        os.makedirs(os.path.join(temp_dir, 'META-INF'))
        
        # Process content into paragraphs
        paragraphs = []
        for para in content.split('\n\n'):
            para = para.strip()
            if not para:
                continue
                
            if para.startswith('# '):
                # Heading
                heading_text = para[2:].strip()
                paragraphs.append({
                    'type': 'heading',
                    'text': heading_text
                })
            else:
                # Regular paragraph
                paragraphs.append({
                    'type': 'paragraph', 
                    'text': para
                })
        
        # Create content.xml with paragraphs
        with open(os.path.join(temp_dir, 'content.xml'), 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<office:document-content ')
            
            # Add all namespaces
            for prefix, uri in ODT_NAMESPACES.items():
                f.write(f'xmlns:{prefix}="{uri}" ')
            
            f.write('office:version="1.2">\n')
            f.write('  <office:body>\n')
            f.write('    <office:text>\n')
            
            # Add paragraphs
            for para in paragraphs:
                if para['type'] == 'heading':
                    f.write(f'      <text:h text:style-name="Heading_20_1" text:outline-level="1">{para["text"]}</text:h>\n')
                else:
                    f.write(f'      <text:p text:style-name="Text_20_body">{para["text"]}</text:p>\n')
            
            f.write('    </office:text>\n')
            f.write('  </office:body>\n')
            f.write('</office:document-content>')
        
        # Create mimetype file
        with open(os.path.join(temp_dir, 'mimetype'), 'w', encoding='utf-8') as f:
            f.write('application/vnd.oasis.opendocument.text')
        
        # Create META-INF/manifest.xml
        with open(os.path.join(temp_dir, 'META-INF', 'manifest.xml'), 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">\n')
            f.write(' <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>\n')
            f.write(' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>\n')
            f.write(' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>\n')
            f.write(' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>\n')
            f.write(' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="META-INF/manifest.xml"/>\n')
            f.write('</manifest:manifest>')
        
        # Create styles.xml
        with open(os.path.join(temp_dir, 'styles.xml'), 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<office:document-styles ')
            for prefix, uri in ODT_NAMESPACES.items():
                if prefix in ['office', 'style', 'text', 'fo']:
                    f.write(f'xmlns:{prefix}="{uri}" ')
            f.write('office:version="1.2">\n')
            f.write('  <office:styles>\n')
            f.write('    <style:style style:name="Standard" style:family="paragraph" style:class="text"/>\n')
            f.write('    <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text">\n')
            f.write('      <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.212cm"/>\n')
            f.write('    </style:paragraph-properties>\n')
            f.write('    <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:class="text">\n')
            f.write('      <style:text-properties fo:font-size="14pt" fo:font-weight="bold"/>\n')
            f.write('    </style:style>\n')
            f.write('    <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:class="text">\n')
            f.write('      <style:text-properties fo:font-size="18pt" fo:font-weight="bold"/>\n')
            f.write('    </style:style>\n')
            f.write('  </office:styles>\n')
            f.write('</office:document-styles>')
        
        # Create meta.xml
        with open(os.path.join(temp_dir, 'meta.xml'), 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<office:document-meta ')
            for prefix, uri in ODT_NAMESPACES.items():
                if prefix in ['office', 'meta', 'dc']:
                    f.write(f'xmlns:{prefix}="{uri}" ')
            f.write('office:version="1.2">\n')
            f.write('  <office:meta>\n')
            f.write('    <dc:title>Created by Chaiotic</dc:title>\n')
            f.write(f'    <dc:date>{datetime.now().isoformat()}</dc:date>\n')
            f.write('  </office:meta>\n')
            f.write('</office:document-meta>')
        
        # Create the ODT file (which is a ZIP file)
        with zipfile.ZipFile(output_path, 'w') as zip_file:
            # Add mimetype first without compression
            zip_file.write(os.path.join(temp_dir, 'mimetype'), 'mimetype', compress_type=zipfile.ZIP_STORED)
            
            # Add other files with compression
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    if file != 'mimetype':
                        file_path_in_temp = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path_in_temp, temp_dir)
                        zip_file.write(file_path_in_temp, arcname, compress_type=zipfile.ZIP_DEFLATED)
    
    return output_path

def find_text_in_complex_paragraph(paragraph, text_to_find):
    """Find text within a complex paragraph with nested elements.
    
    Args:
        paragraph: The paragraph element to search in
        text_to_find: The text to search for
        
    Returns:
        Tuple of (found, element, text_offset) or (False, None, -1)
    """
    # First get the normalized flattened text and check if the text is present
    full_text = ""
    
    try:
        # Use xpath to get all text nodes from the paragraph
        full_text = ''.join(paragraph.xpath('.//text()'))
        
        # Check if the search text is in the full text
        if text_to_find not in full_text:
            # Also check for text followed by punctuation that might be in a different element
            for punct in ['.', ',', ';', ':', '!', '?']:
                if text_to_find + punct in full_text:
                    print(f"DEBUG: Found '{text_to_find}' with trailing punctuation '{punct}'")
                    print(f"This might require manual verification.")
                    return False, None, -1
            
            # If not found at all, return early
            return False, None, -1
            
        # Try to find the element containing the text
        for elem in paragraph.iter():
            if elem.text and text_to_find in elem.text:
                return True, elem, elem.text.find(text_to_find)
            elif elem.tail and text_to_find in elem.tail:
                return True, elem, elem.tail.find(text_to_find)
        
        # If we can't find it in a specific element but we know it's in the full text,
        # it may be split across elements or require special handling
        print(f"DEBUG: Text '{text_to_find}' found in flattened text but not in any single element")
        return True, "COMPLEX_NESTED", 0
        
    except Exception as e:
        print(f"Error in find_text_in_complex_paragraph: {e}")
        import traceback
        traceback.print_exc()
        return False, None, -1

def handle_complex_nested_element(paragraph, original, corrected, 
                                 del_change_id, ins_change_id, nsmap):
    """Handle text replacements in complex nested element structures.
    
    Args:
        paragraph: The paragraph element
        original: Original text to replace
        corrected: Corrected text
        del_change_id: Change ID for deletion
        ins_change_id: Change ID for insertion
        nsmap: XML namespaces
        
    Returns:
        True if successful, False otherwise
    """
    import lxml.etree as LET
    
    try:
        # Debug the paragraph structure
        para_str = LET.tostring(paragraph, encoding='unicode', pretty_print=True)
        print(f"DEBUG: Original paragraph XML:\n{para_str}")
        
        # Get full flattened text from paragraph
        full_text = ''.join(paragraph.xpath('.//text()'))
        print(f"DEBUG: Full flattened text: '{full_text}'")
        
        # Verify the original text is in the flattened text
        if original not in full_text:
            print(f"DEBUG: Original text '{original}' not found in flattened text")
            return False
        
        print(f"DEBUG: Original text '{original}' found in flattened text at position {full_text.find(original)}")
        
        # Try the direct approach: create a new paragraph with the structure we want
        new_para = LET.Element(paragraph.tag, paragraph.attrib)
        
        # Split the full text at the original text
        parts = full_text.split(original, 1)
        
        # Text before our target
        if parts[0]:
            before_span = LET.SubElement(new_para, '{' + nsmap['text'] + '}span')
            before_span.text = parts[0]
        
        # Container for our tracked changes
        change_span = LET.SubElement(new_para, '{' + nsmap['text'] + '}span')
        
        # Add deletion and insertion markers ONLY - don't include the text
        del_start = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-start')
        del_start.set('{' + nsmap['text'] + '}change-id', del_change_id)
        
        del_end = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-end') 
        del_end.set('{' + nsmap['text'] + '}change-id', del_change_id)
        
        ins_start = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-start')
        ins_start.set('{' + nsmap['text'] + '}change-id', ins_change_id)
        
        ins_end = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-end')
        ins_end.set('{' + nsmap['text'] + '}change-id', ins_change_id)
        
        # Text after our target
        if len(parts) > 1 and parts[1]:
            after_span = LET.SubElement(new_para, '{' + nsmap['text'] + '}span')
            after_span.text = parts[1]
        
        # Replace the original paragraph
        parent = paragraph.getparent()
        if parent is not None:
            idx = list(parent).index(paragraph)
            parent.remove(paragraph)
            parent.insert(idx, new_para)
            
            print(f"DEBUG: Successfully rebuilt paragraph with tracked changes for '{original}' -> '{corrected}'")
            print(f"DEBUG: New paragraph XML:\n{LET.tostring(new_para, encoding='unicode', pretty_print=True)}")
            return True
        
        return False
    except Exception as e:
        print(f"Error in complex nested element handler: {e}")
        import traceback
        traceback.print_exc()
        
        # Try a fallback approach for special case of first letter capitalization
        try:
            if original.lower() == corrected.lower() and original[0].isupper() != corrected[0].isupper():
                print(f"DEBUG: Attempting special case handling for capitalization change")
                return handle_capitalization_change(paragraph, original, corrected, del_change_id, ins_change_id, nsmap)
        except Exception as e2:
            print(f"Error in capitalization fallback: {e2}")
        
        return False

def handle_capitalization_change(paragraph, original, corrected, del_change_id, ins_change_id, nsmap):
    """Special handler for capitalization changes.
    
    Args:
        paragraph: The paragraph element
        original: Original text with capitalization issue
        corrected: Corrected text
        del_change_id: Change ID for deletion
        ins_change_id: Change ID for insertion
        nsmap: XML namespaces
        
    Returns:
        True if successful, False otherwise
    """
    import lxml.etree as LET
    
    try:
        # Get the first span or text node that might contain our first character
        text_nodes = []
        
        # First, check the paragraph text
        if paragraph.text:
            text_nodes.append((paragraph, 'text', paragraph.text))
        
        # Then check all child elements
        for elem in paragraph.iter():
            if elem is not paragraph and elem.text:
                text_nodes.append((elem, 'text', elem.text))
            if elem.tail:
                text_nodes.append((elem, 'tail', elem.tail))
        
        # Look for the first character of our original text
        for node, attr, text in text_nodes:
            if original[0] in text:
                # Found it! Now replace just that character
                char_pos = text.find(original[0])
                
                if attr == 'text':
                    # It's in the element's text
                    before_text = text[:char_pos]
                    node.text = before_text
                    
                    # Create change markers
                    change_span = LET.SubElement(node, '{' + nsmap['text'] + '}span')
                    
                    del_start = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-start')
                    del_start.set('{' + nsmap['text'] + '}change-id', del_change_id)
                    
                    del_end = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-end')
                    del_end.set('{' + nsmap['text'] + '}change-id', del_change_id)
                    
                    ins_start = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-start')
                    ins_start.set('{' + nsmap['text'] + '}change-id', ins_change_id)
                    
                    ins_end = LET.SubElement(change_span, '{' + nsmap['text'] + '}change-end')
                    ins_end.set('{' + nsmap['text'] + '}change-id', ins_change_id)
                    
                    # Add corrected character
                    # ins_end.tail = corrected[0]
                    
                    # Add remainder span if needed
                    if len(text) > char_pos + 1:
                        remainder = LET.SubElement(node, '{' + nsmap['text'] + '}span')
                        remainder.text = text[char_pos + 1:]
                    
                    print(f"DEBUG: Successfully handled capitalization change in element text")
                    return True
                elif attr == 'tail':
                    # It's in the element's tail - more complex
                    before_text = text[:char_pos]
                    node.tail = before_text
                    
                    # Create a new element after this one
                    parent = node.getparent()
                    idx = list(parent).index(node)
                    
                    # Create container for changes
                    container = LET.Element('{' + nsmap['text'] + '}span')
                    
                    # Add change markers
                    del_start = LET.SubElement(container, '{' + nsmap['text'] + '}change-start')
                    del_start.set('{' + nsmap['text'] + '}change-id', del_change_id)
                    
                    del_end = LET.SubElement(container, '{' + nsmap['text'] + '}change-end')
                    del_end.set('{' + nsmap['text'] + '}change-id', del_change_id)
                    
                    ins_start = LET.SubElement(container, '{' + nsmap['text'] + '}change-start')
                    ins_start.set('{' + nsmap['text'] + '}change-id', ins_change_id)
                    
                    ins_end = LET.SubElement(container, '{' + nsmap['text'] + '}change-end')
                    ins_end.set('{' + nsmap['text'] + '}change-id', ins_change_id)
                    
                    # Add corrected character and remainder
                    # ins_end.tail = corrected[0] + text[char_pos + 1:]
                    
                    # Insert after current node
                    parent.insert(idx + 1, container)
                    
                    print(f"DEBUG: Successfully handled capitalization change in element tail")
                    return True
        
        return False
    except Exception as e:
        print(f"Error in capitalization handler: {e}")
        return False

def find_and_mark_text_in_paragraphs(office_text, original, corrected, 
                                orig_to_track, corr_to_track,
                                del_change_id_str, ins_change_id_str, nsmap):
    """Find text in paragraphs and mark with tracked changes.
    
    Args:
        office_text: The office:text element
        original: Original full text to find
        corrected: Corrected full text
        orig_to_track: Original text part to track (may be just the changed portion)
        corr_to_track: Corrected text part to track
        del_change_id_str: Change ID for deletion
        ins_change_id_str: Change ID for insertion
        nsmap: XML namespaces
    """
    import lxml.etree as LET
    
    # Add debug logs to help diagnose issues with nested text
    print(f"\n=== DEBUG: Marking text: '{original}' -> '{corrected}' ===")
    
    # Find and update text in paragraphs
    found_and_replaced = False
    
    for paragraph in office_text.findall('.//{' + nsmap['text'] + '}p'):
        try:
            # Skip paragraphs in tracked changes
            if paragraph.getparent().tag == '{' + nsmap['text'] + '}deletion' or \
               paragraph.getparent().tag == '{' + nsmap['text'] + '}insertion':
                continue
            
            # Get normalized paragraph text for debugging
            full_text = get_normalized_paragraph_text(paragraph)
            para_preview = full_text[:100] + ('...' if len(full_text) > 100 else '')
            print(f"Checking paragraph: '{para_preview}'")
            
            # Debug check for text with punctuation
            for punct in ['.', ',', ';', ':', '!', '?']:
                if original + punct in full_text:
                    print(f"DEBUG: Found '{original}' with punctuation '{punct}' - this may be why no change is needed")
            
            # Convert paragraph to string for debugging
            para_str = LET.tostring(paragraph, encoding='unicode')
            print(f"Paragraph XML (preview): '{para_str[:150]}...'")
            
            # Check if the paragraph contains our text
            if original not in full_text:
                print(f"DEBUG: '{original}' NOT found in this paragraph")
                continue
                
            print(f"DEBUG: Found '{original}' in paragraph!")
            
            # Use our improved text search function for complex paragraphs
            found, element, position = find_text_in_complex_paragraph(paragraph, original)
            
            if found:
                if element == "COMPLEX_NESTED":
                    print(f"DEBUG: Text spans multiple elements, using complex handler")
                    # Use the specialized handler for complex nested elements
                    if handle_complex_nested_element(paragraph, original, corrected,
                                                  del_change_id_str, ins_change_id_str, nsmap):
                        print(f"Successfully handled complex nested elements for '{original}'")
                        found_and_replaced = True
                    else:
                        print(f"Failed to handle complex nested elements for '{original}'")
                else:
                    print(f"DEBUG: Found in element type: {element.tag}")
                    
                    # Handle case when text is in element.text
                    if element.text and original in element.text:
                        print(f"DEBUG: Found in element.text: '{element.text}'")
                        # Split the text and add markers
                        parts = element.text.split(original, 1)
                        
                        # Save the element's tail
                        saved_tail = element.tail
                        element.text = parts[0]
                        
                        # Add deletion marker
                        change_start = LET.SubElement(
                            element, '{' + nsmap['text'] + '}change-start',
                            {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                        )
                        
                        change_end = LET.SubElement(
                            element, '{' + nsmap['text'] + '}change-end',
                            {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                        )
                        
                        # Add insertion marker
                        change_start = LET.SubElement(
                            element, '{' + nsmap['text'] + '}change-start',
                            {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                        )
                        
                        change_end = LET.SubElement(
                            element, '{' + nsmap['text'] + '}change-end',
                            {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                        )
                        
                        # DON'T add the text content here - it's already in the change element
                        
                        # Add remainder text
                        if len(parts) > 1 and parts[1]:
                            remainder_span = LET.SubElement(element, '{' + nsmap['text'] + '}span')
                            remainder_span.text = parts[1]
                        
                        # Restore tail
                        element.tail = saved_tail
                        print("DEBUG: Successfully inserted tracked change markers into element.text")
                        found_and_replaced = True
                            
                    # Handle case when text is in element.tail
                    elif element.tail and original in element.tail:
                        print(f"DEBUG: Found in element.tail: '{element.tail}'")
                        # Split the tail and add markers
                        parts = element.tail.split(original, 1)
                        
                        # Save the element's tail start
                        element.tail = parts[0]
                        
                        # Create a new element after this one to hold our change markers
                        parent = element.getparent()
                        idx = list(parent).index(element)
                        
                        # Create a container for our change markers
                        container = LET.Element('{' + nsmap['text'] + '}span')
                        
                        # Add deletion marker
                        change_start = LET.SubElement(
                            container, '{' + nsmap['text'] + '}change-start',
                            {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                        )
                        
                        change_end = LET.SubElement(
                            container, '{' + nsmap['text'] + '}change-end',
                            {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                        )
                        
                        # Add insertion marker
                        change_start = LET.SubElement(
                            container, '{' + nsmap['text'] + '}change-start',
                            {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                        )
                        
                        change_end = LET.SubElement(
                            container, '{' + nsmap['text'] + '}change-end',
                            {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                        )
                        
                        # DON'T add the text content here - it's already in the change element
                        
                        # Add tail remainder if any
                        if len(parts) > 1 and parts[1]:
                            container.tail = parts[1]
                        
                        # Insert our container after the current element
                        parent.insert(idx + 1, container)
                        print("DEBUG: Successfully inserted tracked change markers into element.tail")
                        found_and_replaced = True
                    
                    # Handle other complex cases
                    else:
                        print(f"Unhandled text position - using fallback approach")
                        
                        # Try using the complex element handler as a fallback
                        if handle_complex_nested_element(paragraph, original, corrected,
                                                       del_change_id_str, ins_change_id_str, nsmap):
                            print("Successfully used complex handler as fallback")
                            found_and_replaced = True
                        else:
                            print("Failed to handle with complex handler - text position unclear")
            else:
                print(f"DEBUG: '{original}' NOT found in this paragraph using element search")
                
                # Try the complex handler anyway as a last resort if we know the text is in there
                if original in full_text:
                    print("DEBUG: Trying complex handler as last resort")
                    if handle_complex_nested_element(paragraph, original, corrected, 
                                                  del_change_id_str, ins_change_id_str, nsmap):
                        print("Successfully used complex handler as last resort")
                        found_and_replaced = True
        
        except Exception as e:
            print(f"Error processing paragraph: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    # If we couldn't find the text in paragraphs using normal methods,
    # try one more approach for the entire document
    if not found_and_replaced:
        print("DEBUG: No successful replacement found, trying document-wide search...")
        
        # This is a last resort - try a more aggressive approach to find the text
        for element in office_text.xpath('//*'):
            # Skip tracked-changes elements
            if any(tag in element.tag for tag in ['tracked-changes', 'change-start', 'change-end']):
                continue
                
            try:
                full_text = ''.join(element.xpath('.//text()'))
                if original in full_text:
                    print(f"DEBUG: Found '{original}' in element {element.tag}")
                    if handle_complex_nested_element(element, original, corrected,
                                                  del_change_id_str, ins_change_id_str, nsmap):
                        print("Successfully handled using complex handler on container element")
                        return
            except Exception as e:
                print(f"Error in document-wide search: {e}")
                continue

def get_normalized_paragraph_text(paragraph):
    """Get all text from a paragraph with normalized whitespace.
    
    Args:
        paragraph: The paragraph element
        
    Returns:
        The normalized text content of the paragraph
    """
    try:
        # Use xpath to get all text nodes from the paragraph
        text_nodes = paragraph.xpath('.//text()')
        
        # Combine all text nodes
        full_text = ''.join(text_nodes)
        
        # Normalize whitespace but maintain structure
        # This doesn't collapse all whitespace, just redundant spaces
        return ' '.join(full_text.split())
    except Exception as e:
        print(f"Error in get_normalized_paragraph_text: {e}")
        import traceback
        traceback.print_exc()
        
        # Fallback approach
        try:
            # Simple concatenation of all text nodes
            return ''.join(paragraph.xpath('.//text()'))
        except:
            # Last resort: stringify and extract text
            from lxml import etree
            try:
                para_str = etree.tostring(paragraph, encoding='unicode')
                # Remove all XML tags to get plain text
                import re
                return re.sub(r'<[^>]+>', '', para_str)
            except:
                return ""
