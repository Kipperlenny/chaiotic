"""Module for handling ODT document operations."""

import os
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Union

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
    """Apply corrections to an ODT document using lxml library.
    
    Args:
        file_path: Path to the original ODT document
        corrections: List of correction dictionaries
        output_path: Path to save the corrected document
        
    Returns:
        Path to the saved document
    """
    import lxml.etree as LET
    
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
        
        # Ensure required namespaces are present and remove problematic ones
        if 'text' not in nsmap:
            nsmap['text'] = ODT_NAMESPACES['text']
        if 'dc' not in nsmap:
            nsmap['dc'] = ODT_NAMESPACES['dc']
            
        # Remove any 'officeooo' namespace which causes validation errors
        if any(ns.startswith('officeooo') for ns in nsmap):
            # Create a new nsmap without the problematic namespace
            cleaned_nsmap = {k: v for k, v in nsmap.items() 
                            if not k.startswith('officeooo')}
            nsmap = cleaned_nsmap
        
        # Find text:tracked-changes element or create it
        office_text = None
        for child in root.iter('{' + nsmap['office'] + '}text'):
            office_text = child
            break
        
        if office_text is None:
            print("Warning: office:text element not found")
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
        
        # Process corrections
        change_id = 1
        for correction in corrections:
            original = correction.get('original', '')
            corrected = correction.get('corrected', '')
            
            if not original or original == corrected:
                continue
                
            # Create two separate region IDs for deletion and insertion
            del_change_id_str = f"ctd{change_id}"
            ins_change_id_str = f"cti{change_id}"
            
            # 1. Create changed-region for DELETION with required attributes
            del_region = LET.Element('{' + nsmap['text'] + '}changed-region', 
                                   {
                                       '{' + nsmap['text'] + '}id': del_change_id_str,
                                       '{http://www.w3.org/XML/1998/namespace}id': del_change_id_str
                                   }, 
                                   nsmap=nsmap)
            
            # Add deletion element to the deletion region
            deletion = LET.SubElement(del_region, '{' + nsmap['text'] + '}deletion')
            
            # Add change-info element to deletion
            change_info = LET.SubElement(deletion, '{' + nsmap['office'] + '}change-info')
            
            # Add creator and date to change-info
            creator = LET.SubElement(change_info, '{' + nsmap['dc'] + '}creator')
            creator.text = "Chaiotic Grammar Checker"
            date = LET.SubElement(change_info, '{' + nsmap['dc'] + '}date')
            date.text = datetime.now().isoformat()
            
            # Add deleted text content after change-info
            deleted_text = LET.SubElement(deletion, '{' + nsmap['text'] + '}p')
            deleted_text.text = original
            
            # Add the deletion region to tracked-changes
            tracked_changes.append(del_region)
            
            # 2. Create a SEPARATE changed-region for INSERTION with required attributes
            ins_region = LET.Element('{' + nsmap['text'] + '}changed-region', 
                                    {
                                        '{' + nsmap['text'] + '}id': ins_change_id_str,
                                        '{http://www.w3.org/XML/1998/namespace}id': ins_change_id_str
                                    }, 
                                    nsmap=nsmap)
            
            # Add insertion element to the insertion region
            insertion = LET.SubElement(ins_region, '{' + nsmap['text'] + '}insertion')
            
            # Add change-info element to insertion
            change_info = LET.SubElement(insertion, '{' + nsmap['office'] + '}change-info')
            
            # Add creator and date to change-info
            creator = LET.SubElement(change_info, '{' + nsmap['dc'] + '}creator')
            creator.text = "Chaiotic Grammar Checker"
            date = LET.SubElement(change_info, '{' + nsmap['dc'] + '}date')
            date.text = datetime.now().isoformat()
            
            # Add inserted text content after change-info
            inserted_text = LET.SubElement(insertion, '{' + nsmap['text'] + '}p')
            inserted_text.text = corrected
            
            # Add the insertion region to tracked-changes
            tracked_changes.append(ins_region)
            
            # Find and update text in paragraphs
            for paragraph in office_text.findall('.//{' + nsmap['text'] + '}p'):
                try:
                    # Skip paragraphs in tracked changes
                    if paragraph.getparent().tag == '{' + nsmap['text'] + '}deletion' or \
                       paragraph.getparent().tag == '{' + nsmap['text'] + '}insertion':
                        continue
                    
                    text_content = ''.join(paragraph.xpath('.//text()'))
                    
                    if original in text_content:
                        # Replace text with tracked change markers
                        # Find text node containing the original text
                        for node in paragraph.iter():
                            if node.text and original in node.text:
                                # Split the text and add markers
                                parts = node.text.split(original, 1)
                                
                                # Clear the node
                                saved_tail = node.tail
                                node.clear()
                                node.tail = saved_tail
                                
                                # Add first part
                                if parts[0]:
                                    node.text = parts[0]
                                
                                # Add deletion marker
                                change_start = LET.SubElement(
                                    node, '{' + nsmap['text'] + '}change-start',
                                    {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                                )
                                
                                # Add change end for deletion
                                change_end = LET.SubElement(
                                    node, '{' + nsmap['text'] + '}change-end',
                                    {'{' + nsmap['text'] + '}change-id': del_change_id_str}
                                )
                                
                                # Add insertion marker
                                change_start = LET.SubElement(
                                    node, '{' + nsmap['text'] + '}change-start',
                                    {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                                )
                                
                                # Add corrected text
                                span = LET.SubElement(node, '{' + nsmap['text'] + '}span')
                                span.text = corrected
                                
                                # Add change end for insertion
                                change_end = LET.SubElement(
                                    node, '{' + nsmap['text'] + '}change-end',
                                    {'{' + nsmap['text'] + '}change-id': ins_change_id_str}
                                )
                                
                                # Add rest of text
                                if len(parts) > 1 and parts[1]:
                                    after_span = LET.SubElement(node, '{' + nsmap['text'] + '}span')
                                    after_span.text = parts[1]
                                
                                break
                except Exception as e:
                    print(f"Error processing paragraph: {e}")
                    continue
        
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
            f.write('    </style:style>\n')
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