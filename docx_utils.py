from lxml import etree
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import io
from docx.opc.part import PartFactory
from docx.opc.rel import Relationships
from docx.shared import RGBColor

def add_comments_to_paragraph(paragraph, sorted_corrections):
    for correction in sorted_corrections:
        add_comment(paragraph, correction['run'], correction['text'], correction['author'])

def add_comment(paragraph, run, comment_text, author="Grammatik-Prüfung"):
    # The document object is needed to manipulate comments
    doc = paragraph.part.document
    # Use custom comment creation function that works with python-docx
    add_comment_directly(doc, paragraph, comment_text, author)

def create_comments_part(document):
    """Create a comments part for the document if it doesn't exist."""
    part = document.part
    
    # Check if comments part already exists
    for rel in part.rels.values():
        if rel.reltype == RT.COMMENTS:
            return rel.target_part
    
    # Create a new comments part
    comments_part_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>'
    comments_part = part.package.create_part('/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml', comments_part_xml)
    
    # Add relationship to the comments part
    part.rels.add_relationship(RT.COMMENTS, comments_part)
    
    return comments_part

def get_next_comment_id(document):
    """Get the next available comment ID."""
    comments_part = create_comments_part(document)
    comments_element = etree.fromstring(comments_part.blob)
    existing_comments = comments_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment')
    if not existing_comments:
        return 1
    return max([int(comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')) for comment in existing_comments]) + 1

def create_comment_reference(comment_id):
    """Create XML for a comment reference."""
    return f'<w:commentReference w:id="{comment_id}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'

def create_comment(comment_id, text, author, initials=None):
    """Create XML for a comment."""
    from datetime import datetime
    initials = initials or author[0]
    date = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    return f'<w:comment w:id="{comment_id}" w:author="{author}" w:initials="{initials}" w:date="{date}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>'

def copy_run_formatting(source_run, target_run):
    """Copy formatting from one run to another."""
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    if hasattr(source_run, 'font') and hasattr(target_run, 'font'):
        if hasattr(source_run.font, 'size'):
            target_run.font.size = source_run.font.size
        if hasattr(source_run.font, 'name'):
            target_run.font.name = source_run.font.name
        if hasattr(source_run.font, 'color') and hasattr(target_run.font.color, 'rgb'):
            target_run.font.color.rgb = source_run.font.color.rgb

def add_comment_to_paragraph(document, paragraph, comment_text, author="Chaiotic", initials="CH"):
    """Add a comment to a paragraph. This is a simplified version that should work with python-docx."""
    add_comment_directly(document, paragraph, comment_text, author, initials)

def add_comment_directly(document, paragraph, comment_text, author="Chaiotic", initials="CH"):
    """Add a comment directly to a paragraph by manipulating the XML."""
    from datetime import datetime
    import uuid
    from docx.oxml import OxmlElement
    
    # Get the document part
    part = document.part if hasattr(document, 'part') else document
    
    # Create a unique comment ID
    comment_id = str(uuid.uuid4())[:8]
    
    # Create a comment reference in the paragraph
    run = paragraph.add_run()
    comment_reference = OxmlElement('w:commentReference')
    comment_reference.set(qn('w:id'), comment_id)
    run._element.append(comment_reference)
    
    # Add the comment directly to the document without using comments_part
    # This is a workaround for the lack of comments_part support in python-docx
    
    # Create the comment as a paragraph in the document
    comment_para = paragraph.add_run(f" [{author}: {comment_text}]")
    comment_para.italic = True
    comment_para.font.color.rgb = RGBColor(128, 128, 128)  # Gray color
    
    return comment_id

def _get_new_comment_id(paragraph):
    # This function is deprecated - use get_next_comment_id instead
    return get_next_comment_id(paragraph.part.document)

def _get_or_create_comments_part(part):
    # This function is deprecated - use create_comments_part instead
    return create_comments_part(part.document)

def add_tracked_change_with_comment(paragraph, original_text, corrected_text, explanation):
    """Add a tracked change and a comment to a paragraph in a DOCX document.

    Args:
        paragraph: The paragraph to modify.
        original_text: The original text to replace.
        corrected_text: The corrected text to insert.
        explanation: The explanation for the change, added as a comment.
    """
    # Find the run containing the original text
    for i, run in enumerate(paragraph.runs):
        if original_text in run.text:
            # Get the paragraph text before modifying
            original_para_text = paragraph.text
            
            # Split the run into parts
            before, match, after = run.text.partition(original_text)
            
            # Clear existing runs from paragraph to rebuild it
            p_element = paragraph._element
            for _ in range(len(paragraph.runs)):
                if len(p_element.r_lst) > 0:
                    p_element.remove(p_element.r_lst[0])
            
            # Add the text before the correction
            if before:
                before_run = paragraph.add_run(before)
            
            # Add the original text with strikethrough
            del_run = paragraph.add_run(original_text)
            del_run.font.strike = True
            
            # Add the corrected text in bold with a different color
            from docx.shared import RGBColor
            ins_run = paragraph.add_run(corrected_text)
            ins_run.bold = True
            ins_run.font.color.rgb = RGBColor(255, 0, 0)  # Red color for insertions
            
            # Add the text after the correction
            if after:
                after_run = paragraph.add_run(after)
            
            # Add the explanation as italic gray text in brackets
            comment_run = paragraph.add_run(f" [{explanation}]")
            comment_run.italic = True
            comment_run.font.color.rgb = RGBColor(128, 128, 128)  # Gray color for comments
            
            print(f"Applied change: '{original_text}' → '{corrected_text}' with explanation: {explanation}")
            print(f"Paragraph before: '{original_para_text}'")
            print(f"Paragraph after: '{paragraph.text}'")
            
            return True
    
    # If we reach here, the original text wasn't found
    print(f"WARNING: Could not find text '{original_text}' in paragraph: '{paragraph.text}'")
    return False

def add_tracked_change(document, paragraph, original_text, corrected_text):
    """Add a tracked change to a paragraph by manipulating the XML directly."""
    from datetime import datetime
    
    # Get the paragraph's element
    p_elem = paragraph._element
    
    # Find the run containing the original text
    found = False
    for run in paragraph.runs:
        if original_text in run.text:
            # Split the run into parts
            before, match, after = run.text.partition(original_text)
            
            # Update the run with just the "before" text
            run.text = before
            
            # Add strikethrough text for original
            del_run = paragraph.add_run(original_text)
            del_run.font.strike = True
            
            # Add the corrected text in bold
            ins_run = paragraph.add_run(corrected_text)
            ins_run.bold = True
            
            # Add the "after" text if needed
            if after:
                paragraph.add_run(after)
                
            found = True
            break
    
    return found