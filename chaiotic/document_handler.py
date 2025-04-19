"""Document handler module for reading and writing documents."""

# This file now serves as a facade for our modular document handling system
# and maintains backward compatibility with existing code

from .document_reader import read_document, read_docx, read_odt, read_text_file
from .document_extractor import extract_structured_content
from .document_writer import save_document
from .document_creator import create_sample_document

import os

def create_sample_document(file_path):
    """Create a sample document file for testing.
    
    Args:
        file_path: Path where the sample document will be saved
    """
    # Check the file extension
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.docx':
        try:
            from docx import Document
            from docx.shared import Pt
            
            doc = Document()
            doc.add_heading('Sample Document', 0)
            
            p = doc.add_paragraph('This is a sample document created for testing. ')
            p.add_run('It contains some text with potential grammar issues that can be corrected.')
            
            doc.add_heading('Section with Grammar Issues', level=1)
            
            doc.add_paragraph('This sentence have a verb agreement error.')
            doc.add_paragraph('There are mistakes in this sentance.')
            doc.add_paragraph('This paragraph contains multiple erors that should be corected by the grammar checker.')
            
            doc.add_heading('Section with Formatting', level=1)
            p = doc.add_paragraph('This section has ')
            p.add_run('bold').bold = True
            p.add_run(' and ')
            p.add_run('italic').italic = True
            p.add_run(' text formatting.')
            
            # Save the document
            doc.save(file_path)
            print(f"Created sample DOCX file at {file_path}")
            
        except Exception as e:
            print(f"Error creating sample DOCX: {e}")
            
    elif file_ext == '.odt':
        try:
            # For ODT, we'll create an ODT file using odfpy library
            try:
                from odf.opendocument import OpenDocumentText
                from odf.style import Style, TextProperties, ParagraphProperties
                from odf.text import H, P, Span
                
                textdoc = OpenDocumentText()
                
                # Add styles
                heading_style = Style(name="Heading", family="paragraph")
                heading_style.addElement(TextProperties(fontsize="16pt", fontweight="bold"))
                textdoc.styles.addElement(heading_style)
                
                # Add content
                h = H(outlinelevel=1, text="Sample ODT Document")
                textdoc.text.addElement(h)
                
                p = P(text="This is a sample document created for testing. It contains some text with potential grammar issues that can be corrected.")
                textdoc.text.addElement(p)
                
                h = H(outlinelevel=2, text="Section with Grammar Issues")
                textdoc.text.addElement(h)
                
                p = P(text="This sentence have a verb agreement error.")
                textdoc.text.addElement(p)
                
                p = P(text="There are mistakes in this sentance.")
                textdoc.text.addElement(p)
                
                p = P(text="This paragraph contains multiple erors that should be corected by the grammar checker.")
                textdoc.text.addElement(p)
                
                textdoc.save(file_path)
                print(f"Created sample ODT file at {file_path}")
                return
            except ImportError:
                print("odfpy library not found, using alternative method")
            
            # Alternative method using a template and text replacement
            import zipfile
            import tempfile
            import shutil
            import os
            
            # Create a simple ODT file structure
            with tempfile.TemporaryDirectory() as temp_dir:
                # Create META-INF directory and mimetype
                os.makedirs(os.path.join(temp_dir, 'META-INF'))
                
                # Create content.xml
                with open(os.path.join(temp_dir, 'content.xml'), 'w', encoding='utf-8') as f:
                    f.write('''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" 
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" 
                         xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" 
                         xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" 
                         xmlns:xlink="http://www.w3.org/1999/xlink" 
                         xmlns:dc="http://purl.org/dc/elements/1.1/" 
                         xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" 
                         xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" 
                         xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" 
                         xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" 
                         xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" 
                         xmlns:math="http://www.w3.org/1998/Math/MathML" 
                         xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" 
                         xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" 
                         xmlns:ooo="http://openoffice.org/2004/office" 
                         xmlns:ooow="http://openoffice.org/2004/writer" 
                         xmlns:oooc="http://openoffice.org/2004/calc" 
                         xmlns:dom="http://www.w3.org/2001/xml-events" 
                         xmlns:xforms="http://www.w3.org/2002/xforms" 
                         xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
                         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                         office:version="1.2">
  <office:body>
    <office:text>
      <text:h text:style-name="Heading_20_1" text:outline-level="1">Sample ODT Document</text:h>
      <text:p text:style-name="Text_20_body">This is a sample document created for testing. It contains some text with potential grammar issues that can be corrected.</text:p>
      <text:h text:style-name="Heading_20_2" text:outline-level="2">Section with Grammar Issues</text:h>
      <text:p text:style-name="Text_20_body">This sentence have a verb agreement error.</text:p>
      <text:p text:style-name="Text_20_body">There are mistakes in this sentance.</text:p>
      <text:p text:style-name="Text_20_body">This paragraph contains multiple erors that should be corected by the grammar checker.</text:p>
    </office:text>
  </office:body>
</office:document-content>''')
                
                # Create mimetype file
                with open(os.path.join(temp_dir, 'mimetype'), 'w', encoding='utf-8') as f:
                    f.write('application/vnd.oasis.opendocument.text')
                
                # Create META-INF/manifest.xml
                with open(os.path.join(temp_dir, 'META-INF', 'manifest.xml'), 'w', encoding='utf-8') as f:
                    f.write('''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
 <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="META-INF/manifest.xml"/>
</manifest:manifest>''')
                
                # Create styles.xml
                with open(os.path.join(temp_dir, 'styles.xml'), 'w', encoding='utf-8') as f:
                    f.write('''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                        xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" 
                        xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" 
                        xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" 
                        office:version="1.2">
  <office:styles>
    <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
    <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:class="text">
      <style:text-properties fo:font-size="14pt" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text">
      <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.212cm"/>
    </style:style>
    <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:class="text">
      <style:text-properties fo:font-size="18pt" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:class="text">
      <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
    </style:style>
  </office:styles>
</office:document-styles>''')
                
                # Create meta.xml
                with open(os.path.join(temp_dir, 'meta.xml'), 'w', encoding='utf-8') as f:
                    f.write('''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                     xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" 
                     xmlns:dc="http://purl.org/dc/elements/1.1/" 
                     office:version="1.2">
  <office:meta>
    <dc:title>Sample ODT Document</dc:title>
    <dc:creator>Chaiotic Grammar Checker</dc:creator>
    <dc:date>2023-01-01T00:00:00</dc:date>
  </office:meta>
</office:document-meta>''')
                
                # Create settings.xml
                with open(os.path.join(temp_dir, 'settings.xml'), 'w', encoding='utf-8') as f:
                    f.write('''<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" 
                         xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" 
                         office:version="1.2">
  <office:settings>
    <config:config-item-set config:name="ooo:view-settings"/>
  </office:settings>
</office:document-settings>''')
                
                # Create the ODT file (which is a ZIP file)
                with zipfile.ZipFile(file_path, 'w') as zip_file:
                    # Add mimetype first without compression
                    zip_file.write(os.path.join(temp_dir, 'mimetype'), 'mimetype', compress_type=zipfile.ZIP_STORED)
                    
                    # Add other files with compression
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            if file != 'mimetype':  # Skip mimetype as we already added it
                                file_path_in_zip = os.path.relpath(os.path.join(root, file), temp_dir)
                                zip_file.write(os.path.join(root, file), file_path_in_zip, compress_type=zipfile.ZIP_DEFLATED)
                
                print(f"Created sample ODT file at {file_path}")
                
        except Exception as e:
            print(f"Error creating sample ODT: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"Cannot create sample document with extension {file_ext}. Please use .docx or .odt.")

# Re-export all the functions to maintain backward compatibility
__all__ = [
    'read_document',
    'save_document',
    'create_sample_document',
    'extract_structured_content'
]