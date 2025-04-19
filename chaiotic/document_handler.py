"""Document handler module for reading and writing documents."""

# This file now serves as a facade for our modular document handling system
# and maintains backward compatibility with existing code

from .document_reader import read_document, read_docx, read_odt, read_text_file
from .document_extractor import extract_structured_content
from .document_writer import save_document
from .document_creator import create_sample_document

# Re-export all the functions to maintain backward compatibility
__all__ = [
    'read_document',
    'save_document',
    'create_sample_document',
    'extract_structured_content'
]