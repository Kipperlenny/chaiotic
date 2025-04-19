#!/usr/bin/env python3
"""
Grammar and Logic Checker for document files
Main entry point for the application
"""

import argparse
import os
import sys

def main():
    """Main entry point for the application."""
    # Add parent directory to sys.path to allow imports
    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Chaiotic: Grammar and Logic Checker')
    parser.add_argument('--file', type=str, help='Path to the document file (DOCX or ODT)')
    parser.add_argument('--structured', action='store_true', help='Use structured content extraction')
    parser.add_argument('--clean-checkpoints', action='store_true', help='Clean up all previous checkpoints')
    parser.add_argument('--keep-checkpoints', action='store_true', help='Keep checkpoints after successful completion')
    parser.add_argument('--max-checkpoints', type=int, default=5, help='Maximum number of checkpoints to keep')
    parser.add_argument('--nocache', action='store_true', help='Disable API request caching')
    args = parser.parse_args()
    
    try:
        # Import necessary modules
        from chaiotic.document_handler import read_document, save_document, create_sample_document, extract_structured_content
        from chaiotic.document_writer import save_correction_outputs
        from chaiotic.grammar_checker import check_grammar, display_corrections, CheckpointHandler
        from chaiotic.utils import preprocess_content
        from chaiotic.config import load_config
        
        # Load configuration and set cache settings
        config = load_config()
        if args.nocache:
            config.set_cache_enabled(False)
        
        # Initialize checkpoint handler
        checkpoint_handler = CheckpointHandler(
            max_checkpoints=args.max_checkpoints, 
            keep_last=args.keep_checkpoints
        )
        
        # Clean checkpoints if requested
        if args.clean_checkpoints:
            checkpoint_handler.purge_all_checkpoints()
            print("All checkpoints purged.")
        
        # Get file path from command line args or prompt user
        file_path = args.file
        if not file_path:
            file_path = input("Enter the path to the document file (DOCX or ODT): ")

        # If nothing is entered, use test.odt
        if not file_path:
            file_path = 'test.odt'
            print(f"No file path provided. Using default: {file_path}")
            
            # If default file doesn't exist, create a sample file
            if not os.path.exists(file_path):
                print(f"Default file {file_path} not found. Creating sample file...")
                create_sample_document(file_path)

        # Check file extension
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext == '.docx':
            is_docx = True
        elif file_ext == '.odt':
            is_docx = False
        else:
            print("Please provide a valid .docx or .odt file.")
            return 1
            
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            print("Please provide a valid file path.")
            return 1
        
        # Ask user what they want to do
        print("\nWhat would you like to do with this document?")
        print("1. Check grammar and spelling (GPT-4o-mini)")
        print("2. Check logic and get creative ideas (GPT-4.1) - Coming soon")
        choice = input("Enter your choice (1 or 2): ")

        if not choice:
            choice = "1"
            print("No choice provided. Defaulting to grammar and spelling check.")
        
        if choice == "1":
            # Use structured processing if specified by command line flag
            use_structured = args.structured
            
            # Read the document content
            print(f"Reading document: {file_path}")
            content, doc_obj, is_docx = read_document(file_path)
            
            # Extract structured content if needed
            if use_structured:
                print("Using structured content extraction...")
                # Fix: extract_structured_content expects just the file path
                structured_content = extract_structured_content(file_path)
                
                # Check if we got valid structured content
                if not isinstance(structured_content, (list, dict)) or len(structured_content) == 0:
                    print("No structured content found. Falling back to full text processing.")
                    use_structured = False
                    structured_content = None
                else:
                    print(f"Extracted {len(structured_content)} structured elements.")
            else:
                structured_content = None
            
            # Show preview of the content
            preview_length = min(200, len(content))
            print(f"\nPreview of '{file_path}':")
            print("=" * 40)
            print(content[:preview_length] + ("..." if len(content) > preview_length else ""))
            print("=" * 40)
            
            # Preprocess the content
            content = preprocess_content(content)
            
            print(f"Checking grammar and spelling of German text in {file_path}...")
            # Process with structure-aware grammar checker or standard checker
            corrections = check_grammar(content, structured_content, use_structured, checkpoint_handler=checkpoint_handler)
            
            # Display corrections
            if corrections:
                print("\nCorrections:")
                display_corrections(corrections)
                
                # Ask if user wants to save corrections
                save_prompt = input("\nDo you want to save corrections to a new document? (y/n): ").lower()
                if save_prompt == 'y' or save_prompt == 'yes':
                    # Save corrections
                    json_path, text_path, doc_path = save_correction_outputs(file_path, corrections, doc_obj, is_docx)
                    print(f"\nOutput files: \n- {json_path}\n- {text_path}\n- {doc_path}")
                else:
                    print("Corrections not saved.")
            else:
                print("No corrections found or failed to process corrections.")
            
            # Clean up checkpoints after successful completion (if not keeping them)
            if not args.keep_checkpoints:
                checkpoint_handler.clean_up_on_success()
        
        elif choice == "2":
            print(f"Checking logic and generating creative ideas using GPT-4.1...")
            print("This feature is coming soon. Currently not implemented.")
            # Future implementation
            # from chaiotic.logic_checker import check_logic
            # check_logic(content)
        
        else:
            print("Invalid choice. Exiting.")
        
    except Exception as e:
        import traceback
        print(f"Error: {e}")
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == '__main__':
    sys.exit(main())