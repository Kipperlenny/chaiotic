#!/usr/bin/env python3
"""
Grammar and Logic Checker for document files
Main entry point for the application
"""

import argparse
import os
import sys

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description='Chaiotic: Grammar and Logic Checker')
    parser.add_argument('--file', type=str, help='Path to the document file (ODT)')
    parser.add_argument('--structured', action='store_true', help='Use structured content extraction')
    parser.add_argument('--clean-checkpoints', action='store_true', help='Clean up all previous checkpoints')
    parser.add_argument('--keep-checkpoints', action='store_true', help='Keep checkpoints after successful completion')
    parser.add_argument('--max-checkpoints', type=int, default=5, help='Maximum number of checkpoints to keep')
    parser.add_argument('--nocache', action='store_true', help='Disable API request caching')
    return parser.parse_args()

def main():
    """Main entry point for the application."""
    args = parse_arguments()
    
    # Add parent directory to sys.path to allow imports
    sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
    
    # Get imports
    from chaiotic.document_handler import read_document, save_document, create_sample_document
    from chaiotic.grammar_checker import check_grammar, display_corrections, CheckpointHandler
    from utils.text_utils import preprocess_content
    from chaiotic.config import load_config
    
    # Load configuration
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
    
    try:
        # Get file path
        file_path = args.file
        if not file_path:
            file_path = input("Enter the path to the document file (ODT): ").strip()
        
        # Use default if no input
        if not file_path:
            file_path = 'test.odt'
            print(f"No file path provided. Using default: {file_path}")
            
            # Create sample file if needed
            if not os.path.exists(file_path):
                print(f"Creating sample file: {file_path}")
                create_sample_document(file_path)
        
        # Validate file extension
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext not in ['.odt']:
            print("Please provide a valid .odt file.")
            return 1
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
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
            content, structured_content = read_document(file_path)
            
            # Use structured content if available and requested
            if use_structured and structured_content and isinstance(structured_content, list):
                print(f"Using structured content with {len(structured_content)} elements")
            else:
                use_structured = False
                structured_content = None
                print("Using standard content processing")

            # Show preview of the content
            preview_length = min(200, len(content) if content else 0)
            print(f"\nPreview of '{file_path}':")
            print("=" * 40)
            print(content[:preview_length] + ("..." if len(content) > preview_length else ""))
            print("=" * 40)
            
            # Preprocess the content
            content = preprocess_content(content)
            
            print(f"Checking grammar and spelling of German text in {file_path}...")
            # Process with structure-aware grammar checker or standard checker
            corrections = check_grammar(content, structured_content, use_structured, checkpoint_handler=checkpoint_handler)
            
            # Display corrections if available
            if isinstance(corrections, dict) and 'corrections' in corrections:
                display_corrections(corrections['corrections'])
            elif isinstance(corrections, list):
                display_corrections(corrections)

            # Save document with corrections
            print("\nSaving corrected document...")
            json_path, text_path, doc_path = save_document(
                file_path,
                corrections,
                original_doc=None
            )
            
            if json_path:
                print(f"\nSaved corrections to: {json_path}")
            if text_path:
                print(f"Saved text version to: {text_path}")
            if doc_path:
                print(f"Saved corrected document to: {doc_path}")
            
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
        
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        return 1
    except Exception as e:
        import traceback
        print(f"Error: {e}")
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == '__main__':
    sys.exit(main())