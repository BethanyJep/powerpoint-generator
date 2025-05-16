#!/usr/bin/env python3
"""
Demo script for Word to PowerPoint conversion using the existing Word document in the workspace.
"""

import os
import argparse
from word_to_ppt import WordToPowerPointConverter

def main():
    parser = argparse.ArgumentParser(description='Demo Word to PowerPoint conversion (vanilla)')
    parser.add_argument('input_file', nargs='?', default='develop-ai-agent-with-semantic-kernel.docx', help='Path to the Word document')
    parser.add_argument('--output', '-o', help='Path for the output PowerPoint file')
    
    args = parser.parse_args()
    
    # Find the Word document in the workspace
    word_doc = args.input_file
    
    if not os.path.exists(word_doc):
        print(f"❌ Could not find the Word document {word_doc}")
        print("Please make sure the document exists or modify this script to point to your document.")
        return
    
    # Create output path
    output_path = args.output or os.path.splitext(word_doc)[0] + '_presentation.pptx'
    
    print(f"🔍 Found Word document: {word_doc}")
    print(f"📝 Will generate PowerPoint: {output_path}")
    
    # Confirm with user
    proceed = input("Continue? (y/n): ")
    if proceed.lower() != 'y':
        print("Operation cancelled.")
        return
    
    try:
        # Create vanilla converter
        converter = WordToPowerPointConverter()
        
        print("🔄 Converting document to PowerPoint...")
        
        # Convert the document
        output_file = converter.convert_document(word_doc, output_path)
        
        print(f"\n✅ Successfully created PowerPoint presentation: {output_file}")
        
    except Exception as e:
        print(f"❌ Error during conversion: {str(e)}")

if __name__ == '__main__':
    main()
