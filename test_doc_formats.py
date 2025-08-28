#!/usr/bin/env python3
"""
Test different approaches to read .doc files
"""

import os

def test_file_format(file_path):
    print(f"Testing file: {file_path}")
    
    # Check if it's a zip file (modern .docx)
    try:
        import zipfile
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            files = zip_ref.namelist()
            print(f"  ✅ ZIP format detected, contains {len(files)} files")
            if 'word/document.xml' in files:
                print("  ✅ Modern .docx format confirmed")
                return 'docx'
            else:
                print("  ⚠️ ZIP but not .docx format")
                return 'zip_unknown'
    except:
        print("  ❌ Not a ZIP file")
    
    # Check if it's an OLE file (legacy .doc)
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            if header[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                print("  ✅ OLE/Compound Document format detected (legacy .doc)")
                return 'doc_ole'
    except:
        pass
    
    # Check raw content
    try:
        with open(file_path, 'rb') as f:
            content = f.read(100)
            print(f"  Raw bytes: {content}")
    except:
        pass
    
    return 'unknown'

def test_extraction_methods(file_path):
    print(f"\nTesting extraction methods for: {file_path}")
    
    # Test docx2txt
    try:
        import docx2txt
        text = docx2txt.process(file_path)
        print(f"  docx2txt: {len(text)} chars - '{text[:50]}...'")
    except Exception as e:
        print(f"  docx2txt: FAILED - {e}")
    
    # Test python-docx
    try:
        from docx import Document
        doc = Document(file_path)
        text = '\n'.join([p.text for p in doc.paragraphs])
        print(f"  python-docx: {len(text)} chars - '{text[:50]}...'")
    except Exception as e:
        print(f"  python-docx: FAILED - {e}")

if __name__ == "__main__":
    test_file = "/home/arthurc/dev/projects/Clausier/clauses/04_Objet_du_Contrat/Objet.doc"
    
    format_type = test_file_format(test_file)
    test_extraction_methods(test_file)
    
    print(f"\nConclusion: File is {format_type}")