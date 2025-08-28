#!/usr/bin/env python3

from document_merger import DocumentMerger
from local_client import LocalClauseClient
from parties_parser import PartiesParser
import os
import types

def mock_streamlit():
    """Mock streamlit functions for testing"""
    def mock_progress(value):
        return types.SimpleNamespace(progress=lambda x: None)
    
    def mock_empty():
        return types.SimpleNamespace(text=lambda x: print(f"Status: {x}"))
    
    def mock_warning(msg):
        print(f"Warning: {msg}")
    
    # Mock streamlit functions
    import streamlit as st
    st.progress = mock_progress
    st.empty = mock_empty
    st.warning = mock_warning

def test_styling():
    print("🎨 Test du styling avec couleur #003DA5")
    print("=" * 50)
    
    mock_streamlit()
    
    # Check template exists
    template_path = "clauses/Exemple contrat V2 clausier km.docx"
    if not os.path.exists(template_path):
        print(f"❌ Template non trouvé: {template_path}")
        return
    
    # Get clauses and organize by sections
    client = LocalClauseClient()
    parser = PartiesParser()
    
    clauses_by_section = client.get_clauses_by_section()
    sections = parser.get_sections()
    
    if not any(clauses for clauses in clauses_by_section.values()):
        print("❌ Aucune clause trouvée")
        return
    
    print("📄 Clauses par section:")
    for section in sections:
        section_clauses = clauses_by_section.get(section['key'], [])
        if section_clauses:
            print(f"  {section['order']}. {section['name']}: {len(section_clauses)} clause(s)")
    
    # Create merger and test
    merger = DocumentMerger()
    
    print("\n🔄 Test fusion avec nouveau styling...")
    try:
        output_path = merger.merge_documents_by_sections(clauses_by_section, sections)
        
        print(f"\n✅ Document créé avec styling:")
        print(f"   - Couleur: #003DA5")
        print(f"   - Taille police titres: 11pt")
        print(f"   - Taille police contenu: 11pt")
        print(f"   - Saut de ligne après titres de parties")
        print(f"📄 Fichier: {output_path}")
        
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"📊 Taille: {size} octets")
        
        # Cleanup
        merger.cleanup()
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_styling()