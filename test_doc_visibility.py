#!/usr/bin/env python3

from local_client import LocalClauseClient
import streamlit as st
import types

def mock_streamlit():
    """Mock streamlit functions"""
    def mock_warning(msg):
        print(f"Streamlit Warning: {msg}")
    
    st.warning = mock_warning

def test_doc_visibility():
    print("👁️ Test de visibilité des fichiers .doc")
    print("=" * 50)
    
    mock_streamlit()
    
    client = LocalClauseClient()
    
    # Get all clause files
    clause_files = client.get_clause_files()
    
    print(f"📄 {len(clause_files)} fichier(s) détecté(s) au total:")
    
    doc_files = []
    docx_files = []
    
    for clause in clause_files:
        file_type = "(.doc)" if clause['file_name'].endswith('.doc') else "(.docx)"
        legacy_marker = " [LEGACY]" if clause.get('is_legacy_doc', False) else ""
        
        print(f"   - {clause['name']} {file_type}{legacy_marker}")
        print(f"     Section: {clause['section_name']}")
        print(f"     Fichier: {clause['file_name']}")
        
        if clause['file_name'].endswith('.doc'):
            doc_files.append(clause)
        else:
            docx_files.append(clause)
    
    print(f"\n📊 Répartition:")
    print(f"   - Fichiers .docx: {len(docx_files)}")
    print(f"   - Fichiers .doc: {len(doc_files)}")
    
    # Test clauses by section
    print(f"\n📋 Clauses par section:")
    clauses_by_section = client.get_clauses_by_section()
    
    for section_key, clauses in clauses_by_section.items():
        if clauses:
            print(f"   {section_key}: {len(clauses)} clause(s)")
            for clause in clauses:
                file_type = "(.doc)" if clause['file_name'].endswith('.doc') else "(.docx)"
                print(f"      - {clause['name']} {file_type}")
    
    if doc_files:
        print(f"\n✅ {len(doc_files)} fichier(s) .doc trouvé(s) et reconnu(s)")
    else:
        print(f"\n❌ Aucun fichier .doc visible dans les résultats")
    
    # Cleanup
    client.cleanup()
    
    print("✅ Test terminé")

if __name__ == "__main__":
    test_doc_visibility()