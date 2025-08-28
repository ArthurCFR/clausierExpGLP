#!/usr/bin/env python3

from document_merger import DocumentMerger
from local_client import LocalClauseClient
import os

def test_template_merger():
    print("🧪 Test du DocumentMerger avec template")
    print("=" * 50)
    
    # Check template exists
    template_path = "clauses/Exemple contrat V2 clausier km.docx"
    if not os.path.exists(template_path):
        print(f"❌ Template non trouvé: {template_path}")
        return
    
    print(f"✅ Template trouvé: {template_path}")
    
    # Get some clauses from local client
    client = LocalClauseClient()
    clause_files = client.get_clause_files()
    
    if not clause_files:
        print("❌ Aucune clause trouvée dans les dossiers locaux")
        return
    
    print(f"📄 {len(clause_files)} clauses trouvées")
    
    # Take first few clauses for test
    test_clauses = clause_files[:3]  # Take first 3 clauses
    
    print("📋 Clauses à fusionner:")
    for clause in test_clauses:
        print(f"  - {clause['name']} ({clause['section_name']})")
    
    # Prepare paths and names
    file_paths = [clause['file_path'] for clause in test_clauses]
    clause_names = [clause['name'] for clause in test_clauses]
    
    # Create merger with template
    merger = DocumentMerger()
    
    print("\n🔄 Fusion en cours...")
    try:
        # This would normally use streamlit progress bars, but we'll mock them
        import types
        
        # Mock streamlit functions for testing
        def mock_progress(value):
            return types.SimpleNamespace(progress=lambda x: None)
        
        def mock_empty():
            return types.SimpleNamespace(text=lambda x: print(f"Status: {x}"))
        
        # Monkey patch for testing
        import streamlit as st
        original_progress = getattr(st, 'progress', None)
        original_empty = getattr(st, 'empty', None)
        
        st.progress = mock_progress
        st.empty = mock_empty
        
        # Merge documents
        output_path = merger.merge_documents(file_paths, clause_names)
        
        # Restore original functions
        if original_progress:
            st.progress = original_progress
        if original_empty:
            st.empty = original_empty
        
        print(f"\n✅ Fusion terminée!")
        print(f"📄 Document créé: {output_path}")
        
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"📊 Taille: {size} octets")
        
        # Cleanup
        merger.cleanup()
        
    except Exception as e:
        print(f"❌ Erreur lors de la fusion: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_template_merger()