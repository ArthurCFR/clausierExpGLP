#!/usr/bin/env python3

import streamlit as st
from local_client import LocalClauseClient
from parties_parser import PartiesParser

def test_streamlit_visibility():
    st.title("🧪 Test de visibilité des fichiers .doc")
    
    if st.button("🔍 Tester la détection des fichiers"):
        client = LocalClauseClient()
        parser = PartiesParser()
        
        # Get clauses
        clause_files = client.get_clause_files()
        clauses_by_section = client.get_clauses_by_section()
        sections = parser.get_sections()
        
        st.success(f"✅ {len(clause_files)} fichiers détectés au total")
        
        # Show file breakdown
        doc_count = sum(1 for c in clause_files if c['file_name'].endswith('.doc'))
        docx_count = sum(1 for c in clause_files if c['file_name'].endswith('.docx'))
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Fichiers .docx", docx_count)
        with col2:
            st.metric("Fichiers .doc", doc_count)
        
        # Show sections with dropdowns
        st.subheader("📋 Test des listes déroulantes")
        
        for section in sections:
            section_clauses = clauses_by_section.get(section['key'], [])
            
            if section_clauses:
                st.markdown(f"### {section['order']}. {section['name']}")
                
                clause_options = []
                for clause in section_clauses:
                    file_type = ".doc" if clause['file_name'].endswith('.doc') else ".docx"
                    option_label = f"{clause['name']} ({file_type})"
                    clause_options.append(option_label)
                
                selected = st.multiselect(
                    f"Clauses pour: {section['name']}",
                    options=clause_options,
                    key=f"test_section_{section['key']}"
                )
                
                if selected:
                    st.write(f"**Sélectionnées:** {', '.join(selected)}")
        
        # Cleanup
        client.cleanup()

if __name__ == "__main__":
    test_streamlit_visibility()