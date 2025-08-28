#!/usr/bin/env python3

import streamlit as st
from local_client import LocalClauseClient
from parties_parser import PartiesParser

def debug_main():
    st.title("🐛 Debug Clausier")
    
    # Initialize
    parser = PartiesParser()
    client = LocalClauseClient()
    
    # Debug info
    sections = parser.get_sections()
    st.write(f"**Sections chargées:** {len(sections)}")
    
    clause_files = client.get_clause_files()
    st.write(f"**Fichiers de clauses:** {len(clause_files)}")
    
    clauses_by_section = client.get_clauses_by_section()
    st.write(f"**Sections avec clauses:** {len([s for s in clauses_by_section.values() if s])}")
    
    # Show clauses by section
    st.subheader("Debug: Clauses par section")
    
    for section in sections:
        section_clauses = clauses_by_section.get(section['key'], [])
        if section_clauses:
            st.write(f"**{section['order']}. {section['name']}** ({len(section_clauses)} clause(s))")
            
            clause_options = [clause['name'] for clause in section_clauses]
            selected = st.multiselect(
                f"Clauses pour: {section['name']}",
                options=clause_options,
                key=f"debug_section_{section['key']}"
            )
            
            if selected:
                st.write(f"Sélectionnées: {selected}")

if __name__ == "__main__":
    debug_main()