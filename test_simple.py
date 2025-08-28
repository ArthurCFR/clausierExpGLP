#!/usr/bin/env python3

import streamlit as st
from local_client import LocalClauseClient

def simple_test():
    st.title("🧪 Test Simple Clausier")
    
    if st.button("Charger les clauses"):
        client = LocalClauseClient()
        clauses_by_section = client.get_clauses_by_section()
        
        st.write(f"Clauses chargées dans {len(clauses_by_section)} sections")
        
        # Check if any sections have clauses
        sections_with_clauses = [k for k, v in clauses_by_section.items() if v]
        st.write(f"Sections avec clauses: {len(sections_with_clauses)}")
        
        # Test the condition
        has_clauses = clauses_by_section and any(clauses for clauses in clauses_by_section.values())
        st.write(f"Condition any() clauses: {has_clauses}")
        
        if has_clauses:
            st.success("✅ Des clauses sont disponibles pour affichage")
            
            for section_key, clauses in clauses_by_section.items():
                if clauses:
                    st.subheader(f"Section: {section_key}")
                    clause_names = [clause['name'] for clause in clauses]
                    selected = st.multiselect(
                        f"Clauses pour {section_key}:",
                        options=clause_names,
                        key=f"test_{section_key}"
                    )
                    if selected:
                        st.write(f"Sélectionnées: {selected}")
        else:
            st.error("❌ Aucune clause disponible")

if __name__ == "__main__":
    simple_test()