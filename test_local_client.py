#!/usr/bin/env python3

from local_client import LocalClauseClient
from parties_parser import PartiesParser
import os

def test_local_client():
    print("🧪 Test du LocalClauseClient")
    print("=" * 50)
    
    # Test existence du dossier clauses
    clauses_dir = "clauses"
    if not os.path.exists(clauses_dir):
        print(f"❌ Le dossier {clauses_dir} n'existe pas")
        return
    
    print(f"✅ Dossier {clauses_dir} trouvé")
    
    # List directories
    subdirs = [d for d in os.listdir(clauses_dir) if os.path.isdir(os.path.join(clauses_dir, d))]
    print(f"📁 {len(subdirs)} sous-dossiers trouvés:")
    for subdir in sorted(subdirs):
        files = os.listdir(os.path.join(clauses_dir, subdir))
        docx_files = [f for f in files if f.endswith('.docx')]
        print(f"   - {subdir}: {len(docx_files)} fichier(s) .docx")
    
    print("\n" + "=" * 50)
    
    # Test PartiesParser
    print("🧪 Test du PartiesParser")
    parser = PartiesParser()
    sections = parser.get_sections()
    print(f"📋 {len(sections)} sections chargées depuis parties.ini")
    
    print("\n" + "=" * 50)
    
    # Test LocalClauseClient
    print("🧪 Test du LocalClauseClient")
    client = LocalClauseClient()
    
    # Get clause files
    clause_files = client.get_clause_files()
    print(f"📄 {len(clause_files)} fichiers de clauses trouvés:")
    for clause in clause_files:
        print(f"   - {clause['name']} (section: {clause.get('section_name', 'N/A')})")
    
    print("\n" + "=" * 50)
    
    # Get clauses by section
    print("🧪 Test du regroupement par section")
    clauses_by_section = client.get_clauses_by_section()
    
    for section_key, clauses in clauses_by_section.items():
        if clauses:  # Only show sections with clauses
            section_info = parser.find_section_by_key(section_key)
            section_name = section_info['name'] if section_info else section_key
            print(f"📋 {section_name} ({len(clauses)} clause(s)):")
            for clause in clauses:
                print(f"   - {clause['name']}")
    
    print("\n" + "=" * 50)
    print("✅ Test terminé")

if __name__ == "__main__":
    test_local_client()