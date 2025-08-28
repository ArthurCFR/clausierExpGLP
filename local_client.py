import os
import tempfile
import shutil
from typing import List, Dict, Optional
import streamlit as st
from parties_parser import PartiesParser
from doc_converter import DocConverter

class LocalClauseClient:
    """Client for reading clauses from local directory structure"""
    
    def __init__(self, clauses_dir: str = "clauses"):
        self.clauses_dir = clauses_dir
        self.parties_parser = PartiesParser()
        self._temp_dir = tempfile.mkdtemp()
        self.doc_converter = DocConverter()
    
    def get_clause_files(self) -> List[Dict[str, str]]:
        """Get list of clause files from local directories"""
        if not os.path.exists(self.clauses_dir):
            st.error(f"Le dossier {self.clauses_dir} n'existe pas")
            return []
        
        clause_files = []
        
        # Iterate through section directories
        for section_dir in sorted(os.listdir(self.clauses_dir)):
            section_path = os.path.join(self.clauses_dir, section_dir)
            
            if os.path.isdir(section_path):
                # Extract section info from directory name
                section_info = self._parse_directory_name(section_dir)
                
                # Find all .doc and .docx files in this section
                for filename in os.listdir(section_path):
                    if filename.endswith(('.doc', '.docx')):
                        file_path = os.path.join(section_path, filename)
                        
                        # Validate that the file is actually readable or convertible
                        is_valid, reason = self._is_valid_word_file(file_path)
                        if is_valid:
                            clause_name = filename.replace('.docx', '').replace('.doc', '')
                            
                            clause_files.append({
                                'name': clause_name,
                                'file_name': filename,
                                'file_path': file_path,
                                'section_tag': section_info['key'],
                                'section_order': section_info['order'],
                                'section_name': section_info['name'],
                                'is_legacy_doc': filename.endswith('.doc') and self.doc_converter.is_legacy_doc_file(file_path)
                            })
                        else:
                            # Use streamlit warning instead of print
                            import streamlit as st
                            st.warning(f"⚠️ Fichier ignoré: {filename} - {reason}")
        
        return sorted(clause_files, key=lambda x: (x['section_order'], x['name']))
    
    def _parse_directory_name(self, dir_name: str) -> Dict[str, any]:
        """Parse directory name to extract section info"""
        # Format: "01_Designation_des_Parties"
        parts = dir_name.split('_', 1)
        
        if len(parts) >= 2:
            try:
                order = int(parts[0])
                name_part = parts[1].replace('_', ' ')
                
                # Find corresponding section in parties.ini
                for section in self.parties_parser.get_sections():
                    if section['order'] == order:
                        return {
                            'order': order,
                            'key': section['key'],
                            'name': section['name']
                        }
                
                # Fallback if not found in parties.ini
                return {
                    'order': order,
                    'key': name_part.lower().replace(' ', '_'),
                    'name': name_part
                }
            except ValueError:
                pass
        
        # Fallback for malformed directory names
        return {
            'order': 999,
            'key': dir_name.lower(),
            'name': dir_name.replace('_', ' ')
        }
    
    def get_clauses_by_section(self) -> Dict[str, List[Dict[str, str]]]:
        """Get clauses grouped by section"""
        all_clauses = self.get_clause_files()
        clauses_by_section = {}
        
        # Initialize with all sections
        for section in self.parties_parser.get_sections():
            clauses_by_section[section['key']] = []
        
        # Group clauses by section
        for clause in all_clauses:
            section_key = clause.get('section_tag', 'uncategorized')
            if section_key not in clauses_by_section:
                clauses_by_section[section_key] = []
            clauses_by_section[section_key].append(clause)
        
        return clauses_by_section
    
    def _is_valid_word_file(self, file_path: str) -> tuple[bool, str]:
        """Check if a file is a valid Word document that can be read or converted"""
        try:
            from docx import Document
            # Try to open the document directly
            doc = Document(file_path)
            # If we can access paragraphs, it's a valid modern Word document
            _ = len(doc.paragraphs)
            return True, "Document Word moderne valide"
        except Exception as e:
            # If direct loading fails, check if it's a legacy .doc file we can convert
            if file_path.endswith('.doc'):
                # Check if it's a true legacy .doc file that we can convert
                try:
                    if self.doc_converter.is_legacy_doc_file(file_path):
                        return True, "Fichier legacy Word 97-2003 (sera converti automatiquement)"
                    else:
                        return False, "Fichier .doc invalide ou corrompu"
                except Exception as conv_error:
                    return False, f"Erreur lors de la validation .doc: {str(conv_error)}"
            
            # For .docx files or other errors, it's not valid
            if "not a Word file" in str(e):
                return False, "Format de fichier non supporté"
            elif "content type" in str(e):
                return False, "Extension incorrecte ou fichier corrompu"
            else:
                return False, f"Erreur de lecture: {str(e)[:50]}"
    
    def download_selected_clauses(self, selected_clauses: List[Dict[str, str]]) -> List[str]:
        """Copy selected clause files to temp directory and return paths"""
        downloaded_files = []
        
        if not selected_clauses:
            return downloaded_files
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, clause in enumerate(selected_clauses):
            status_text.text(f"Copie: {clause['name']}")
            
            try:
                source_path = clause['file_path']
                temp_filename = f"{i+1:03d}_{clause['file_name']}"
                temp_path = os.path.join(self._temp_dir, temp_filename)
                
                shutil.copy2(source_path, temp_path)
                downloaded_files.append(temp_path)
                
            except Exception as e:
                st.warning(f"Erreur lors de la copie de {clause['name']}: {str(e)}")
                continue
            
            progress_bar.progress((i + 1) / len(selected_clauses))
        
        status_text.text("Copie terminée!")
        return downloaded_files
    
    def cleanup(self):
        """Clean up temporary files"""
        try:
            shutil.rmtree(self._temp_dir, ignore_errors=True)
            # Also cleanup converter temporary files
            self.doc_converter.cleanup()
        except Exception:
            pass