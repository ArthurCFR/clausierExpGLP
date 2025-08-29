import os
import tempfile
import re
from typing import List, Dict, Optional
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import streamlit as st
from config import SharePointConfig
from parties_parser import PartiesParser

class SharePointClient:
    """Client for interacting with SharePoint documents"""
    
    def __init__(self, config: SharePointConfig):
        self.config = config
        self.ctx: Optional[ClientContext] = None
        self._temp_dir = tempfile.mkdtemp()
        self.parties_parser = PartiesParser()
    
    def authenticate(self) -> bool:
        """Authenticate with SharePoint"""
        try:
            credentials = UserCredential(self.config.username, self.config.password)
            self.ctx = ClientContext(self.config.site_url).with_credentials(credentials)
            # Test connection
            self.ctx.web.get().execute_query()
            return True
        except Exception as e:
            st.error(f"Erreur d'authentification SharePoint: {str(e)}")
            return False
    
    def get_clause_files(self) -> List[Dict[str, str]]:
        """Get list of clause files from SharePoint"""
        if not self.ctx:
            if not self.authenticate():
                return []
        
        try:
            folder_path = f"/{self.config.document_library}/{self.config.clauses_folder}"
            folder = self.ctx.web.get_folder_by_server_relative_url(folder_path)
            files = folder.files.get().execute_query()
            
            clause_files = []
            for file in files:
                if file.name.endswith(('.doc', '.docx')):
                    clause_name = file.name.replace('.docx', '').replace('.doc', '')
                    section_tag = self._extract_section_tag(clause_name)
                    clause_files.append({
                        'name': clause_name,
                        'file_name': file.name,
                        'server_relative_url': file.serverRelativeUrl,
                        'section_tag': section_tag,
                        'section_order': self.parties_parser.get_section_order(section_tag) if section_tag else 999
                    })
            
            return sorted(clause_files, key=lambda x: (x['section_order'], x['name']))
        
        except Exception as e:
            st.error(f"Erreur lors de la récupération des clauses: {str(e)}")
            return []
    
    def download_clause_file(self, server_relative_url: str, file_name: str) -> Optional[str]:
        """Download a clause file and return local path"""
        if not self.ctx:
            return None
        
        try:
            local_path = os.path.join(self._temp_dir, file_name)
            file = self.ctx.web.get_file_by_server_relative_url(server_relative_url)
            
            with open(local_path, 'wb') as local_file:
                file.download(local_file).execute_query()
            
            return local_path
        
        except Exception as e:
            st.error(f"Erreur lors du téléchargement de {file_name}: {str(e)}")
            return None
    
    def download_selected_clauses(self, selected_clauses: List[Dict[str, str]]) -> List[str]:
        """Download multiple clause files and return list of local paths"""
        downloaded_files = []
        
        for i, clause in enumerate(selected_clauses):
            local_path = self.download_clause_file(
                clause['server_relative_url'], 
                clause['file_name']
            )
            
            if local_path:
                downloaded_files.append(local_path)
        return downloaded_files
    
    def _extract_section_tag(self, clause_name: str) -> Optional[str]:
        """
        Extract section tag from clause name
        Expected format: [TAG] Clause Name or TAG - Clause Name
        """
        # Pattern 1: [TAG] Clause Name
        match = re.match(r'\[([^\]]+)\]\s*(.+)', clause_name)
        if match:
            tag = match.group(1).strip().lower()
            return self._normalize_tag(tag)
        
        # Pattern 2: TAG - Clause Name
        match = re.match(r'^([^-]+?)\s*-\s*(.+)', clause_name)
        if match:
            tag = match.group(1).strip().lower()
            return self._normalize_tag(tag)
        
        # Pattern 3: Try to match section names directly
        for section in self.parties_parser.get_sections():
            section_name = section['name'].lower()
            if section_name in clause_name.lower():
                return section['key']
        
        return None
    
    def _normalize_tag(self, tag: str) -> Optional[str]:
        """Normalize tag to match section keys"""
        # Remove accents and special characters
        normalized = re.sub(r'[^\w\s-]', '', tag.lower())
        normalized = re.sub(r'[-\s]+', '_', normalized)
        
        # Try exact match first
        section = self.parties_parser.find_section_by_key(normalized)
        if section:
            return section['key']
        
        # Try fuzzy matching with section names
        for section in self.parties_parser.get_sections():
            if normalized in section['key'] or section['key'] in normalized:
                return section['key']
        
        return normalized
    
    def get_clauses_by_section(self) -> Dict[str, List[Dict[str, str]]]:
        """Get clauses grouped by section"""
        all_clauses = self.get_clause_files()
        clauses_by_section = {}
        
        # Initialize with all sections
        for section in self.parties_parser.get_sections():
            clauses_by_section[section['key']] = []
        
        # Add uncategorized section
        clauses_by_section['uncategorized'] = []
        
        # Group clauses by section
        for clause in all_clauses:
            section_key = clause.get('section_tag', 'uncategorized')
            if section_key not in clauses_by_section:
                clauses_by_section[section_key] = []
            clauses_by_section[section_key].append(clause)
        
        return clauses_by_section
    
    def cleanup(self):
        """Clean up temporary files"""
        try:
            import shutil
            shutil.rmtree(self._temp_dir, ignore_errors=True)
        except Exception:
            pass