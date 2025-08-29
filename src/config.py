import os
from typing import Optional

class SharePointConfig:
    """Configuration for SharePoint connection"""
    
    def __init__(self):
        self.site_url: str = os.getenv('SHAREPOINT_SITE_URL', '')
        self.username: str = os.getenv('SHAREPOINT_USERNAME', '')
        self.password: str = os.getenv('SHAREPOINT_PASSWORD', '')
        self.document_library: str = os.getenv('SHAREPOINT_DOC_LIBRARY', 'Documents partagés')
        self.clauses_folder: str = os.getenv('SHAREPOINT_CLAUSES_FOLDER', 'Clauses')
    
    def is_configured(self) -> bool:
        """Check if all required configuration is present"""
        return all([self.site_url, self.username, self.password])
    
    @classmethod
    def from_streamlit_secrets(cls) -> 'SharePointConfig':
        """Create config from Streamlit secrets"""
        import streamlit as st
        config = cls()
        if hasattr(st, 'secrets') and 'sharepoint' in st.secrets:
            secrets = st.secrets['sharepoint']
            config.site_url = secrets.get('site_url', '')
            config.username = secrets.get('username', '')
            config.password = secrets.get('password', '')
            config.document_library = secrets.get('document_library', 'Documents partagés')
            config.clauses_folder = secrets.get('clauses_folder', 'Clauses')
        return config