import os
from typing import List, Dict, Optional

class PartiesParser:
    """Parser for the parties.ini file to extract contract sections"""
    
    def __init__(self, parties_file_path: str = "parties.ini"):
        self.parties_file_path = parties_file_path
        self.sections: List[Dict[str, any]] = []
        self._load_sections()
    
    def _load_sections(self):
        """Load sections from parties.ini file"""
        if not os.path.exists(self.parties_file_path):
            # Default sections if file doesn't exist
            self.sections = [
                {"order": 1, "name": "Désignation des Parties", "key": "designation_parties"},
                {"order": 2, "name": "Préambule", "key": "preambule"},
                {"order": 3, "name": "Définitions", "key": "definitions"}
            ]
            return
        
        try:
            with open(self.parties_file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            self.sections = []
            order = 1
            
            for line in lines:
                line = line.strip()
                
                if line and '→' in line:
                    # Parse line format: "1→Désignation des Parties"
                    parts = line.split('→', 1)
                    if len(parts) == 2:
                        try:
                            order_num = int(parts[0].strip())
                            name = parts[1].strip()
                            if name:  # Skip empty names
                                key = self._generate_key(name)
                                self.sections.append({
                                    "order": order_num,
                                    "name": name,
                                    "key": key
                                })
                        except ValueError:
                            continue
                
                elif line and not line.isspace():
                    # Parse simple format: just the section name
                    name = line.strip()
                    if name:  # Skip empty names
                        key = self._generate_key(name)
                        self.sections.append({
                            "order": order,
                            "name": name,
                            "key": key
                        })
                        order += 1
            
            # Sort by order
            self.sections.sort(key=lambda x: x['order'])
            
        except Exception as e:
            print(f"Error loading parties.ini: {e}")
            self.sections = []
    
    def _generate_key(self, name: str) -> str:
        """Generate a key from section name"""
        import re
        # Remove accents and special characters, replace spaces with underscores
        key = re.sub(r'[^\w\s-]', '', name.lower())
        key = re.sub(r'[-\s]+', '_', key)
        return key
    
    def get_sections(self) -> List[Dict[str, any]]:
        """Get all contract sections"""
        return self.sections.copy()
    
    def get_section_names(self) -> List[str]:
        """Get list of section names in order"""
        return [section['name'] for section in self.sections]
    
    def get_section_keys(self) -> List[str]:
        """Get list of section keys in order"""
        return [section['key'] for section in self.sections]
    
    def find_section_by_key(self, key: str) -> Optional[Dict[str, any]]:
        """Find section by key"""
        for section in self.sections:
            if section['key'] == key:
                return section
        return None
    
    def find_section_by_name(self, name: str) -> Optional[Dict[str, any]]:
        """Find section by name"""
        for section in self.sections:
            if section['name'].lower() == name.lower():
                return section
        return None
    
    def get_section_order(self, key: str) -> int:
        """Get the order of a section by key"""
        section = self.find_section_by_key(key)
        return section['order'] if section else 999