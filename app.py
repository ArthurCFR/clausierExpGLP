import streamlit as st
import os
import tempfile
from datetime import datetime
from config import SharePointConfig
from sharepoint_client import SharePointClient
from local_client import LocalClauseClient
from document_merger import DocumentMerger
from parties_parser import PartiesParser
from doc_converter import DocConverter
from docx import Document
import html

def main():
    st.set_page_config(
        page_title="Agrégateur de clauses",
        page_icon="📄",
        layout="wide"
    )
    
    # Intro landing screen state
    if 'show_intro' not in st.session_state:
        st.session_state.show_intro = True

    # Landing page (only these elements, centered, blur-in 1s)
    if st.session_state.show_intro:
        st.markdown(
            """
<style>
/* Reset default paddings around app */
section.main > div {padding-top: 0 !important;}

.hero-wrap { 
  display: flex; 
  flex-direction: column; 
  align-items: center; 
  text-align: center;
  padding: 40px 20px 20px 20px;
}

@keyframes blurIn { from { filter: blur(20px); opacity: 0; } to { filter: blur(0); opacity: 1; } }
@keyframes fadeOut { from { opacity: 1; } to { opacity: 0; } }

.hero {
  animation: blurIn 1s ease-out both;
}

.hero-content {
  animation: blurIn 1s ease-out both;
}

.white-overlay {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: white;
  z-index: 9999;
  animation: fadeOut 1s ease-out 0.5s both;
  pointer-events: none;
}

.hero h1 { 
  font-size: 2.5rem; 
  line-height: 1.05; 
  margin: 0 0 10px 0; 
  color: #003DA5; 
  font-weight: 800; 
  font-family: 'Montserrat ExtraBold', Montserrat, system-ui, -apple-system, Segoe UI, Roboto, 'Helvetica Neue', Arial, sans-serif;
}

.hero h3 { 
  font-size: 1.2rem; 
  font-style: italic; 
  color: #003DA5; 
  margin: 0 0 15px 0; 
}

/* Style the start contract button */
div.stButton > button { 
  margin: 0; 
  padding: 0.8rem 1.6rem; 
  font-size: 1.1rem; 
  border-radius: 8px; 
  background: #003DA5; 
  color: #fff; 
  border: none; 
  width: 100%;
  animation: blurIn 1s ease-out both;
}

div.stButton > button:hover,
div.stButton > button:focus,
div.stButton > button:active {
  background: #003DA5 !important;
  color: #fff !important;
  border: none !important;
}

.logo-container {
  margin-top: 30px;
  display: flex;
  justify-content: center;
  animation: blurIn 1s ease-out both;
}

.logo-container img {
  max-width: 200px;
  height: auto;
}

</style>

<div class="hero-wrap">
  <div class="hero">
    <h1>AGRÉGATEUR DE CLAUSES</h1>
    <h3><em>Expérimentations IA de la Direction Juridique du groupe La Poste</em></h3>
  </div>
</div>
<div class="white-overlay"></div>
""",
            unsafe_allow_html=True,
        )

        # Logo La Poste (affiché avant le bouton pour éviter le flash)
        st.markdown(
            """
            <div class="logo-container">
                <img src="data:image/png;base64,{}" alt="Logo La Poste">
            </div>
            """.format(_get_base64_image("la_poste-logo-freelogovectors.net_.png")),
            unsafe_allow_html=True
        )
        
        # Use Streamlit columns with better proportions for true centering
        col1, col2, col3 = st.columns([2, 1, 2])
        with col2:
            if st.button("Commencer un contrat", key="start_contract", use_container_width=True):
                st.session_state.show_intro = False
                st.rerun()
        return

    st.title("🎆 Agrégateur de clauses")
    st.markdown("*Sélectionnez et assemblez vos clauses contractuelles par section*")
    
    # CSS for styling multiselect selected items in green
    st.markdown(
        """
        <style>
        /* Style selected items in multiselect */
        .stMultiSelect div[data-baseweb="select"] div[data-baseweb="tag"] {
            background-color: #28a745 !important;
            color: white !important;
        }
        .stMultiSelect div[data-baseweb="select"] div[data-baseweb="tag"] span {
            color: white !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    
    # Initialize session state
    if 'sharepoint_client' not in st.session_state:
        st.session_state.sharepoint_client = None
    if 'local_client' not in st.session_state:
        st.session_state.local_client = None
    if 'clause_files' not in st.session_state:
        st.session_state.clause_files = []
    if 'clauses_by_section' not in st.session_state:
        st.session_state.clauses_by_section = {}
    if 'merger' not in st.session_state:
        st.session_state.merger = DocumentMerger(enable_summary=False)
    if 'parties_parser' not in st.session_state:
        st.session_state.parties_parser = PartiesParser()
    if 'connection_mode' not in st.session_state:
        st.session_state.connection_mode = "local"
    if 'preview_converter' not in st.session_state:
        st.session_state.preview_converter = DocConverter()
    
    # Inline preview panel (scrollable, non-disabled)
    show_preview = (
        st.session_state.get('preview_content') and 
        not st.session_state.get('hide_preview', False)
    )
    
    if show_preview:
        st.markdown("### Aperçu de la clause")
        st.markdown(f"**{st.session_state.get('preview_title', 'Clause')}**")
        safe_html = html.escape(st.session_state['preview_content'])
        st.markdown(
            f"""
<div style="max-height: 320px; overflow-y: auto; padding: 8px; border: 1px solid #e6e6e6; border-radius: 6px; background: #fafafa; white-space: pre-wrap; line-height: 1.45; margin-bottom: 8px;">
{safe_html}
</div>
""",
            unsafe_allow_html=True,
        )
        if st.button("Fermer l'aperçu", key="close_preview_panel"):
            # Clean up preview state without rerun
            for key in ['preview_content', 'preview_title', 'hide_preview']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Connection mode selection
        connection_mode = st.radio(
            "Mode de connexion:",
            ["Dossiers locaux", "SharePoint"],
            key="connection_mode_radio"
        )
        new_mode = connection_mode.lower().replace(" ", "_").replace("dossiers_locaux", "local").replace("sharepoint", "sharepoint")
        
        # Clear data when switching modes
        if new_mode != st.session_state.connection_mode:
            st.session_state.local_client = None
            st.session_state.sharepoint_client = None
            st.session_state.clause_files = []
            st.session_state.clauses_by_section = {}
        
        st.session_state.connection_mode = new_mode
        
        if st.session_state.connection_mode == "local":
            st.info("📁 Mode local activé")
            st.markdown("Les clauses seront lues depuis le dossier `clauses/`")
            
            # Auto-load clauses when switching to local mode
            if not st.session_state.local_client:
                st.session_state.local_client = LocalClauseClient()
                st.session_state.clause_files = st.session_state.local_client.get_clause_files()
                st.session_state.clauses_by_section = st.session_state.local_client.get_clauses_by_section()
            
            # Show status
            if st.session_state.clause_files:
                st.success(f"✅ {len(st.session_state.clause_files)} clauses chargées automatiquement!")
            else:
                st.warning("⚠️ Aucune clause trouvée dans le dossier local")
            
            # Optional reload button
            if st.button("🔄 Recharger les clauses locales"):
                st.session_state.clause_files = st.session_state.local_client.get_clause_files()
                st.session_state.clauses_by_section = st.session_state.local_client.get_clauses_by_section()
                if st.session_state.clause_files:
                    st.success(f"✅ {len(st.session_state.clause_files)} clauses rechargées!")
                else:
                    st.warning("⚠️ Aucune clause trouvée dans le dossier local")
                st.rerun()
        
        else:
            # SharePoint configuration options
            config_method = st.radio(
                "Méthode de configuration SharePoint:",
                ["Variables d'environnement", "Saisie manuelle", "Secrets Streamlit"]
            )
            
            if config_method == "Saisie manuelle":
                site_url = st.text_input("URL du site SharePoint:", placeholder="https://monentreprise.sharepoint.com/sites/monsite")
                username = st.text_input("Nom d'utilisateur:", placeholder="utilisateur@monentreprise.com")
                password = st.text_input("Mot de passe:", type="password")
                doc_library = st.text_input("Bibliothèque de documents:", value="Documents partagés")
                clauses_folder = st.text_input("Dossier des clauses:", value="Clauses")
                
                config = SharePointConfig()
                config.site_url = site_url
                config.username = username
                config.password = password
                config.document_library = doc_library
                config.clauses_folder = clauses_folder
                
            elif config_method == "Secrets Streamlit":
                config = SharePointConfig.from_streamlit_secrets()
                st.info("Configuration chargée depuis les secrets Streamlit")
                
            else:
                config = SharePointConfig()
                st.info("Configuration chargée depuis les variables d'environnement")
            
            # SharePoint connection button
            if st.button("🔌 Se connecter à SharePoint"):
                if config.is_configured():
                    st.session_state.sharepoint_client = SharePointClient(config)
                    if st.session_state.sharepoint_client.authenticate():
                        st.success("✅ Connexion réussie!")
                        st.session_state.clause_files = st.session_state.sharepoint_client.get_clause_files()
                        st.session_state.clauses_by_section = st.session_state.sharepoint_client.get_clauses_by_section()
                    else:
                        st.error("❌ Échec de la connexion")
                else:
                    st.error("❌ Configuration incomplète")
    
    # Main content
    active_client = None
    if st.session_state.connection_mode == "local" and st.session_state.local_client:
        active_client = st.session_state.local_client
    elif st.session_state.connection_mode == "sharepoint" and st.session_state.sharepoint_client:
        active_client = st.session_state.sharepoint_client
    
    if not active_client:
        if st.session_state.connection_mode == "local":
            st.info("👈 Veuillez charger les clauses locales via le panneau latéral")
        else:
            st.info("👈 Veuillez vous connecter à SharePoint via le panneau latéral")
        
        # Demo section for testing without SharePoint
        st.markdown("---")
        st.subheader("🧪 Mode Démo")
        st.markdown("*Pour tester l'application sans SharePoint, vous pouvez uploader des fichiers Word*")
        
        demo_files = st.file_uploader(
            "Uploadez des fichiers Word (.docx):",
            type=['docx'],
            accept_multiple_files=True
        )
        
        if demo_files:
            st.success(f"✅ {len(demo_files)} fichier(s) uploadé(s)")
            
            # Create demo clause list
            demo_clauses = []
            for file in demo_files:
                demo_clauses.append({
                    'name': file.name.replace('.docx', ''),
                    'file_name': file.name,
                    'file_obj': file
                })
            
            selected_demo_clauses = st.multiselect(
                "Sélectionnez les clauses à assembler:",
                options=[clause['name'] for clause in demo_clauses],
                default=[clause['name'] for clause in demo_clauses]
            )
            
            if selected_demo_clauses and st.button("🔗 Assembler les clauses (Mode Démo)"):
                # Save uploaded files temporarily
                temp_files = []
                selected_names = []
                
                for clause in demo_clauses:
                    if clause['name'] in selected_demo_clauses:
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                        temp_file.write(clause['file_obj'].read())
                        temp_file.close()
                        temp_files.append(temp_file.name)
                        selected_names.append(clause['name'])
                
                # Merge documents
                try:
                    merged_doc_path = st.session_state.merger.merge_documents(temp_files, selected_names)
                    
                    # Offer download
                    with open(merged_doc_path, 'rb') as f:
                        st.download_button(
                            label="📥 Télécharger le document assemblé",
                            data=f.read(),
                            file_name=f"clauses_assemblees_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    st.success("✅ Document assemblé avec succès!")
                    
                    # Cleanup
                    for temp_file in temp_files:
                        try:
                            os.unlink(temp_file)
                        except:
                            pass
                            
                except Exception as e:
                    st.error(f"❌ Erreur lors de l'assemblage: {str(e)}")
        
    else:
        # Active client mode (local or SharePoint)
        # Check if we have clauses to display
        has_clauses = (
            st.session_state.clauses_by_section and 
            any(clauses for clauses in st.session_state.clauses_by_section.values())
        )
        
        if has_clauses:

            # Create selection interface by sections
            selected_clauses_all = []
            
            # Use tabs for better organization
            sections = st.session_state.parties_parser.get_sections()
            
            # Create columns for better layout
            col1, col2 = st.columns([2, 1])
            
            with col1:
                # Selection by sections
                selected_clauses_all = []
                for section in sections:
                    section_clauses = st.session_state.clauses_by_section.get(section['key'], [])
                    
                    if section_clauses:
                        st.markdown(f"### {section['order']}. {section['name']}")
                        
                        # Multiselect for selection
                        clause_options = []
                        clause_mapping = {}
                        for clause in section_clauses:
                            if clause['file_name'].endswith('.doc'):
                                legacy_indicator = " ⚠️"
                                option_label = f"{clause['name']}{legacy_indicator}"
                            else:
                                option_label = clause['name']
                            clause_options.append(option_label)
                            clause_mapping[option_label] = clause
                        selected_for_section = st.multiselect(
                            f"Clauses pour: {section['name']}",
                            options=clause_options,
                            key=f"section_{section['key']}",
                            help=f"Sélectionnez les clauses pour la section '{section['name']}'. ⚠️ = fichier .doc non supporté"
                        )
                        
                        # Add selected clauses to the main list
                        for label in selected_for_section:
                            base_name = label.replace(' ⚠️', '')
                            clause_obj = next((c for c in section_clauses if c['name'] == base_name), None)
                            if clause_obj:
                                selected_clauses_all.append(clause_obj)
                        
                        st.markdown("---")
                
                # Handle uncategorized clauses
                uncategorized_clauses = st.session_state.clauses_by_section.get('uncategorized', [])
                if uncategorized_clauses:
                    st.markdown("### 📝 Clauses non catégorisées")
                    
                    # Create options with file type indicators for uncategorized
                    uncategorized_options = []
                    uncategorized_mapping = {}
                    
                    # Multiselect for selection (uncategorized)
                    uncategorized_options = []
                    uncategorized_mapping = {}
                    for clause in uncategorized_clauses:
                        if clause['file_name'].endswith('.doc'):
                            legacy_indicator = " ⚠️"
                            option_label = f"{clause['name']}{legacy_indicator}"
                        else:
                            option_label = clause['name']
                        uncategorized_options.append(option_label)
                        uncategorized_mapping[option_label] = clause

                    selected_uncategorized = st.multiselect(
                        "Clauses non catégorisées:",
                        options=uncategorized_options,
                        key="section_uncategorized",
                        help="Clauses qui n'ont pas pu être automatiquement assignées à une section. ⚠️ = fichier .doc non supporté"
                    )
                    
                    # Add selected uncategorized clauses to the main list
                    for label in selected_uncategorized:
                        base_name = label.replace(' ⚠️', '')
                        clause_obj = next((c for c in uncategorized_clauses if c['name'] == base_name), None)
                        if clause_obj:
                            selected_clauses_all.append(clause_obj)

            
            with col2:
                # Summary sidebar
                st.markdown("### 📊 Résumé de sélection")
                
                # Contract preview button
                if selected_clauses_all:
                    if st.button("📋 Aperçu du contrat", key="contract_preview_btn", use_container_width=True):
                        contract_preview = _generate_contract_preview(selected_clauses_all)
                        st.session_state['preview_title'] = "Aperçu du contrat complet"
                        st.session_state['preview_content'] = contract_preview
                        st.session_state.hide_preview = False
                        st.rerun()
                
                if selected_clauses_all:
                    st.success(f"**{len(selected_clauses_all)} clause(s) sélectionnée(s)**")
                    
                    # Group selected clauses by section for display
                    selected_by_section = {}
                    for clause in selected_clauses_all:
                        section_key = clause.get('section_tag', 'uncategorized')
                        if section_key not in selected_by_section:
                            selected_by_section[section_key] = []
                        selected_by_section[section_key].append(clause['name'])
                    
                    # Display selection summary with compact preview buttons
                    for section in sections:
                        if section['key'] in selected_by_section:
                            st.write(f"**{section['name']}:**")
                            for idx, clause_name in enumerate(selected_by_section[section['key']]):
                                cols_item = st.columns([0.9, 0.1])
                                with cols_item[0]:
                                    st.write(f"• {clause_name}")
                                with cols_item[1]:
                                    if st.button("👁️", key=f"sum_prev_{section['key']}_{idx}", help="Aperçu"):
                                        clause_obj = next((c for c in st.session_state.clauses_by_section.get(section['key'], []) if c['name'] == clause_name), None)
                                        if clause_obj:
                                            preview_text = _get_clause_preview(clause_obj)
                                            st.session_state['preview_title'] = clause_obj['name']
                                            st.session_state['preview_content'] = preview_text or "(Aucun aperçu disponible)"
                                            st.session_state.hide_preview = False
                                            st.rerun()
                    
                    if 'uncategorized' in selected_by_section:
                        st.write("**Non catégorisées:**")
                        for idx, clause_name in enumerate(selected_by_section['uncategorized']):
                            cols_item = st.columns([0.9, 0.1])
                            with cols_item[0]:
                                st.write(f"• {clause_name}")
                            with cols_item[1]:
                                if st.button("👁️", key=f"sum_prev_uncat_{idx}", help="Aperçu"):
                                    clause_obj = next((c for c in st.session_state.clauses_by_section.get('uncategorized', []) if c['name'] == clause_name), None)
                                    if clause_obj:
                                        preview_text = _get_clause_preview(clause_obj)
                                        st.session_state['preview_title'] = clause_obj['name']
                                        st.session_state['preview_content'] = preview_text or "(Aucun aperçu disponible)"
                                        st.session_state.hide_preview = False
                                        st.rerun()
                else:
                    st.info("Aucune clause sélectionnée")
                
                st.markdown("---")
                
                # Configuration options
                
                # Initialize AI summary toggle state if not exists
                if 'ai_summary_enabled' not in st.session_state:
                    st.session_state.ai_summary_enabled = st.session_state.merger.enable_summary
                    
                enable_summary = st.toggle(
                    "Générer une synthèse IA du contrat", 
                    value=st.session_state.ai_summary_enabled,
                    key="ai_summary_toggle"
                )
                
                # Only update if actually changed
                if enable_summary != st.session_state.ai_summary_enabled:
                    st.session_state.ai_summary_enabled = enable_summary
                    st.session_state.merger.enable_summary = enable_summary
                custom_filename = st.text_input(
                    "Nom du fichier (optionnel):",
                    placeholder="document_final"
                )
                
            # Assembly button outside the columns
            if selected_clauses_all and st.button("🔗 Assembler les clauses", type="primary"):
                # Sort selected clauses by section order
                selected_clauses_all.sort(key=lambda x: (x.get('section_order', 999), x['name']))
                
                # Download/copy files
                if st.session_state.connection_mode == "local":
                    with st.spinner("Copie des clauses locales..."):
                        downloaded_files = active_client.download_selected_clauses(selected_clauses_all)
                else:
                    with st.spinner("Téléchargement des clauses depuis SharePoint..."):
                        downloaded_files = active_client.download_selected_clauses(selected_clauses_all)
                    
                if downloaded_files or st.session_state.connection_mode == "local":
                    # Merge documents using section-based approach
                    try:
                        with st.spinner("Assemblage des clauses avec template..."):
                            # Organize selected clauses by section
                            selected_by_section = {}
                            for clause in selected_clauses_all:
                                section_key = clause.get('section_tag', 'uncategorized')
                                if section_key not in selected_by_section:
                                    selected_by_section[section_key] = []
                                selected_by_section[section_key].append(clause)
                            
                            # Get sections in order
                            sections_order = st.session_state.parties_parser.get_sections()
                            
                            # Use new section-based merge method
                            merged_doc_path = st.session_state.merger.merge_documents_by_sections(
                                selected_by_section,
                                sections_order
                            )
                            
                            # Generate filename - simple format with custom name + date
                            if custom_filename:
                                filename = f"{custom_filename}_{datetime.now().strftime('%Y%m%d')}.docx"
                            else:
                                filename = f"document_{datetime.now().strftime('%Y%m%d')}.docx"
                            
                            # Offer download
                            with open(merged_doc_path, 'rb') as f:
                                st.download_button(
                                    label="📥 Télécharger le document assemblé",
                                    data=f.read(),
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                            
                            st.success("✅ Document assemblé avec succès!")
                            st.balloons()
                            
                    except Exception as e:
                        st.error(f"❌ Erreur lors de l'assemblage: {str(e)}")
                else:
                    st.error("❌ Aucun fichier n'a pu être téléchargé")
        
        else:
            if st.session_state.connection_mode == "local":
                st.info("📝 Aucune clause disponible pour sélection")
                st.markdown("""
                **Pour ajouter des clauses :**
                1. Placez vos fichiers Word (.doc ou .docx) dans les dossiers correspondants sous `clauses/`
                2. Exemple : `clauses/01_Designation_des_Parties/ma_clause.docx`
                3. Cliquez sur "🔄 Recharger les clauses locales" dans le panneau latéral
                """)
            else:
                st.warning("⚠️ Aucune clause trouvée. Vérifiez la configuration du dossier SharePoint.")
                
                if st.button("🔄 Recharger les clauses"):
                    if active_client:
                        st.session_state.clause_files = active_client.get_clause_files()
                        st.session_state.clauses_by_section = active_client.get_clauses_by_section()
                        st.rerun()
    
    # Footer
    st.markdown("---")
    st.markdown("*Clausier v1.0 - Assembleur de clauses contractuelles*")

def _get_clause_preview(clause: dict) -> str:
    """Return a short text preview of a clause (.docx directly or .doc via conversion)."""
    try:
        path = clause.get('file_path') or clause.get('local_path')
        if not path:
            return ""
        if path.endswith('.doc'):
            # Convert to docx temporarily
            converter = st.session_state.get('preview_converter')
            docx_path = converter.convert_doc_to_docx(path)
            return _extract_docx_text(docx_path)
        else:
            return _extract_docx_text(path)
    except Exception as e:
        return f"Erreur d'aperçu: {str(e)}"


def _generate_contract_preview(selected_clauses: list) -> str:
    """Generate a complete contract preview by concatenating all selected clauses."""
    try:
        sections = st.session_state.parties_parser.get_sections()
        
        # Organize clauses by section
        clauses_by_section = {}
        for clause in selected_clauses:
            section_key = clause.get('section_tag', 'uncategorized')
            if section_key not in clauses_by_section:
                clauses_by_section[section_key] = []
            clauses_by_section[section_key].append(clause)
        
        contract_parts = []
        contract_parts.append("=== APERÇU DU CONTRAT COMPLET ===\n")
        
        # Process sections in order
        for i, section in enumerate(sections):
            section_key = section['key']
            
            if section_key in clauses_by_section:
                contract_parts.append(f"\n**--- {section['order']}. {section['name']} ---**\n")
                
                for clause in clauses_by_section[section_key]:
                    contract_parts.append(f"\n[{clause['name']}]\n")
                    clause_content = _get_clause_preview(clause)
                    if clause_content:
                        contract_parts.append(clause_content)
                    else:
                        contract_parts.append("(Contenu non disponible)")
                    contract_parts.append("\n")
                
                # Add separator after non-empty section
                contract_parts.append("="*50 + "\n")
            else:
                contract_parts.append(f"\n**--- {section['order']}. {section['name']} ---** : aucune clause sélectionnée\n")
                
                # Add separator after empty section if next section has content
                next_section_has_content = False
                for j in range(i + 1, len(sections)):
                    next_section_key = sections[j]['key']
                    if next_section_key in clauses_by_section:
                        next_section_has_content = True
                        break
                
                # Also check if uncategorized section exists
                if not next_section_has_content and 'uncategorized' in clauses_by_section:
                    next_section_has_content = True
                
                if next_section_has_content:
                    contract_parts.append("="*50 + "\n")
        
        # Add uncategorized clauses at the end
        if 'uncategorized' in clauses_by_section:
            contract_parts.append("\n**--- Clauses non catégorisées ---**\n")
            for clause in clauses_by_section['uncategorized']:
                contract_parts.append(f"\n[{clause['name']}]\n")
                clause_content = _get_clause_preview(clause)
                if clause_content:
                    contract_parts.append(clause_content)
                else:
                    contract_parts.append("(Contenu non disponible)")
                contract_parts.append("\n")
            
            # Add separator after uncategorized section if it has clauses
            contract_parts.append("="*50 + "\n")
        
        return "\n".join(contract_parts)
        
    except Exception as e:
        return f"Erreur lors de la génération de l'aperçu du contrat: {str(e)}"


def _extract_docx_text(docx_path: str, max_chars: int = None) -> str:
    try:
        doc = Document(docx_path)
        texts = []
        for p in doc.paragraphs:
            if p.text and p.text.strip():
                texts.append(p.text.strip())
            if max_chars is not None and sum(len(t) for t in texts) > max_chars:
                break
        content = '\n'.join(texts)
        if max_chars is not None and len(content) > max_chars:
            content = content[:max_chars] + '…'
        return content
    except Exception as e:
        return f"Erreur de lecture .docx: {str(e)}"

def _get_base64_image(image_path: str) -> str:
    """Convert image to base64 string for embedding in HTML"""
    import base64
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return ""

if __name__ == "__main__":
    main()