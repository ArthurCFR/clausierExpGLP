# 🎆 Agrégateur de Clauses

Application Streamlit pour l'assemblage automatique de clauses contractuelles développée pour la Direction Juridique du groupe La Poste.

## 🚀 Fonctionnalités

- **Connexion SharePoint** : Accès sécurisé aux documents Word stockés dans SharePoint
- **Catégorisation automatique** : Organisation des clauses par sections contractuelles basée sur parties.ini
- **Interface par sections** : Sélection des clauses organisées par parties du contrat
- **Assemblage ordonné** : Fusion des clauses dans l'ordre des sections contractuelles
- **Tagging intelligent** : Reconnaissance automatique des tags dans les noms de fichiers
- **Mode démo** : Test de l'application sans SharePoint (upload de fichiers locaux)
- **Export personnalisé** : Téléchargement du document final au format Word
- **Conversion automatique** : Support des fichiers Word 97-2003 (.doc) avec conversion automatique

## 📋 Prérequis

- Python 3.8+
- Accès à un site SharePoint avec des documents Word
- Compte utilisateur SharePoint avec permissions de lecture

## 📄 Formats supportés

- **Word moderne (.docx)** : Support natif complet
- **Word 97-2003 (.doc)** : Conversion automatique avec extraction de texte
- **Template** : Utilisation du template fourni pour la mise en page

## 🛠 Installation

1. Clonez le projet :
```bash
git clone <repo-url>
cd Clausier
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

3. Configurez l'accès SharePoint (voir section Configuration)

## ⚙️ Configuration

### Option 1 : Variables d'environnement

Copiez le fichier `.env.example` vers `.env` et modifiez les valeurs :

```bash
cp .env.example .env
```

### Option 2 : Secrets Streamlit

Créez un dossier `.streamlit` et copiez `secrets.toml.example` :

```bash
mkdir .streamlit
cp secrets.toml.example .streamlit/secrets.toml
```

### Option 3 : Saisie manuelle

Utilisez l'interface de l'application pour saisir les informations de connexion.

### Paramètres de configuration

- `SHAREPOINT_SITE_URL` : URL complète de votre site SharePoint
- `SHAREPOINT_USERNAME` : Nom d'utilisateur (email)
- `SHAREPOINT_PASSWORD` : Mot de passe
- `SHAREPOINT_DOC_LIBRARY` : Nom de la bibliothèque de documents (par défaut: "Documents partagés")
- `SHAREPOINT_CLAUSES_FOLDER` : Dossier contenant les clauses (par défaut: "Clauses")

## 🚀 Utilisation

1. Démarrez l'application :
```bash
streamlit run app.py
```

2. Ouvrez votre navigateur à l'adresse affichée (généralement http://localhost:8501)

3. Configurez la connexion SharePoint via le panneau latéral

4. **Organisez vos clauses** : Nommez vos fichiers Word avec des tags pour la catégorisation automatique (voir `examples/clause_naming_examples.md`)

5. **Sélectionnez par sections** : Choisissez les clauses à inclure pour chaque partie du contrat

6. Téléchargez le document final assemblé dans l'ordre des sections

### Mode Démo

Sans connexion SharePoint, vous pouvez tester l'application en uploadant des fichiers Word directement via l'interface.

## 📁 Structure du projet

```
Clausier/
├── app.py                    # Application Streamlit principale
├── config.py                 # Configuration SharePoint
├── sharepoint_client.py      # Client SharePoint avec catégorisation
├── document_merger.py        # Fusion des documents Word
├── parties_parser.py         # Parser du fichier parties.ini
├── parties.ini              # Définition des sections contractuelles
├── requirements.txt          # Dépendances Python
├── .env.example             # Exemple de configuration
├── secrets.toml.example     # Exemple secrets Streamlit
├── examples/                # Exemples et documentation
│   └── clause_naming_examples.md
└── README.md               # Documentation
```

## 🔧 Architecture

### Modules principaux

- **config.py** : Gestion de la configuration SharePoint
- **sharepoint_client.py** : Authentification, téléchargement et catégorisation des clauses
- **parties_parser.py** : Analyse du fichier parties.ini pour les sections contractuelles
- **document_merger.py** : Assemblage des documents Word avec python-docx
- **app.py** : Interface utilisateur Streamlit avec sélection par sections

### Flux de traitement

1. **Authentification** : Connexion à SharePoint via Office365-REST-Python-Client
2. **Découverte** : Listage des fichiers Word dans le dossier des clauses
3. **Catégorisation** : Analyse automatique des noms de fichiers et extraction des tags de section
4. **Organisation** : Regroupement des clauses par sections contractuelles selon parties.ini
5. **Sélection** : Interface utilisateur organisée par sections pour choisir les clauses
6. **Téléchargement** : Récupération des fichiers depuis SharePoint
7. **Fusion** : Assemblage des documents dans l'ordre des sections avec python-docx
8. **Export** : Génération et téléchargement du document final

## 🏷️ Organisation des clauses

### Nommage des fichiers

Pour bénéficier de la catégorisation automatique, nommez vos fichiers Word selon l'une des conventions :

- **Format avec crochets** : `[DESIGNATION] Nom de la clause.docx`
- **Format avec tiret** : `DESIGNATION - Nom de la clause.docx`
- **Format avec inclusion** : `Désignation des Parties - Clause.docx`

### Sections supportées

L'application reconnaît automatiquement 37 sections contractuelles définies dans `parties.ini` :

1. Désignation des Parties
2. Préambule
3. Définitions
4. Objet du Contrat
5. Périmètre géographique
... (voir `parties.ini` pour la liste complète)

### Exemples d'organisation

Consultez le fichier `examples/clause_naming_examples.md` pour des exemples détaillés de nommage et d'organisation de vos clauses dans SharePoint.

## 🛡️ Sécurité

- Les mots de passe ne sont jamais stockés dans le code
- Utilisation de fichiers temporaires pour le traitement
- Nettoyage automatique des fichiers temporaires
- Support des secrets Streamlit pour la production

## 🐛 Dépannage

### Erreur d'authentification
- Vérifiez vos identifiants SharePoint
- Assurez-vous que l'URL du site est correcte
- Vérifiez que l'utilisateur a les permissions de lecture

### Aucune clause trouvée
- Vérifiez le nom de la bibliothèque de documents
- Vérifiez le nom du dossier des clauses
- Assurez-vous que les fichiers sont au format .doc ou .docx

### Clauses non catégorisées
- Vérifiez le nommage de vos fichiers (utilisez les conventions de tags)
- Consultez `examples/clause_naming_examples.md` pour les formats supportés
- Les clauses mal nommées apparaîtront dans la section "Clauses non catégorisées"

### Problèmes avec fichiers .doc
- **Conversion automatique** : Les fichiers Word 97-2003 sont automatiquement convertis
- **Qualité de conversion** : Le texte est extrait, la mise en forme peut être simplifiée
- **Alternative** : Pour une meilleure qualité, convertissez manuellement en .docx avec Word/LibreOffice

### Erreur de fusion
- Vérifiez que les fichiers Word ne sont pas corrompus
- Assurez-vous qu'il n'y a pas de fichiers protégés par mot de passe

## 📄 Licence

Ce projet est fourni tel quel à des fins éducatives et de démonstration.