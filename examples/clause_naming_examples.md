# Exemples de nommage des clauses

Pour que l'application puisse automatiquement catégoriser vos clauses, vous devez nommer vos fichiers Word (.doc ou .docx) selon l'une des conventions suivantes :

## Conventions de nommage supportées

### Format 1 : [TAG] Nom de la clause
```
[DESIGNATION] Clause d'identification des parties.docx
[PREAMBULE] Contexte et historique du contrat.doc
[DEFINITIONS] Terminologie contractuelle.docx
[OBJET] Description de l'objet du contrat.doc
```

### Format 2 : TAG - Nom de la clause
```
DESIGNATION - Identification des parties contractantes.docx
PREAMBULE - Contexte du projet.doc
DEFINITIONS - Glossaire des termes.docx
OBJET - Périmètre de la prestation.doc
```

### Format 3 : Inclusion du nom de section
```
Désignation des Parties - Clause principale.docx
Préambule du contrat logistique.doc
Définitions et terminologie.docx
Objet du contrat de service.doc
```

## Tags recommandés (basés sur parties.ini)

| Section | Tag suggéré | Exemple de fichier |
|---------|-------------|-------------------|
| Désignation des Parties | DESIGNATION | `[DESIGNATION] Parties contractantes.docx` |
| Préambule | PREAMBULE | `[PREAMBULE] Contexte du projet.docx` |
| Définitions | DEFINITIONS | `[DEFINITIONS] Glossaire technique.docx` |
| Objet du Contrat | OBJET | `[OBJET] Périmètre de la prestation.docx` |
| Périmètre géographique | GEOGRAPHIE | `[GEOGRAPHIE] Zone d'intervention.docx` |
| Conditions tarifaires | TARIFS | `[TARIFS] Grille de prix.docx` |
| Facturation – Paiement | FACTURATION | `[FACTURATION] Modalités de règlement.docx` |
| Responsabilité | RESPONSABILITE | `[RESPONSABILITE] Limitation des risques.docx` |
| Confidentialité | CONFIDENTIALITE | `[CONFIDENTIALITE] Clause de non-divulgation.docx` |
| Durée – Entrée en vigueur | DUREE | `[DUREE] Période contractuelle.docx` |
| Résiliation | RESILIATION | `[RESILIATION] Conditions de rupture.docx` |

## Bonnes pratiques

1. **Utilisez des tags courts** : préférez "DESIGNATION" à "DESIGNATION_DES_PARTIES"
2. **Soyez cohérents** : utilisez toujours le même format dans votre bibliothèque SharePoint
3. **Noms descriptifs** : après le tag, utilisez un nom explicite pour la clause
4. **Évitez les caractères spéciaux** : limitez-vous aux lettres, chiffres, espaces et tirets

## Gestion des clauses non catégorisées

Les clauses qui ne correspondent à aucun tag seront placées dans la section "Clauses non catégorisées" et apparaîtront à la fin du document final.

## Exemple complet d'organisation SharePoint

```
📁 Clauses/
├── [DESIGNATION] Identification société cliente.docx
├── [DESIGNATION] Identification prestataire logistique.docx
├── [PREAMBULE] Contexte commercial.docx
├── [DEFINITIONS] Termes logistiques.docx
├── [DEFINITIONS] Terminologie produits.docx
├── [OBJET] Prestation d'entreposage.docx
├── [OBJET] Service de distribution.docx
├── [TARIFS] Grille tarifaire entreposage.docx
├── [TARIFS] Coûts de transport.docx
├── [FACTURATION] Modalités de paiement.docx
├── [RESILIATION] Résiliation pour manquement.docx
└── [RESILIATION] Résiliation de convenance.docx
```