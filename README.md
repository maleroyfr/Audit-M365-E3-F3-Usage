# Audit-M365-E3-F3-Usage

Audit de l'usage des licences **Microsoft 365 E3, F3 et E5** (y compris les variantes EEA / sans Teams / HUB) via Microsoft Graph.

Le script collecte les rapports d'activité Microsoft 365 et produit des fichiers CSV détaillés permettant d'identifier les licences sous-utilisées ou non utilisées.

> ℹ️ La console et les colonnes CSV sont en **français**. Le code source est en anglais.

---

## Prérequis

### Module PowerShell

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Permissions Microsoft Graph requises

| Permission | Usage |
|---|---|
| `Reports.Read.All` | Téléchargement des rapports d'usage |
| `User.Read.All` | Lecture des utilisateurs et de leurs licences |
| `Organization.Read.All` | Lecture des informations du tenant |
| `Directory.Read.All` | Lecture des licences souscrites (SKUs) |

> ⚠️ Si les utilisateurs apparaissent anonymisés dans les rapports, vérifiez le paramètre **Rapports > Confidentialité** dans le centre d'administration Microsoft 365.

---

## Utilisation

### Exécution simple (période par défaut : 90 jours)

```powershell
.\Audit-M365-E3-F3-Usage.ps1
```

### Paramètres disponibles

| Paramètre | Type | Valeur par défaut | Description |
|---|---|---|---|
| `-Period` | `D7` \| `D30` \| `D90` \| `D180` | `D90` | Période d'analyse des rapports d'usage |
| `-OutputFolder` | `string` | `.\Audit-Usage-Licences-M365` | Dossier de sortie des fichiers générés |
| `-CsvDelimiter` | `string` | `;` | Séparateur utilisé dans les fichiers CSV |
| `-MicrosoftLicensingReferenceCsvUrl` | `string` | URL Microsoft officielle | URL du CSV de référence des licences Microsoft |
| `-SkipSkuReferenceDownload` | `switch` | — | Ignore le téléchargement de la référence Microsoft et utilise le mapping local |

### Exemples

```powershell
# Analyse sur 30 jours
.\Audit-M365-E3-F3-Usage.ps1 -Period D30

# Dossier de sortie personnalisé
.\Audit-M365-E3-F3-Usage.ps1 -OutputFolder "C:\Audits\M365-2025"

# Analyse 180 jours avec virgule comme séparateur CSV
.\Audit-M365-E3-F3-Usage.ps1 -Period D180 -CsvDelimiter ","

# Sans téléchargement de la référence Microsoft
.\Audit-M365-E3-F3-Usage.ps1 -SkipSkuReferenceDownload
```

---

## Fichiers générés

Le script crée la structure suivante dans le dossier de sortie :

```
Audit-Usage-Licences-M365/
├── Usage_Licences_M365_E3_F3_E5_Detail_<Period>.csv      # Détail par utilisateur
├── Usage_Licences_M365_E3_F3_E5_Synthese_<Period>.csv    # Synthèse par licence
├── Usage_Licences_M365_E3_F3_E5_SKU_Audites.csv          # SKUs ciblés et résolus
├── Rapports-Bruts/                                        # Rapports bruts Graph API (CSV)
└── Reference-Microsoft/                                   # Référence Microsoft des licences
```

### Fichier de détail (`Detail_*.csv`)

Un enregistrement par utilisateur avec les colonnes suivantes :

**Identité & Licence**
- Nom complet, UPN, Compte actif, Type utilisateur
- Licence, Famille de licence (E3/F3/E5), SKU, GUID SKU, Référence SKU
- Teams inclus dans le bundle

**Usage global**
- Usage détecté (Oui/Non)
- Usage cœur Microsoft 365 (Exchange, OneDrive, SharePoint, Apps)
- Dernière activité détectée
- Score d'activité total, cœur M365 et Teams

**Exchange**
- Usage Exchange, Dernière activité Exchange
- Mails envoyés / reçus / lus

**OneDrive**
- Usage OneDrive, Dernière activité OneDrive
- Fichiers consultés/modifiés, synchronisés, partagés (interne/externe)

**SharePoint**
- Usage SharePoint, Dernière activité SharePoint
- Fichiers consultés/modifiés, pages visitées, partages (interne/externe)

**Teams**
- Usage Teams, Dernière activité Teams
- Messages canal, messages privés, appels, réunions

**Microsoft 365 Apps**
- Usage Apps, Dernière activation, Dernière activité
- Plateformes utilisées : Windows, Mac, Mobile, Web
- Applications utilisées : Outlook, Word, Excel, PowerPoint, Teams

**Recommandation**
- Aucun usage détecté → candidat retrait ou réaffectation, à valider métier
- Usage faible → candidat optimisation ou changement de licence, à valider métier
- Usage Teams uniquement (bundle sans Teams) → non compté comme usage du bundle
- Usage détecté → conserver ou analyser plus finement

### Fichier de synthèse (`Synthese_*.csv`)

Un enregistrement par type de licence avec les compteurs agrégés :
- Nombre total d'utilisateurs
- Utilisateurs avec usage détecté / sans usage
- Compteurs par service (Exchange, OneDrive, SharePoint, Teams, Apps)
- Candidats usage faible / sans usage / Teams seul non compté

---

## Licences auditées

Le script couvre toutes les variantes Microsoft 365 E3, F3 et E5, y compris :

| Produit | Famille | Teams inclus |
|---|---|---|
| Microsoft 365 E3 | E3 | Oui |
| Microsoft 365 E3 EEA (no Teams) | E3 | Non |
| Microsoft 365 E3 EEA (no Teams) - HUB 500 seats minimum | E3 | Non |
| Microsoft 365 E3 - HUB 500 seats minimum | E3 | Oui |
| Microsoft 365 F3 | F3 | Oui |
| Microsoft 365 F3 EEA (no Teams) | F3 | Non |
| Microsoft 365 E5 | E5 | Oui |
| Microsoft 365 E5 - HUB 500 seats minimum | E5 | Oui |
| Microsoft 365 E5 EEA (no Teams) | E5 | Non |
| Microsoft 365 E5 EEA (no Teams) - HUB 500 seats minimum | E5 | Non |
| Microsoft 365 E5 EEA (no Teams) without Audio Conferencing | E5 | Non |
| Microsoft 365 E5 EEA (no Teams) without Audio Conferencing - HUB | E5 | Non |
| Microsoft 365 E5 EEA (no Teams) with Calling Minutes | E5 | Non |

La résolution des SKUs est effectuée en priorité via le [CSV de référence officiel Microsoft](https://learn.microsoft.com/fr-fr/entra/identity/users/licensing-service-plan-reference), avec un fallback sur un mapping local intégré.

---

## Comportement Teams pour les bundles sans Teams

Pour les variantes **EEA (no Teams)**, l'usage Teams est visible dans le fichier de détail mais **n'est pas comptabilisé** comme usage du bundle Microsoft 365. Cela évite de considérer une licence sans Teams comme "utilisée" uniquement parce que l'utilisateur dispose de Teams via une autre licence.

---

## Référence Microsoft

- [Noms de produits et identificateurs de plan de service pour les licences](https://learn.microsoft.com/fr-fr/entra/identity/users/licensing-service-plan-reference)
- [API Microsoft Graph Reports](https://learn.microsoft.com/fr-fr/graph/api/resources/report)
