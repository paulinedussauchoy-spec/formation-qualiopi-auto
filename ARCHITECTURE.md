# Architecture technique — Automatisation Qualiopi Koban

**Projet** : Générateur automatique de documents de formation (conventions, certificats, feuilles de présence)
**Utilisatrice** : Estelle Lecanu (eBazten), PC Windows, non-technicienne
**Date** : 01/03/2026

---

## Décision d'architecture : approche en deux phases

La logique métier (parsing Excel, groupement des conventions, génération Word) est le risque principal
du projet — pas l'interface. On valide d'abord cette logique avec une UI minimale, puis on construit
une interface plus riche si les besoins évoluent.

```
Phase 1 (MVP)  →  Python + Streamlit Community Cloud
Phase 2 (V1)   →  FastAPI (Python) + Next.js  [si multi-clients / tableau de bord / API Koban]
```

Le code Python du `core/` (Phase 1) migre directement dans FastAPI sans modification en Phase 2.

---

## Phase 1 — MVP

### Stack

| Couche | Choix | Justification |
|---|---|---|
| Langage | Python 3.11+ | Seul écosystème avec des bibliothèques matures pour Word, Excel et PDF |
| Interface web | Streamlit | UI générée en Python pur, zéro HTML/CSS, déployable en 10 min |
| Hébergement | Streamlit Community Cloud | Gratuit, accessible via URL depuis n'importe quel navigateur, aucune installation côté Estelle |
| Parsing Excel | openpyxl (direct) + pandas | openpyxl pour les cellules brutes (lignes 1-7 de métadonnées), pandas pour la manipulation des données stagiaires |
| Génération Word | docxtpl | Moteur Jinja2 directement dans les fichiers .docx — Estelle garde ses vrais templates Word |
| Export PDF | LibreOffice headless | Serveurs Streamlit Cloud = Linux, LibreOffice headless natif. Fiable et gratuit. |
| Configuration | JSON | Référentiel des 16 modules + objectifs pédagogiques + données fixes eBazten |

### Pourquoi openpyxl direct (et pas pandas seul)

Le tableau Excel d'Estelle a 7 lignes de métadonnées avant les données stagiaires, avec des cellules
fusionnées sur plusieurs colonnes. `pandas.read_excel()` ne gère pas correctement les cellules fusionnées :
les cellules non-top-left retournent `None`. Il faut passer par `ws.iter_rows()` d'openpyxl pour
reconstruire la structure proprement, puis transmettre à pandas pour la manipulation.

### Pourquoi LibreOffice headless (et pas docx2pdf)

`docx2pdf` est abandonné depuis 2020 et requiert Microsoft Word installé sur la machine (AppleScript
sur Mac, COM sur Windows). Sur Streamlit Cloud (Linux), il ne fonctionne pas du tout. LibreOffice
headless est la solution standard côté serveur, gratuite et fiable.

### Sources de données disponibles

Les fichiers de référence sont dans `base de travail/` :

| Fichier | Rôle |
|---|---|
| `GENERFEU - Formations et inscrits.xlsx` | Fichier Excel principal (modèle de test) |
| `Liste des formations et des objectifs.xlsx` | Référentiel des 16 modules + objectifs |
| `ok - GENERFEU - Utilisateurs et Prérequis.xlsx` | Données stagiaires (prérequis, coordonnées) |
| `A signer - GENERFEU - CONVENTION *.docx` (×5) | Conventions réelles GENERFEU → servent de base pour le template convention |
| `[CLIENT] - CERTIF REALISATION.docx` | Template certificat de réalisation |
| `[CLIENT] FEUILLE DE PRESENCE.docx` | Template feuille de présence |

> L'onglet à lire dans le fichier Excel est nommé `"Formations et inscrits"` (nom standard d'Estelle).
> Fallback : premier onglet si le nom diffère pour un autre client.

### Format de sortie

Chaque génération produit **les deux formats** dans un ZIP :
- `.docx` — pour qu'Estelle puisse corriger manuellement si besoin
- `.pdf` — pour envoi direct au client

```
GENERFEU_2026-03-01.zip
├── conventions/
│   ├── CONVENTION_Administrateurs.docx
│   ├── CONVENTION_Administrateurs.pdf
│   ├── CONVENTION_Commerciaux_BE_Gr01.docx
│   └── ...
├── certificats/
│   ├── CERTIF_DUPONT_Christelle.docx
│   ├── CERTIF_DUPONT_Christelle.pdf
│   └── ...
└── feuilles_presence/
    ├── PRESENCE_1-01_05-04-2026.docx
    ├── PRESENCE_1-01_05-04-2026.pdf
    └── ...
```

### Structure du projet

```
koban-qualiopi/
├── app.py                          # Interface Streamlit (upload, config, génération, download)
│
├── core/
│   ├── excel_parser.py             # Lecture openpyxl : reconstruit les 7 lignes de métadonnées
│   ├── group_builder.py            # Logique : quels stagiaires → quelle convention
│   │                               # (regroupement par ensemble de modules cochés, pas par en-tête)
│   ├── module_mapper.py            # Table de correspondance colonnes Excel → codes modules Koban
│   └── document_gen.py             # Génération .docx via docxtpl + conversion PDF LibreOffice
│
├── templates/                      # Fichiers .docx avec tags Jinja2 (basés sur les vrais docs d'Estelle)
│   ├── convention.docx
│   ├── certificat.docx
│   └── feuille_presence.docx
│
├── config/
│   └── modules.json                # 16 modules Koban : code, intitulé, objectifs pédagogiques
│
├── requirements.txt
└── .streamlit/
    └── config.toml                 # Config Streamlit (thème, upload max size)
```

### Dépendances (requirements.txt)

```
streamlit>=1.32
openpyxl>=3.1
pandas>=2.2
docxtpl>=0.16
python-docx>=1.1
```

> LibreOffice est disponible nativement sur Streamlit Cloud — aucun package Python à installer pour la conversion PDF.

### Workflow en deux phases

La génération se fait en **deux moments distincts** correspondant au cycle de vie d'une formation :

**Phase A — Avant la formation (pré-formation)**
→ Génération des **Conventions** uniquement.
Les dates sont une fourchette large saisie manuellement ("du 15/03/2026 au 15/11/2026").
L'Excel peut ne pas encore avoir les dates/horaires définitifs (lignes 4-5 vides).

**Phase B — Après la formation (post-formation)**
→ Génération des **Certificats de réalisation** + **Feuilles de présence**.
Les vraies dates et horaires par session sont nécessaires.
Source : lignes 4 (Date) et 5 (Horaires) de l'Excel — Estelle les remplit après planification avec le client.
Si l'Excel n'a pas encore ces données, Estelle les saisit manuellement dans l'interface.

### Flux utilisateur (Estelle)

**Avant la formation :**
```
1. Ouvre l'URL Streamlit dans son navigateur (PC Windows)
2. Upload le fichier Excel "Formations et inscrits"
3. Saisit les infos variables du client (nom société, adresse, représentant légal + fonction)
4. Saisit la fourchette de dates prévisionnelles et les frais de mission (souvent 0)
5. Aperçu des conventions détectées (groupes de stagiaires, modules, montants calculés)
6. Valide ou ajuste la liste des stagiaires (coche/décoche les "optionnel")
7. Clique "Générer les conventions"
8. Télécharge le ZIP (conventions .docx + .pdf)
```

**Après la formation :**
```
1. Re-upload le même Excel (maintenant complété avec les dates/horaires réels, lignes 4-5)
2. Clique "Générer les certificats et feuilles de présence"
3. Télécharge le ZIP (certificats .docx/.pdf + feuilles de présence .docx/.pdf)
```

### Logique de groupement des conventions

Le groupement **ne se fait pas sur les en-têtes de colonnes** (ligne 1 du tableau) mais sur
**l'ensemble des stagiaires cochés** pour un même groupe de modules. Deux stagiaires avec
exactement les mêmes colonnes cochées → même convention.

Cas particuliers gérés :
- **TNS** (`Gilles BERTRAND` = `TNS`) → participe aux formations mais n'apparaît pas dans la convention
- **"optionnel" / "à inviter"** → inclus par défaut (inclusion maximale). Estelle peut les décocher manuellement dans l'interface avant de générer.

### Calculs automatiques

| Calcul | Formule |
|---|---|
| Nombre de journées | `demi-journées / 2` |
| Nombre d'heures par stagiaire | `journées × 7` |
| Montant HT | `demi-journées × 450€` |
| Total TTC | `Montant HT × 1.20` |

### Données fixes (identiques sur tous les documents)

Stockées dans `config/modules.json` — ne nécessitent pas de saisie à chaque génération :

- Organisme : eBazten, 2495 route de Forcalquier, 04300 Pierrerue
- Représentant OF : Guillaume Lecanu, Gérant
- N° déclaration : 82691206769 — DataDock : 0029644 — SIREN : 539304683
- Formatrice : Estelle Lecanu
- Questionnaires (liens Google Forms à chaud et à froid)
- Tarif : 450€ HT / demi-journée — TVA : 20% — Dédommagement annulation : 400€ HT

---

## Phase 2 — V1 (conditions de déclenchement)

La migration vers cette stack se fait **seulement si** au moins une de ces conditions est remplie :
- Estelle gère plusieurs clients et a besoin d'un tableau de bord
- D'autres personnes doivent accéder à l'outil
- L'intégration avec l'API Koban CRM devient prioritaire
- L'UI Streamlit est perçue comme trop limitante

### Stack Phase 2

| Couche | Choix |
|---|---|
| Backend API | FastAPI (Python) |
| Frontend | Next.js 14 + TypeScript + shadcn/ui (même stack que cc-sales-tunnel) |
| Base de données | SQLite → PostgreSQL (via SQLAlchemy + Prisma) |
| Auth | NextAuth |
| Déploiement backend | Railway |
| Déploiement frontend | Vercel |

> Le `core/` Python (excel_parser, group_builder, module_mapper, document_gen) migre directement
> dans les routes FastAPI sans modification. Seule l'interface change.

---

## Note sur la tarification

**Formule confirmée : `Montant HT = nombre_demi_journées_facturées × 450€`**

Le nombre de demi-journées facturées **n'est pas automatiquement déduit du tableau Excel**.
Estelle peut décider de ne facturer qu'une partie des modules prévus (ex : Convention Administrateurs
GENERFEU = 4 modules dans l'Excel, mais seulement 1 journée = 2 demi-journées commandées et facturées).

→ L'interface proposera le nombre de demi-journées calculé depuis l'Excel comme valeur par défaut,
mais Estelle pourra le modifier avant de générer la convention.

## Points ouverts

Aucun point ouvert. L'architecture est complète et validée.
