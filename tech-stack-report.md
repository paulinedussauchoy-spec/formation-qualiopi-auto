# Rapport : Choix de tech stack — Automatisation convention formation Koban

**Projet** : Outil de génération automatique de documents Qualiopi
**Date** : 01/03/2026
**Contexte** : MVP usage personnel (Estelle, formatrice eBazten) → V1 multi-clients

---

## 1. Les 5 options évaluées

| # | Option | En une ligne |
|---|---|---|
| A | **Python + Streamlit** | Backend Python + UI web auto-générée |
| B | **Next.js full-stack** | TypeScript côté serveur et client |
| C | **FastAPI + Next.js** | API Python + frontend React découplés |
| D | **Python CLI** | Script en ligne de commande |
| E | **Electron** | Application desktop installable |

---

## 2. Grille de scores

| Critère | A · Streamlit | B · Next.js | C · FastAPI+Next | D · CLI | E · Electron |
|---|:---:|:---:|:---:|:---:|:---:|
| **Génération Word/PDF** | 5 | 2 | 5 | 5 | 3 |
| **Parsing Excel complexe** | 5 | 3 | 5 | 5 | 3 |
| **UI pour Estelle** | 3 | 4 | 5 | 1 | 5 |
| **Rapidité de dev MVP** | 5 | 2 | 2 | 5 | 3 |
| **Déploiement** | 4 | 4 | 2 | 2 | 4 |
| **Coût hébergement** | 5 | 4 | 3 | 5 | 5 |
| **Évolutivité V1+** | 3 | 5 | 5 | 1 | 2 |
| **Moyenne** | **4.3** | **3.4** | **3.9** | **3.4** | **3.6** |

---

## 3. Analyse détaillée par option

### A — Python + Streamlit

**Points forts :**
- `docxtpl` (templating Jinja2 sur vrais fichiers .docx) = Estelle garde ses templates Word, le code injecte les données — aucune autre option ne fait ça aussi simplement
- `pandas` + `openpyxl` = parsing de tableaux croisés complexes en quelques lignes
- MVP fonctionnel en 2-3 jours de développement
- Streamlit Community Cloud : hébergement gratuit depuis GitHub, accessible depuis n'importe quel navigateur, sans installation côté Estelle
- `libreoffice --headless` ou `docx2pdf` pour la conversion Word → PDF incluse

**Points faibles :**
- L'UI est contrainte : Streamlit re-exécute tout le script à chaque interaction, ce qui limite les workflows complexes (confirmation avant génération, aperçu interactif)
- Pas d'authentification native — acceptable en usage solo, problématique si multi-utilisateurs
- Architecture "script" qui se reexécute : difficile d'ajouter un état persistant (historique, tableau de bord) sans se battre contre le framework

**Bibliothèques clés :**
```
docxtpl        # templates Word avec Jinja2
openpyxl       # parsing Excel brut (cellules fusionnées, styles)
pandas         # transformation des données tabulaires
streamlit      # interface web
python-docx    # manipulation avancée .docx si besoin
```

---

### B — Next.js full-stack (TypeScript)

**Points forts :**
- Pauline connaît déjà la stack (Next.js 14 + shadcn/ui + TypeScript — même que cc-sales-tunnel)
- Architecture naturellement extensible : NextAuth, Prisma, API Routes, tout est déjà documenté
- Vercel : déploiement en quelques clics, CI/CD automatique
- Meilleure UI possible avec React + shadcn/ui

**Points faibles — structurels :**
- **Génération Word** : `docx-templates` (Node.js) est moins mature que `docxtpl`. Les styles hérités depuis un vrai fichier .docx sont difficiles à préserver. Risque de régression sur la mise en page des documents Qualiopi
- **Excel complexe** : SheetJS (gratuit) couvre les cas basiques, mais la logique de tableau croisé (stagiaires × modules, cellules fusionnées, colonnes à mapper) nécessite du code manuel que pandas gère nativement
- **Timeout Vercel** : la génération de 60 documents en batch peut dépasser 10 s sur le plan gratuit. Contournement possible (background jobs) mais ajoute de la complexité
- **Délai MVP** : 1 à 2 semaines minimum pour un résultat comparable à Streamlit en 2-3 jours

---

### C — FastAPI (Python) + Next.js frontend

**Points forts :**
- Combine le meilleur des deux mondes : Python pour la logique documentaire, React pour l'UI
- Architecture maximalement extensible : chaque service évolue indépendamment
- Quand Koban, email, et signature électronique arrivent, ils s'ajoutent proprement comme routes FastAPI ou nouveaux microservices

**Points faibles — critiques pour le MVP :**
- **Deux projets à déployer** : FastAPI (Railway/Render/VPS) + Next.js (Vercel) = deux pipelines, deux configs, deux logs, deux domaines
- **Render tier gratuit** : cold start de 30-60 secondes après 15 minutes d'inactivité — irritant pour une utilisatrice
- **CORS, types partagés, contrats d'API** : chaque couche d'interface ajoute de la friction en développement
- **Délai MVP** : 2 à 3 semaines pour un résultat propre

> Cette option est la meilleure à long terme, mais la plus risquée pour un MVP rapide.

---

### D — CLI Python

**Disqualifié.** L'UI (terminal) est inutilisable pour Estelle sans accompagnement technique. Et l'architecture est un cul-de-sac pour la V1.

---

### E — Electron

**Points forts :**
- Application desktop installée = zero friction pour Estelle (comme Word ou Excel)
- Accès direct au système de fichiers, pas de serveur

**Points faibles :**
- Bibliothèques Node.js pour Word/PDF moins performantes que Python
- Multi-clients et tableau de bord = réécriture quasi-complète en V1
- Code signing (Apple Developer Program) = coût annuel si distribution formelle
- Taille du bundle : une app Electron embarque Chromium (~150 MB)

---

## 4. Les tensions clés du choix

**Tension 1 — Python (documents) vs. JavaScript (UI)**
Python domine sur Word/Excel. JavaScript domine sur les UI web modernes.
→ Aucune option ne les combine sans coût (FastAPI + Next.js le fait, mais avec complexité).

**Tension 2 — Rapidité MVP vs. Évolutivité V1**
Streamlit et CLI sont rapides mais plafonnent vite.
Next.js et FastAPI+Next.js sont extensibles mais lents à démarrer.

**Tension 3 — Local vs. Cloud**
Estelle veut accéder à l'outil depuis n'importe où ? → Cloud obligatoire.
Estelle travaille toujours depuis le même poste ? → Local possible.

**Tension 4 — Valider la logique métier d'abord vs. Construire l'architecture finale**
La vraie difficulté de ce projet n'est pas l'UI : c'est le parsing du tableau Excel GENERFEU avec ses colonnes ambiguës, les TNS à exclure, et le mapping module. Cette logique doit être validée avec Estelle avant de construire une UI sophistiquée.

---

## 5. Recommandation

### ✅ Approche en deux phases : Streamlit (MVP) → FastAPI + Next.js (V1)

#### Phase 1 — MVP : Python + Streamlit

**Livrable** : outil fonctionnel utilisable par Estelle en 2-3 semaines de développement.

**Pourquoi :**

1. **Le risque principal est documentaire, pas UI** : le parsing du tableau Excel GENERFEU (colonnes ambiguës, TNS, mapping) et la génération fidèle des conventions Word sont les défis réels. Python résout ces deux problèmes mieux que toute autre option. Il faut valider ça en premier.

2. **Streamlit est suffisant pour un usage solo** : Estelle accède à une URL, uploade son Excel, configure les paramètres clients, clique "Générer" et télécharge un ZIP contenant ses 60 documents. C'est tout ce dont elle a besoin en MVP.

3. **Le code Python est 100% réutilisable en V1** : la logique de parsing et génération écrite en Phase 1 sera simplement encapsulée dans des routes FastAPI en Phase 2. Rien n't est jeté.

4. **Streamlit Community Cloud = gratuit et zéro serveur à gérer** : déploiement en 10 minutes depuis GitHub.

**Stack Phase 1 :**
```
Python 3.11+
├── streamlit          # interface web
├── docxtpl            # génération Word depuis templates
├── pandas             # parsing Excel
├── openpyxl           # accès bas niveau Excel
└── python-docx        # manipulation .docx avancée
```

**Architecture Phase 1 :**
```
app.py                     # interface Streamlit
├── core/
│   ├── excel_parser.py    # lecture et transformation du tableau Excel
│   ├── document_gen.py    # génération Word depuis templates
│   └── mapping.py         # table de correspondance modules Koban
├── templates/
│   ├── convention.docx    # template Word convention (vrai fichier)
│   ├── certificat.docx    # template Word certificat
│   └── feuille_presence.docx
└── config/
    └── modules.json       # référentiel des 16 modules + objectifs
```

---

#### Phase 2 — V1 : FastAPI + Next.js

**Déclencheur** : quand l'une de ces conditions est remplie :
- Estelle veut gérer plusieurs clients avec un tableau de bord
- D'autres utilisateurs doivent accéder à l'outil
- L'intégration Koban CRM ou l'envoi email est prioritaire
- L'UI Streamlit est perçue comme trop limitante

**Migration :**
- `core/` (parser Excel, génération Word) → migre tel quel dans FastAPI, sans modification
- Streamlit UI → remplacée par Next.js + shadcn/ui (stack connue de Pauline via cc-sales-tunnel)
- Ajout : Prisma + SQLite/PostgreSQL pour la base clients, NextAuth pour l'auth

**Stack Phase 2 :**
```
Backend : FastAPI + python-docx + pandas + docxtpl
Frontend : Next.js 14 + TypeScript + shadcn/ui + Prisma
Déploiement : Railway (FastAPI) + Vercel (Next.js)
```

---

## 6. Ce que cette approche évite

| Risque évité | Pourquoi |
|---|---|
| Passer 3 semaines sur l'UI avant de savoir si le parsing Excel fonctionne | Phase 1 valide la logique métier en premier |
| Réécrire la génération Word en JS pour découvrir que docx-templates ne gère pas les styles | Python + docxtpl dès le début |
| Livrer un outil inutilisable pour Estelle faute d'UI | Streamlit donne une interface opérationnelle rapidement |
| Payer un serveur Railway dès le MVP pour un usage solo | Streamlit Cloud = gratuit |
| Jeter le code Python quand on passe à la V1 | La logique métier migre directement dans FastAPI |

---

## 7. Décision finale

```
MVP  →  Python + Streamlit   (2-3 semaines, gratuit, validations logique métier)
V1   →  FastAPI + Next.js    (si multi-clients, tableau de bord, intégrations)
```

Le code Python de la Phase 1 est l'investissement — il ne sera pas jeté.
L'UI Streamlit est temporaire — elle sera remplacée si nécessaire.
La vraie valeur de ce projet est dans la logique de parsing Excel et de génération documentaire, pas dans le framework.
