# Projet : Automatisation de la paperasse formation Qualiopi

## Contexte

**Qui** : Estelle Lecanu, formatrice chez eBazten (éditeur du CRM Koban), organisme certifié Qualiopi.

**Problème** : Pour chaque client, Estelle doit produire manuellement (copier-coller dans Word) des dizaines de documents imposés par la certification Qualiopi. Sur un client comme GENERFEU (39 stagiaires, 5 groupes, ~15 sessions), cela représente environ **60 documents**. Elle y passe **4 à 7 heures par mois**.

**Enjeu ROI** : Un outil spécialisé type Digiforma coûte 200-350€/mois. L'objectif est de voir si une automatisation sur mesure permet de faire mieux, pour moins cher.

---

## 1. Qu'est-ce qu'on souhaite construire ?

Un **outil de génération automatique de documents de formation Qualiopi** qui, à partir des données Excel existantes, produit en lot les 3 types de documents obligatoires.

### Fonctionnalités

- Génération automatique des **Conventions de formation** (Word/PDF) — 1 par groupe de stagiaires
- Génération automatique des **Certificats de réalisation** (Word/PDF) — 1 par stagiaire
- Génération automatique des **Feuilles de présence** (Word/PDF) — 1 par demi-journée de formation
- Calculs automatiques (demi-journées → journées → montant HT/TTC)
- Injection automatique des **objectifs pédagogiques** par module (depuis la liste fixe)
- Injection automatique des **noms/prénoms** des stagiaires depuis le tableau de répartition
- Interface web simple pour configurer et lancer la génération

---

## 2. Objectif du projet

**MVP pour utilisation personnelle** (Estelle), avec potentiel de réutilisation pour tous ses clients (pas seulement GENERFEU).

---

## 3. Fonctionnalités par étape

### MVP — Ce qui DOIT fonctionner

- Lire le tableau Excel "Formations et inscrits" (groupes × modules × stagiaires)
- Lire la liste des objectifs par module
- Appliquer la table de mapping (voir section 6) pour traduire les en-têtes Excel en codes modules
- Générer les **Conventions** Word pré-remplies à partir du template existant
- Générer les **Certificats de réalisation** individuels
- Générer les **Feuilles de présence** par demi-journée
- Export en lot (tous les docs d'un client d'un coup)
- Interface minimale (formulaire web simple)

### V1 — Fonctionnalités additionnelles importantes

- Interface web propre pour saisir/modifier les infos client, dates, modalités
- Gestion multi-clients (pas juste GENERFEU)
- Export PDF automatique (pas juste Word)
- Tableau de bord de suivi : quels docs sont générés, envoyés, signés
- Pré-remplissage des infos client depuis un annuaire/base

### + Tard — Bonus, pas prioritaire

- Intégration avec Koban CRM (récupérer les contacts directement)
- Envoi email automatique des documents
- Signature électronique intégrée
- Gestion des questionnaires à chaud/à froid
- Suivi des indicateurs Qualiopi

### Hors périmètre

- Remplacer complètement Digiforma (audit Qualiopi, BPF, etc.)
- Facturation / comptabilité
- Planning / calendrier des formations

---

## 4. Les 3 documents à générer

### 4.1 Convention de formation

**Fréquence** : 1 par groupe de stagiaires ayant les mêmes modules
**Template** : Word avec mentions légales fixes
**Contrainte Qualiopi** : toutes les personnes listées dans une convention doivent suivre TOUS les modules de cette convention (pas de mix)

**Champs variables :**

| Champ | Source | Exemple |
|---|---|---|
| Nom de l'entreprise cliente | Saisie manuelle ou base client | GENERFEU |
| Adresse de l'entreprise | Saisie manuelle ou base client | PARC D'ACTIVITES DES ECLAPONS... |
| Représentant + fonction | Saisie manuelle ou base client | Gilles BERTRAND, Dirigeant |
| Liste des modules (codes + intitulés) | Table de mapping + Excel | 1-01 - Socle commun - Administration socle commun |
| Effectif formé | Comptage auto depuis Excel | 5 |
| Nombre de demi-journées | Comptage auto des modules | 4 |
| Nombre de journées | Calcul : demi-journées / 2 | 2 |
| Nombre d'heures par stagiaire | Calcul : journées × 7 | 14 heures |
| Modalité | Depuis Excel (ligne 6) | A distance / Sur site |
| Dates prévisionnelles | Saisie manuelle | Du 01/04/2026 au 31/10/2026 |
| Noms et prénoms des stagiaires | Depuis Excel (colonnes B/C) | DUPONT Christelle, FAVRE Audrey... |
| Montant HT | Calcul : demi-journées × 450€ | 900.00€ HT |
| Frais de mission | Saisie manuelle (souvent 0) | 0.00€ HT |
| Total HT | Calcul | 900.00€ HT |
| Total TTC | Calcul : HT × 1.20 | 1080.00€ TTC |
| Objectifs par module | Depuis la liste des objectifs | (texte long, voir section 7) |
| Date du document | Saisie ou date du jour | 26/02/2026 |

**Champs fixes (identiques pour toutes les conventions) :**

- Organisme de formation : eBazten, 2495 route de Forcalquier, 04300 Pierrerue
- Représentant OF : Guillaume Lecanu, Gérant
- N° déclaration d'activité : 82691206769
- Identifiant DataDock : 0029644
- SIREN : 539304683
- Formatrice : Estelle Lecanu
- Sections : Moyens pédagogiques, Validation des acquis, Sanctions, Moyens de suivi, Non-réalisation, Dédommagement, Litiges
- Liens questionnaires à chaud et à froid (Google Forms)
- Signatures et cachets

### 4.2 Certificat de réalisation

**Fréquence** : 1 par stagiaire (individuel)
**Template** : Word court (1 page)

**Champs variables :**

| Champ | Source | Exemple |
|---|---|---|
| NOM du stagiaire | Depuis Excel | DUPONT |
| Prénom du stagiaire | Depuis Excel | Christelle |
| Nom et adresse de la société | Base client | GENERFEU... |
| Nom de la formation (modules suivis) | Convention/mapping | 1-01 - Socle commun... |
| Nombre de journées | Depuis convention | 2 |
| Nombre d'heures par stagiaire | Depuis convention | 14 heures |
| Modalités / Lieu | Depuis Excel | A distance |
| Dates et horaires réels de session | Saisie manuelle (post-formation) | 05/04/2026 9h-12h30... |
| Date du document | Saisie ou date du jour | |

**Champs fixes :**

- Signataire : Estelle Lecanu, SARL Ebazten
- Lieu : Pierrerue
- Signatures/cachets

### 4.3 Feuille de présence

**Fréquence** : 1 par demi-journée de formation (par session)
**Template** : Word avec tableau de signatures

**Champs variables :**

| Champ | Source | Exemple |
|---|---|---|
| Nom et adresse de la société | Base client | GENERFEU... |
| Nom de la formation (1 module) | Mapping | 1-01 - Administration socle commun |
| Nombre de journées | Fixe : 0.5 | 0.5 |
| Nombre d'heures | Fixe : 3.5 | 3.5 heures |
| Modalités | Depuis Excel | VISIO |
| Date et horaire de la session | Saisie manuelle | 05/04/2026, 9h-12h30 |
| Tableau des stagiaires (NOM + Prénom) | Depuis Excel (le groupe) | DUPONT Christelle, FAVRE Audrey... |
| Date du document | Saisie | |

**Champs fixes :**

- Formatrice : Estelle Lecanu
- Signatures formatrice (pré-insérées)

---

## 5. Les sources de données

### 5.1 Tableau "Formations et inscrits" (Excel)

C'est le **fichier maître**. Pour chaque client, un tableau croisé :

**Structure :**

- **Ligne 1** : Type de groupe (ex: "Administrateurs", "Utilisateurs")
- **Ligne 2** : Domaine du module (ex: "Socle commun", "CRM", "SAV", "Facturation clients / CONTRATS")
- **Ligne 3** : Durée (toujours 0.5 = demi-journée)
- **Ligne 4** : Date (à remplir après planification)
- **Ligne 5** : Horaires (à remplir après planification)
- **Ligne 6** : Modalités ("VISIO" ou "SUR SITE")
- **Ligne 7** : Capacité min/max
- **Lignes 9+** : Stagiaires (col 1 = N°, col 2 = NOM, col 3 = Prénom, col 4 = Email, puis "x" ou "TNS" dans les colonnes modules)

**Cas GENERFEU — Colonnes modules :**

| Colonnes | Type (L1) | Module (L2) | Modalité | Groupe de convention |
|---|---|---|---|---|
| Col 5-8 | Administrateurs | Socle commun, CRM, SAV, Facturation | VISIO | Convention Administrateurs |
| Col 9-10 | Utilisateurs | Socle commun, CRM | SUR SITE | Convention Commerciaux BE Gr01 |
| Col 11-12 | Utilisateurs | Socle commun, CRM | SUR SITE | Convention Commerciaux BE Gr01 (bis) |
| Col 13-14 | Utilisateurs | Socle commun, CRM | SUR SITE | Convention Commerciaux BE Gr02 |
| Col 15 | Utilisateurs | Socle commun | SUR SITE | Convention SAV Contrats Gr03 |
| Col 16 | Utilisateurs | SAV +++ | SUR SITE | Convention SAV Contrats Gr03 |
| Col 17 | Utilisateurs | Facturation / CONTRATS | SUR SITE | Convention SAV Contrats Gr03 |
| Col 18-19 | Utilisateurs | SAV, SAV | SUR SITE | Convention Techniciens Gr04 |

### 5.2 Liste des formations et objectifs (Excel)

Fichier référentiel fixe contenant les **16 modules** Koban avec :
- Le code (ex: "1-01")
- L'intitulé complet (ex: "Socle commun - Administration socle commun")
- Les objectifs pédagogiques (texte long)

### 5.3 Tableau "Utilisateurs et Prérequis" (Excel)

Fiche détaillée par stagiaire :
- NOM, Prénom, Email, Mobile, Ligne directe
- Fonction dans la société
- Modules utilisés dans Koban
- Ancienneté, connaissance des process internes
- Maîtrise bureautique, navigation web
- Expérience CRM antérieure
- Type de poste (PC/Mac)
- Disponibilité télétravail + connexion wifi
- Situation de handicap

→ Utile principalement pour les **prérequis** mentionnés dans les conventions et pour avoir les coordonnées complètes.

---

## 6. Table de mapping : En-têtes Excel → Codes modules

C'est la clé de voûte de l'automatisation. La correspondance entre ce qu'Estelle met dans son tableau Excel et les codes modules officiels :

### Règle générale

| Type (Ligne 1) | Domaine (Ligne 2) | Code module | Intitulé officiel |
|---|---|---|---|
| Administrateurs | Socle commun | **1-01** | Socle commun - Administration socle commun |
| Administrateurs | CRM | **1-03** | CRM - Administration CRM |
| Administrateurs | SAV | **1-09** | Service client SAV - Administration service client SAV |
| Administrateurs | Facturation clients / CONTRATS | **1-05** | Facturation - Administration facturation client - Niveau 1 |
| Utilisateurs | Socle commun | **1-02** | Socle commun - Utilisation socle commun |
| Utilisateurs | CRM | **1-0A** | CRM - Utilisation CRM |
| Utilisateurs | SAV / SAV +++ | **1-10** | Service client SAV - Utilisation service client SAV |
| Utilisateurs | Facturation clients / CONTRATS | **2-06** | Facturation - Utilisation facturation client - Niveau 2 - Contrats |

### Modules additionnels disponibles (catalogue complet)

Ces modules n'apparaissent pas dans le cas GENERFEU mais pourraient être utilisés pour d'autres clients :

| Code | Intitulé |
|---|---|
| 1-06 | Facturation - Utilisation facturation client - Niveau 1 |
| 1-07 | Marketing - Marketing niveau 1 - Emailing |
| 2-05 | Facturation - Administration facturation client - Niveau 2 - Contrats |
| 2-07 | Marketing - Marketing niveau 2 - Tracking et points engagements |
| 2-08 | Marketing - Marketing niveau 2 - Scenarii marketing |
| 2-09 | Gestion - Gestion niveau 3 - Achats - Paramétrage et utilisation |
| 2-10 | Gestion - Gestion niveau 3 - Stocks - Paramétrage et utilisation |
| 2-11 | Gestion - Gestion niveau 3 - Projets - Paramétrage et utilisation |

### Comment le mapping est déduit

La correspondance a été établie en croisant :
1. Les **conventions existantes** (qui contiennent les codes modules ET les noms des stagiaires)
2. Le **tableau Excel** (qui contient les mêmes stagiaires avec des "x" dans les colonnes)

En retrouvant les mêmes personnes dans les deux sources, on déduit quelle colonne Excel correspond à quel code module.

**Logique** : Le type "Administrateurs" pointe toujours vers les modules de type "Administration", et le type "Utilisateurs" vers les modules de type "Utilisation", dans le même domaine fonctionnel.

---

## 7. Objectifs pédagogiques par module (référentiel fixe)

### 1-01 — Socle commun - Administration socle commun
**Objectifs opérationnels** : Être autonome dans les paramétrages quotidiens de Koban

A l'issue du parcours, l'apprenant sera capable de :
- Gérer les utilisateurs de la solution
- Paramétrer les fonctionnalités transverses du logiciel pour assurer la bonne utilisation et compréhension du socle commun au sein de la société
- Guider les utilisateurs dans la résolution de questions basiques
- Guider tout nouvel utilisateur de Koban dans sa découverte de l'outil

### 1-02 — Socle commun - Utilisation socle commun
**Objectifs opérationnels** : Utiliser Koban de façon autonome en accord avec les process de l'entreprise

A l'issue du parcours, l'apprenant sera capable de :
- Gérer et mettre à jour les informations de ses prospects/clients et contacts
- Suivre et planifier son activité commerciale
- Partager les informations avec les autres utilisateurs de Koban

### 1-03 — CRM - Administration CRM
**Objectifs opérationnels** : Être autonome dans les paramétrages CRM de Koban

A l'issue du parcours, l'apprenant sera capable de :
- Paramétrer les processus commerciaux en adéquation avec le process commercial internet et les indicateurs souhaités
- Mettre à jour le catalogue de produits vente
- Personnaliser les devis
- Créer et suivre les indicateurs

### 1-0A — CRM - Utilisation CRM
**Objectifs opérationnels** : Utiliser Koban de façon autonome en accord avec les process de l'entreprise

A l'issue du parcours, l'apprenant sera capable de :
- Comprendre la notion de processus commercial, de pipeline et d'opportunité
- Créer et suivre ses opportunités
- Faire un devis et le suivre jusqu'au résultat

### 1-05 — Facturation - Administration facturation client - Niveau 1
**Objectifs opérationnels** : Être autonome dans les paramétrages de la facturation clients (hors contrats) de Koban

A l'issue du parcours, l'apprenant sera capable de :
- Ajuster le catalogue produits
- Réaliser et ajuster les paramétrages nécessaires au processus de facturation clients
- Créer et suivre les indicateurs

### 1-06 — Facturation - Utilisation facturation client - Niveau 1
**Objectifs opérationnels** : Facturer de façon autonome et gérer les recouvrements clients

A l'issue du parcours, l'apprenant sera capable de :
- Gérer le processus de facturation client depuis la signature du devis
- Saisir les paiements clients et suivre le recouvrement

### 2-05 — Facturation - Administration facturation client - Niveau 2 - Contrats
**Objectifs opérationnels** : Être autonome pour paramétrer le module Contrats et les automatismes liés

A l'issue du parcours, l'apprenant sera capable de :
- Ajuster le catalogue produits en fonction des modalités de ses contrats
- Paramétrer les automatismes du renouvellement
- Paramétrer la facturation automatique

### 2-06 — Facturation - Utilisation facturation client - Niveau 2 - Contrats
**Objectifs opérationnels** : Gérer les contrats client et assurer renouvellement et facturation

A l'issue du parcours, l'apprenant sera capable de :
- Créer un contrat complet depuis une commande client ou en direct
- Comprendre, compléter et ajuster le détail du contrat
- Gérer le renouvellement
- Gérer la facturation

### 1-07 — Marketing - Marketing niveau 1 - Emailing
**Objectifs opérationnels** : Être autonome dans le paramétrage et la gestion d'une campagne emailing

A l'issue du parcours, l'apprenant sera capable de :
- Créer ses modèles emails
- Assurer l'ensemble des paramétrages emailings préalables
- Créer, envoyer et analyser ses campagnes emailings

### 2-07 — Marketing - Marketing niveau 2 - Tracking et points engagements
**Objectifs opérationnels** : Être autonome pour tracker et engager depuis son site web

A l'issue du parcours, l'apprenant sera capable de :
- Tracker les pages à valeur de son site
- Créer et intégrer un point d'engagement standard sur son site
- Paramétrer les points d'engagement en vue de suivre les leads générés

### 2-08 — Marketing - Marketing niveau 2 - Scenarii marketing
**Objectifs opérationnels** : Être autonome pour créer et lancer des scenarii de marketing automation

A l'issue du parcours, l'apprenant sera capable de :
- Créer un scénario en maîtrisant les types d'étapes et les bonnes pratiques
- Lancer un scénario depuis un contact ou une requête
- Analyser le scénario en cours

### 1-09 — Service client SAV - Administration service client SAV
**Objectifs opérationnels** : Être autonome dans le paramétrage du service clients : gestion et personnalisation des tickets

A l'issue du parcours, l'apprenant sera capable de :
- Comprendre et transmettre la manière de gérer un ticket
- Personnaliser le processus de gestion de tickets et les informations constituant un ticket

### 1-10 — Service client SAV - Utilisation service client SAV
**Objectifs opérationnels** : Créer et suivre de façon autonome les demandes clients par le biais des tickets du module

A l'issue du parcours, l'apprenant sera capable de :
- Créer un ticket et remplir les informations demandées en fonction de l'étape et du process défini en interne
- Utiliser les fonctionnalités disponibles pour échanger et partager l'information
- Suivre le ticket et clôturer le ticket

### 2-09 — Gestion - Gestion niveau 3 - Achats - Paramétrage et utilisation
**Objectifs opérationnels** : Être autonome pour créer fournisseurs et produits achats liés, créer une commande fournisseur, lier une facture fournisseur, générer un bon de réception

A l'issue du parcours, l'apprenant sera capable de :
- Créer un statut fournisseur
- Créer un produit achat et le lier à n fournisseurs
- Générer une commande fournisseur, la valider, la réceptionner via un bon de réception
- Solder la commande fournisseur en enregistrant la facture du fournisseur
- Enregistrer le règlement d'une facture fournisseur

### 2-10 — Gestion - Gestion niveau 3 - Stocks - Paramétrage et utilisation
**Objectifs opérationnels** : Être autonome pour gérer les stocks : entrées et sorties, inventaires

A l'issue du parcours, l'apprenant sera capable de :
- Créer ses dépôts
- Paramétrer les produits devant être gérés en stocks
- Créer une entrée en stocks
- Créer une sortie de stocks
- Générer un inventaire

### 2-11 — Gestion - Gestion niveau 3 - Projets - Paramétrage et utilisation
**Objectifs opérationnels** : Être autonome pour paramétrer les types de projets, créer et planifier les actions d'un projet, gérer le temps passé, enregistrer les commandes fournisseurs liées pour suivre la rentabilité

A l'issue du parcours, l'apprenant sera capable de :
- Créer des types de projets et leurs étapes personnalisées
- Créer un projet et planifier les actions de chaque étape
- Pointer le temps passé sur une action

---

## 8. Exemple concret : Cas GENERFEU

### Infos client

- **Société** : GENERFEU
- **Adresse** : PARC D'ACTIVITES DES ECLAPONS - 3 CHEMIN DES ECLAPONS, 69390 VOURLES
- **Représentant** : Gilles BERTRAND, Dirigeant
- **Nombre total de stagiaires** : 39
- **Tarif** : 450€ HT / demi-journée

### Répartition en 5 conventions

#### Convention 1 — Administrateurs (VISIO)

**Modules** : 1-01, 1-03, 1-05, 1-09 (4 demi-journées = 2 journées = 14h)
**Montant** : 4 × 450€ = 1 800€ HT → 2 160€ TTC (NB: convention réelle montre 900€ HT, soit 2 demi-journées facturées — à vérifier avec Estelle)

| NOM | Prénom |
|---|---|
| BERTRAND | Gilles (TNS — pas dans la convention car dirigeant) |
| DUPONT | Christelle |
| FAVRE | Audrey |
| POMARES | Laure |
| SANCHEZ | Stéphane |
| NATALE | Livio |

#### Convention 2 — Commerciaux BE Groupe 01 (SUR SITE)

**Modules** : 1-02, 1-0A (2 demi-journées = 1 journée = 7h)
**Montant** : 2 × 450€ = 900€ HT → 1 080€ TTC

| NOM | Prénom |
|---|---|
| CASTILLE | Guy |
| CHARLES | Bertrand |
| KUONY | Gilles |
| LAFARGUE | Stéphane |
| MARCHAND | Mickaël |
| SELOSSE | Frédéric |

#### Convention 3 — Commerciaux BE Groupe 02 (SUR SITE)

**Modules** : 1-02, 1-0A (2 demi-journées = 1 journée = 7h)
**Montant** : 2 × 450€ = 900€ HT → 1 080€ TTC

| NOM | Prénom |
|---|---|
| MACHADO | Jean-Philippe |
| MACQUET | Jean-Philippe |
| CASULA | Carole |
| BATAILLE | Jean-Baptiste |
| FROMENT | Lucas |

#### Convention 4 — SAV / Contrats Groupe 03 (SUR SITE)

**Modules** : 1-02, 1-10, 2-06 (3 demi-journées = 1.5 journées = 10.5h)
**Montant** : 3 × 450€ = 1 350€ HT → 1 620€ TTC

| NOM | Prénom |
|---|---|
| MARQUES | Cynthia |
| VANBECELAERE | Léa |
| DAHAN | Laetitia |
| BOUTIER | Nelly |
| REHAIMINE | Ingrid |

#### Convention 5 — Techniciens Groupe 04 (SUR SITE)

**Modules** : 1-10 (2 demi-journées = 1 journée = 7h — SAV × 2 sessions car 17 personnes)
**Montant** : 2 × 450€ = 900€ HT → 1 080€ TTC

| NOM | Prénom |
|---|---|
| AGHBALOU | Iliass |
| ALMERAS | Frédéric |
| CASTILLE | Lucas |
| DRISS | Sébastien |
| ENDRES | Sébastien |
| FOUACHE | Yann |
| HENRY | Gautier |
| JAMAY | Stéphane |
| LEMONCHE | Titouan |
| LEPOT | Nicolas |
| MORAUX | Peter |
| QUESLIN | Julien |
| RIPOLL | Vincent |
| SOUALMI | Raydoine |
| STROOBANTS | Herwig |
| TABONNI | Grégory |
| TRAIY | Youssef |

### Documents à produire pour GENERFEU

| Type de document | Quantité | Détail |
|---|---|---|
| Conventions | 5 | 1 par groupe |
| Certificats de réalisation | ~39 | 1 par stagiaire |
| Feuilles de présence | ~15 | 1 par demi-journée × par groupe |
| **TOTAL** | **~59 documents** | |

---

## 9. Informations fixes de l'organisme de formation

Ces informations sont identiques sur tous les documents, tous clients confondus :

| Champ | Valeur |
|---|---|
| Organisme de formation | eBazten (éditeur de Koban) |
| Adresse OF | 2495 route de Forcalquier, 04300 Pierrerue |
| Téléphone | 06 69 64 80 09 |
| Email | contact@koban-crm.com |
| Représentant légal | Guillaume Lecanu, Gérant |
| N° déclaration d'activité | 82691206769 (Préfet Rhône-Alpes) |
| Identifiant DataDock | 0029644 |
| SIREN | 539304683 |
| Formatrice | Estelle Lecanu |
| Signataire certificats | Estelle Lecanu, SARL Ebazten |
| Questionnaire à chaud | https://docs.google.com/forms/d/1pyD2p8OKvsyM-zOeQv2QQqEy0rprCIAs-nfM3vN6aE8/prefill |
| Questionnaire à froid | https://docs.google.com/forms/d/1N73xIjPQfW7mae6OlJaXJRn8In1lx5HFUNXL4_xJG7I/prefill |
| Tarif standard | 450€ HT / demi-journée |
| TVA | 20% |
| Dédommagement annulation | 400€ HT |
| Lieu de juridiction | Tribunal de Lyon |

---

## 10. Notes et points d'attention

### Cas particulier : TNS
Gilles BERTRAND est marqué "TNS" (Travailleur Non Salarié) dans le tableau. Il participe aux formations mais **n'apparaît pas dans la convention** en tant que stagiaire. L'outil devra gérer ce cas (exclusion des TNS de la liste des participants).

### Cas particulier : "optionnel" et "à inviter"
Certains stagiaires sont marqués "optionnel" ou "à inviter" dans le tableau au lieu de "x". L'outil devra permettre de décider si ces personnes sont incluses ou non.

### Groupes de colonnes = Groupes de convention
Dans le tableau Excel, les colonnes consécutives avec le même en-tête de ligne 1 (ex: "Administrateurs" en col 5-8) forment un **groupe de convention**. Mais attention : des colonnes "Utilisateurs" non-consécutives peuvent correspondre à des conventions différentes (ex: col 9-10 vs col 13-14 = deux groupes Commerciaux distincts). Le regroupement se fait par **les personnes cochées**, pas seulement par les en-têtes.

### Calcul du prix
Le prix affiché dans la convention des Administrateurs (900€ HT) ne correspond pas au calcul attendu (4 demi-journées × 450€ = 1 800€). Il est possible que le tarif soit par session (et non par demi-journée) ou qu'il y ait une logique de tarification différente. **Point à clarifier avec Estelle.**
