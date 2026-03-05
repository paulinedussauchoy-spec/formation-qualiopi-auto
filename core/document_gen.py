"""
document_gen.py — Génération des conventions, certificats et feuilles de présence (.docx)

Utilise docxtpl (moteur Jinja2 dans Word) pour remplir :
  - templates/convention.docx      → ConventionRenderer
  - templates/certificat.docx      → CertificatRenderer
  - templates/feuille_presence.docx → FeuillePresenceRenderer

Usage :
    from core.document_gen import ConventionRenderer, CertificatRenderer, FeuillePresenceRenderer, ClientInfo
    renderer = ConventionRenderer()
    paths = renderer.generate_all(groups, client, dates="Du 01/04/2026 au 31/10/2026")

    certif = CertificatRenderer()
    paths = certif.generate_all(groups, client, dates="Du 01/04/2026 au 31/10/2026")

    presence = FeuillePresenceRenderer()
    paths = presence.generate_all(groups, client, dates="Du 01/04/2026 au 31/10/2026")
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field, replace as _dc_replace
from datetime import date
from pathlib import Path
from typing import Optional

from docxtpl import DocxTemplate

from core.group_builder import ConventionGroup
from core.module_mapper import get_mapper

# ---------------------------------------------------------------------------
# Chemins
# ---------------------------------------------------------------------------

ROOT = Path(__file__).parent.parent
TEMPLATES_DIR = ROOT / "templates"
CONFIG_DIR = ROOT / "config"
DEFAULT_OUTPUT_DIR = ROOT / "output" / "conventions"
DEFAULT_CERTIF_DIR   = ROOT / "output" / "certificats"
DEFAULT_PRESENCE_DIR = ROOT / "output" / "feuilles_presence"


# ---------------------------------------------------------------------------
# Données client
# ---------------------------------------------------------------------------

@dataclass
class ClientInfo:
    """Informations variables du client (saisies par Estelle dans l'UI)."""
    nom: str
    adresse: str          # Adresse complète, peut contenir \n
    representant: str     # Nom + Prénom du signataire
    fonction: str         # Titre du signataire (ex: "Dirigeant")
    frais_mission_ht: float = 0.0  # Souvent 0, modifiable par Estelle


# ---------------------------------------------------------------------------
# Formatage
# ---------------------------------------------------------------------------

def _eur(amount: float) -> str:
    """
    Formate un montant en euros style français.
    Ex: 1800.0 → "1 800,00 €"
    """
    # Séparateur milliers = espace insécable, décimale = virgule
    s = f"{amount:,.2f}"          # "1,800.00"
    s = s.replace(",", "\u00A0")  # virgule milliers → espace insécable
    s = s.replace(".", ",")       # point décimal → virgule
    return f"{s} €"


def _slug(text: str) -> str:
    """Transforme un label en nom de fichier sûr."""
    text = text.replace(" — ", "_").replace("—", "_")
    text = re.sub(r"[^\w\-]", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def _modalite_display(modalite: str) -> str:
    """Traduit la modalité Excel en libellé lisible pour la convention."""
    return {
        "VISIO": "À distance",
        "SUR SITE": "Sur site",
    }.get(modalite.upper().strip(), modalite)


def _nb_heures_str(nb_heures: float) -> str:
    """Formate le nombre d'heures pour le certificat. Ex: 14.0 → '14 heures'"""
    if nb_heures == int(nb_heures):
        return f"{int(nb_heures)} heures"
    return f"{str(nb_heures).replace('.', ',')} heures"


def _duree_detail(nb_journees: float, nb_heures: float) -> str:
    """
    Génère la ligne de durée pour le tableau statistiques.
    Ex: "2 journées x 7 heures = 14 heures"
        "1,5 journée x 7 heures = 10,5 heures"
    """
    # Formatage du nombre de journées (sans .0 si entier)
    if nb_journees == int(nb_journees):
        j_str = str(int(nb_journees))
        j_label = "journée" if nb_journees == 1 else "journées"
    else:
        j_str = str(nb_journees).replace(".", ",")
        j_label = "journée" if nb_journees <= 1 else "journées"

    # Formatage des heures
    if nb_heures == int(nb_heures):
        h_str = str(int(nb_heures))
    else:
        h_str = str(nb_heures).replace(".", ",")

    return f"{j_str} {j_label} x 7 heures = {h_str} heures"


# ---------------------------------------------------------------------------
# Chargement de la config organisme
# ---------------------------------------------------------------------------

def _load_organisme() -> dict:
    path = CONFIG_DIR / "organisme.json"
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def _load_org_settings() -> dict:
    """Charge les paramètres sensibles depuis org_settings.json (gitignored)."""
    path = CONFIG_DIR / "org_settings.json"
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        data = json.load(f)
    return {k: v for k, v in data.items() if not k.startswith("_")}


# ---------------------------------------------------------------------------
# Rendu d'une convention
# ---------------------------------------------------------------------------

def generate_convention(
    group: ConventionGroup,
    client: ClientInfo,
    dates_previsionnelles: str,
    date_document: Optional[str] = None,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    template_path: Optional[Path] = None,
) -> Path:
    """
    Génère le fichier Word (.docx) d'une convention de formation.

    Args:
        group:                 ConventionGroup issu de group_builder
        client:                Infos variables du client
        dates_previsionnelles: Ex. "Du 01/04/2026 au 31/10/2026"
        date_document:         Date du document (défaut : aujourd'hui)
        output_dir:            Dossier de sortie
        template_path:         Chemin du template (défaut : templates/convention.docx)

    Returns:
        Path du fichier .docx généré
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if template_path is None:
        template_path = TEMPLATES_DIR / "convention.docx"

    if date_document is None:
        date_document = date.today().strftime("%d/%m/%Y")

    # -----------------------------------------------------------------------
    # Calculs financiers
    # -----------------------------------------------------------------------
    montant_ht = group.montant_ht
    frais_ht = client.frais_mission_ht
    total_ht = montant_ht + frais_ht
    tva = _load_organisme().get("tva", 0.20)
    total_ttc = total_ht * (1 + tva)

    # -----------------------------------------------------------------------
    # Contexte Jinja2 pour docxtpl
    # -----------------------------------------------------------------------
    context = {
        # --- Client ---
        "client_nom": client.nom,
        "client_adresse": client.adresse,
        "client_representant": client.representant,
        "client_fonction": client.fonction,

        # --- Modules (liste de dicts pour Jinja2) ---
        "modules": [
            {
                "code": m.code,
                "intitule": m.intitule,
                "objectif_operationnel": m.objectif_operationnel,
                "objectifs_texte": "\n".join(f"- {o}" for o in m.objectifs),
            }
            for m in group.modules
        ],

        # --- Statistiques ---
        "effectif": group.effectif_convention,
        "nb_journees_str": (
            str(int(group.nb_journees))
            if group.nb_journees == int(group.nb_journees)
            else str(group.nb_journees).replace(".", ",")
        ),
        "duree_detail": _duree_detail(group.nb_journees, group.nb_heures_par_stagiaire),
        "modalite": _modalite_display(group.modalite),
        "dates_previsionnelles": dates_previsionnelles,

        # --- Stagiaires ---
        "stagiaires": [
            {"nom": s.nom, "prenom": s.prenom}
            for s in group.convention_stagiaires
        ],

        # --- Montants ---
        "montant_ht_str": f"{_eur(montant_ht)} HT",
        "frais_mission_str": f"{_eur(frais_ht)} HT",
        "total_ht_str": f"{_eur(total_ht)} HT",
        "total_ttc_str": f"{_eur(total_ttc)} TTC",

        # --- Document ---
        "date_document": date_document,
    }

    # Données sensibles chargées depuis org_settings.json (gitignored)
    _org = _load_org_settings()
    context["organisme_telephone"] = _org.get("tel", "")
    context["organisme_email"]     = _org.get("email", "")

    # -----------------------------------------------------------------------
    # Rendu docxtpl
    # -----------------------------------------------------------------------
    tpl = DocxTemplate(template_path)
    tpl.render(context)

    # -----------------------------------------------------------------------
    # Nom du fichier de sortie
    # -----------------------------------------------------------------------
    label_safe = _slug(group.label)
    filename = f"CONVENTION_{_slug(client.nom)}_{label_safe}.docx"
    output_path = output_dir / filename
    tpl.save(output_path)

    return output_path


# ---------------------------------------------------------------------------
# Renderer : génère un lot de conventions
# ---------------------------------------------------------------------------

class ConventionRenderer:
    """
    Gère la génération en lot des conventions pour un client.

    Usage :
        renderer = ConventionRenderer()
        paths = renderer.generate_all(groups, client, dates, output_dir)
    """

    def __init__(self, template_path: Optional[Path] = None):
        self.template_path = template_path or (TEMPLATES_DIR / "convention.docx")
        if not self.template_path.exists():
            raise FileNotFoundError(
                f"Template introuvable : {self.template_path}\n"
                "Lancez d'abord : python scripts/build_templates.py"
            )

    def generate_all(
        self,
        groups: list[ConventionGroup],
        client: ClientInfo,
        dates_previsionnelles: str,
        date_document: Optional[str] = None,
        output_dir: Path = DEFAULT_OUTPUT_DIR,
        frais_per_group: Optional[list[float]] = None,
    ) -> list[Path]:
        """
        Génère les conventions pour tous les groupes d'un client.

        Args:
            frais_per_group: Frais de mission HT par groupe (index = index du groupe).
                             Si None ou index absent, utilise client.frais_mission_ht.

        Returns:
            Liste des chemins vers les fichiers générés (dans l'ordre des groupes).
        """
        paths = []
        for i, group in enumerate(groups):
            group_client = client
            if frais_per_group is not None and i < len(frais_per_group):
                group_client = _dc_replace(client, frais_mission_ht=frais_per_group[i])
            path = generate_convention(
                group=group,
                client=group_client,
                dates_previsionnelles=dates_previsionnelles,
                date_document=date_document,
                output_dir=output_dir,
                template_path=self.template_path,
            )
            paths.append(path)
        return paths


# ---------------------------------------------------------------------------
# Rendu d'un certificat individuel
# ---------------------------------------------------------------------------

def generate_certificat(
    group: ConventionGroup,
    stagiaire,
    client: ClientInfo,
    dates_previsionnelles: str,
    date_document: Optional[str] = None,
    ref_dossier: str = "",
    dates_realisees: Optional[str] = None,
    output_dir: Path = DEFAULT_CERTIF_DIR,
    template_path: Optional[Path] = None,
) -> Path:
    """
    Génère le certificat de réalisation (.docx) d'un stagiaire.

    Args:
        group:                 ConventionGroup du stagiaire
        stagiaire:             Objet Stagiaire (nom, prenom)
        client:                Infos variables du client
        dates_previsionnelles: Ex. "Du 01/04/2026 au 31/10/2026"
        date_document:         Date du document (défaut : aujourd'hui)
        ref_dossier:           Numéro/référence du dossier (optionnel)
        output_dir:            Dossier de sortie
        template_path:         Chemin du template (défaut : templates/certificat.docx)

    Returns:
        Path du fichier .docx généré
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if template_path is None:
        template_path = TEMPLATES_DIR / "certificat.docx"

    if date_document is None:
        date_document = date.today().strftime("%d/%m/%Y")

    # -----------------------------------------------------------------------
    # Contexte Jinja2 pour docxtpl
    # -----------------------------------------------------------------------
    nb_journees = group.nb_journees

    # Horaires des sessions (dédupliqués, ordre de première apparition)
    horaires_list = [mc.horaires for mc in group.module_columns if mc.horaires]
    horaires_uniques = list(dict.fromkeys(horaires_list))
    horaires_sessions = ", ".join(horaires_uniques) if horaires_uniques else "À définir"

    context = {
        # --- Stagiaire ---
        "stagiaire_nom": stagiaire.nom,
        "stagiaire_prenom": stagiaire.prenom,

        # --- Client ---
        "client_nom": client.nom,
        "client_adresse": client.adresse,

        # --- Modules ---
        "modules": [
            {
                "code": m.code,
                "intitule": m.intitule,
            }
            for m in group.modules
        ],

        # --- Statistiques ---
        "nb_journees_str": (
            str(int(nb_journees))
            if nb_journees == int(nb_journees)
            else str(nb_journees).replace(".", ",")
        ),
        "nb_heures_str": _nb_heures_str(group.nb_heures_par_stagiaire),
        "modalite": _modalite_display(group.modalite),
        "dates_previsionnelles": dates_previsionnelles,
        # dates_realisees = dates effectives saisies dans l'UI,
        # ou fallback sur la plage prévisionnelle
        "dates_realisees": dates_realisees if dates_realisees else dates_previsionnelles,
        "horaires_sessions": horaires_sessions,

        # --- Document ---
        "date_document": date_document,
        "ref_dossier": ref_dossier,
    }

    # -----------------------------------------------------------------------
    # Rendu docxtpl
    # -----------------------------------------------------------------------
    tpl = DocxTemplate(template_path)
    tpl.render(context)

    # -----------------------------------------------------------------------
    # Nom du fichier de sortie
    # -----------------------------------------------------------------------
    client_slug = _slug(client.nom)
    nom_slug = _slug(stagiaire.nom)
    prenom_slug = _slug(stagiaire.prenom)
    group_slug = _slug(group.label)
    filename = f"CERTIF_{client_slug}_{nom_slug}_{prenom_slug}_{group_slug}.docx"
    output_path = output_dir / filename
    tpl.save(output_path)

    return output_path


# ---------------------------------------------------------------------------
# Renderer : génère un lot de certificats
# ---------------------------------------------------------------------------

class CertificatRenderer:
    """
    Gère la génération en lot des certificats individuels pour un client.

    Génère 1 certificat par stagiaire (tous les stagiaires de tous les groupes,
    y compris les TNS qui participent aux formations).

    Usage :
        renderer = CertificatRenderer()
        paths = renderer.generate_all(groups, client, dates, output_dir)
    """

    def __init__(self, template_path: Optional[Path] = None):
        self.template_path = template_path or (TEMPLATES_DIR / "certificat.docx")
        if not self.template_path.exists():
            raise FileNotFoundError(
                f"Template introuvable : {self.template_path}\n"
                "Lancez d'abord : python scripts/build_certif_template.py"
            )

    def generate_all(
        self,
        groups: list[ConventionGroup],
        client: ClientInfo,
        dates_previsionnelles: str,
        date_document: Optional[str] = None,
        ref_dossier: str = "",
        output_dir: Path = DEFAULT_CERTIF_DIR,
        dates_realisees_per_group: Optional[list[Optional[str]]] = None,
    ) -> list[Path]:
        """
        Génère les certificats pour tous les stagiaires de tous les groupes.

        Inclut les stagiaires TNS (qui participent aux formations même s'ils
        n'apparaissent pas dans la convention).

        Args:
            dates_realisees_per_group: liste de N strings (une par groupe) contenant
                les dates effectives de session, ex. "01/04/2026, 08/04/2026".
                None pour un groupe = fallback sur dates_previsionnelles.

        Returns:
            Liste des chemins vers les fichiers générés.
        """
        paths = []
        for i, group in enumerate(groups):
            dates_realisees = (
                dates_realisees_per_group[i]
                if dates_realisees_per_group and i < len(dates_realisees_per_group)
                else None
            )
            for stagiaire in group.all_stagiaires:
                path = generate_certificat(
                    group=group,
                    stagiaire=stagiaire,
                    client=client,
                    dates_previsionnelles=dates_previsionnelles,
                    date_document=date_document,
                    ref_dossier=ref_dossier,
                    dates_realisees=dates_realisees,
                    output_dir=output_dir,
                    template_path=self.template_path,
                )
                paths.append(path)
        return paths


# ---------------------------------------------------------------------------
# Rendu d'une feuille de présence (1 par demi-journée)
# ---------------------------------------------------------------------------

def generate_feuille_presence(
    group: ConventionGroup,
    num_dj: int,
    client: ClientInfo,
    dates_previsionnelles: str,
    date_document: Optional[str] = None,
    ref_dossier: str = "",
    output_dir: Path = DEFAULT_PRESENCE_DIR,
    template_path: Optional[Path] = None,
) -> Path:
    """
    Génère la feuille de présence pour une demi-journée spécifique.

    Args:
        group:                 ConventionGroup du groupe de stagiaires
        num_dj:                Numéro de la demi-journée (1-based, index dans module_columns)
        client:                Infos variables du client
        dates_previsionnelles: Plage de dates du groupe (ex. "Du 01/04/2026 au 31/10/2026")
        date_document:         Date du document (défaut : aujourd'hui)
        ref_dossier:           Référence du dossier (optionnel)
        output_dir:            Dossier de sortie
        template_path:         Chemin du template (défaut : templates/feuille_presence.docx)

    Returns:
        Path du fichier .docx généré
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if template_path is None:
        template_path = TEMPLATES_DIR / "feuille_presence.docx"

    if date_document is None:
        date_document = date.today().strftime("%d/%m/%Y")

    # -----------------------------------------------------------------------
    # Module spécifique à cette demi-journée
    # -----------------------------------------------------------------------
    mc = group.module_columns[num_dj - 1]  # 0-based
    mapper = get_mapper()
    module = mapper.resolve(mc.type_groupe, mc.domaine)

    module_code    = module.code if module else ""
    module_intitule = module.intitule if module else mc.domaine

    # -----------------------------------------------------------------------
    # Date / horaires de la session
    # -----------------------------------------------------------------------
    date_session = mc.date or "À définir"
    if mc.date and mc.horaires:
        dates_session_detail = f"{mc.date} {mc.horaires}".strip()
    elif mc.date:
        dates_session_detail = mc.date
    else:
        dates_session_detail = dates_previsionnelles

    # -----------------------------------------------------------------------
    # Contexte Jinja2 pour docxtpl
    # -----------------------------------------------------------------------
    context = {
        # --- Client ---
        "client_nom": client.nom,
        "client_adresse": client.adresse,

        # --- Module de cette demi-journée ---
        "module_code": module_code,
        "module_intitule": module_intitule,

        # --- Stats fixes pour 1 demi-journée ---
        "nb_journees_str": "0,5",
        "nb_heures_str": "3,5 heures",
        "modalite": _modalite_display(mc.modalite),

        # --- Date de session ---
        "date_session": date_session,
        "dates_session_detail": dates_session_detail,

        # --- Stagiaires (tous, TNS inclus — présents lors des sessions) ---
        "stagiaires": [
            {"nom": s.nom, "prenom": s.prenom}
            for s in group.all_stagiaires
        ],

        # --- Document ---
        "ref_dossier": ref_dossier,
        "num_dj": num_dj,
    }

    # -----------------------------------------------------------------------
    # Rendu docxtpl
    # -----------------------------------------------------------------------
    tpl = DocxTemplate(template_path)
    tpl.render(context)

    # -----------------------------------------------------------------------
    # Nom du fichier de sortie
    # G{min_col:02d} garantit l'unicité même si deux groupes ont le même label
    # -----------------------------------------------------------------------
    client_slug = _slug(client.nom)
    group_start = min(group.col_indexes)
    group_slug  = _slug(group.label)
    filename = f"PRESENCE_{client_slug}_G{group_start:02d}_{group_slug}_DJ{num_dj:02d}.docx"
    output_path = output_dir / filename
    tpl.save(output_path)

    return output_path


# ---------------------------------------------------------------------------
# Renderer : génère un lot de feuilles de présence
# ---------------------------------------------------------------------------

class FeuillePresenceRenderer:
    """
    Gère la génération en lot des feuilles de présence pour un client.

    Génère 1 feuille par demi-journée (= par module_column de chaque groupe).
    Pour GENERFEU : 4+4+2+3+2 = 15 feuilles.

    Usage :
        renderer = FeuillePresenceRenderer()
        paths = renderer.generate_all(groups, client, dates, output_dir)
    """

    def __init__(self, template_path: Optional[Path] = None):
        self.template_path = template_path or (TEMPLATES_DIR / "feuille_presence.docx")
        if not self.template_path.exists():
            raise FileNotFoundError(
                f"Template introuvable : {self.template_path}\n"
                "Lancez d'abord : python scripts/build_presence_template.py"
            )

    def generate_all(
        self,
        groups: list[ConventionGroup],
        client: ClientInfo,
        dates_previsionnelles: str,
        date_document: Optional[str] = None,
        ref_dossier: str = "",
        output_dir: Path = DEFAULT_PRESENCE_DIR,
    ) -> list[Path]:
        """
        Génère les feuilles de présence pour tous les groupes.

        Itère sur chaque module_column de chaque groupe (1 feuille par demi-journée).

        Returns:
            Liste des chemins vers les fichiers générés (ordonnés par groupe puis demi-journée).
        """
        paths = []
        for group in groups:
            for num_dj in range(1, len(group.module_columns) + 1):
                path = generate_feuille_presence(
                    group=group,
                    num_dj=num_dj,
                    client=client,
                    dates_previsionnelles=dates_previsionnelles,
                    date_document=date_document,
                    ref_dossier=ref_dossier,
                    output_dir=output_dir,
                    template_path=self.template_path,
                )
                paths.append(path)
        return paths
