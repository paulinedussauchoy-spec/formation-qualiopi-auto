"""
group_builder.py — Regroupement des stagiaires en conventions

Principe (PRD § architecture) :
  Deux stagiaires avec exactement les mêmes colonnes cochées → même convention.
  Le regroupement se fait sur les personnes, pas sur les en-têtes de colonnes.

Clé de groupement :
  frozenset des col_indexes où mark ∈ {x, optionnel, TNS}.
  "à inviter" est EXCLU de la clé → SELOSSE (x sur 9-12, à inviter sur 16-17)
  reste dans Gr01 {9,10,11,12}, pas dans SAV Gr03.

Calcul des demi-journées :
  = nombre de colonnes Excel (demi-journées de formation réelles).
  Exemple Gr01 : 4 colonnes (9,10,11,12) → 4 demi-journées → 2 journées / 14h.
  Un même code module peut apparaître sur plusieurs colonnes (ex: 1-02 × 2 = 2 DJ).
  ARCHITECTURE.md précise qu'Estelle peut modifier ce chiffre dans l'interface.

Cas particuliers gérés :
  - TNS (ex: BERTRAND) : dans le groupe (participe) mais absent de la convention
  - optionnel / à inviter : inclus par défaut, Estelle peut décocher dans l'UI
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from core.excel_parser import (
    ExcelData, ModuleColumn, Stagiaire,
    MARK_CONFIRMED, MARK_TNS, MARK_OPTIONAL, MARK_TO_INVITE,
)
from core.module_mapper import Module, ModuleMapper

# Tarif standard — peut être surchargé par convention
TARIF_DEMI_JOURNEE_HT = 450.0
TVA = 0.20


# ---------------------------------------------------------------------------
# Dataclass ConventionGroup
# ---------------------------------------------------------------------------

@dataclass
class ConventionGroup:
    """
    Un groupe de stagiaires partageant exactement le même ensemble de colonnes
    cochées dans le tableau Excel — correspond à une convention de formation.
    """
    # Colonnes Excel (1-based) qui définissent ce groupe
    col_indexes: frozenset[int]

    # Métadonnées des colonnes (triées par col_index)
    module_columns: list[ModuleColumn]

    # Modules uniques de ce groupe (dédupliqués par code, ordonnés par col_index)
    modules: list[Module]

    # Tous les stagiaires du groupe (y compris TNS)
    all_stagiaires: list[Stagiaire]

    # Stagiaires à inscrire dans la convention (TNS exclus)
    convention_stagiaires: list[Stagiaire]

    # Stagiaires marqués TNS (participent mais n'apparaissent pas dans la convention)
    tns_stagiaires: list[Stagiaire]

    # Stagiaires optionnels ou "à inviter" dans les colonnes de CE groupe
    optional_stagiaires: list[Stagiaire]

    # Stagiaires d'un AUTRE groupe ayant "à inviter" sur des colonnes de ce groupe
    # (suggestion pour Estelle : elle peut les ajouter manuellement)
    candidats_a_inviter: list[Stagiaire] = field(default_factory=list)

    # Modalité de formation (ex: "VISIO", "SUR SITE")
    modalite: str = ""

    # Nom suggéré (calculé, modifiable par Estelle dans l'UI)
    label: str = ""

    # Tarif (modifiable par Estelle dans l'UI)
    tarif_demi_journee_ht: float = TARIF_DEMI_JOURNEE_HT

    # Nombre de demi-journées facturées (= nb_modules_uniques par défaut, modifiable)
    nb_demi_journees: int = 0

    # -----------------------------------------------------------------------
    # Propriétés calculées
    # -----------------------------------------------------------------------

    @property
    def nb_modules_uniques(self) -> int:
        return len(self.modules)

    @property
    def nb_colonnes(self) -> int:
        """Nombre brut de colonnes Excel (avant déduplication)."""
        return len(self.col_indexes)

    @property
    def nb_journees(self) -> float:
        return self.nb_demi_journees / 2

    @property
    def nb_heures_par_stagiaire(self) -> float:
        return self.nb_journees * 7

    @property
    def montant_ht(self) -> float:
        return self.nb_demi_journees * self.tarif_demi_journee_ht

    @property
    def montant_ttc(self) -> float:
        return self.montant_ht * (1 + TVA)

    @property
    def effectif_convention(self) -> int:
        """Nombre de stagiaires inscrits dans la convention (hors TNS)."""
        return len(self.convention_stagiaires)

    @property
    def has_optional(self) -> bool:
        return bool(self.optional_stagiaires)

    @property
    def codes_modules(self) -> list[str]:
        return [m.code for m in self.modules]

    def col_to_module(self, col_index: int) -> Optional[Module]:
        """Retourne le module résolu pour une colonne donnée."""
        for mc, module in zip(self.module_columns, _expand_modules(self.module_columns, self.modules)):
            if mc.col_index == col_index:
                return module
        return None

    def summary(self) -> str:
        """Résumé lisible pour le debug / logs."""
        return (
            f"[{self.label}] "
            f"cols={sorted(self.col_indexes)} "
            f"modules={self.codes_modules} "
            f"{self.nb_demi_journees} DJ "
            f"{self.effectif_convention} stagiaires "
            f"({self.montant_ht:.0f}€ HT)"
        )


def _expand_modules(
    module_columns: list[ModuleColumn],
    unique_modules: list[Module],
) -> list[Optional[Module]]:
    """
    Mappe chaque ModuleColumn à son Module unique (par correspondance de code).
    Retourne une liste parallèle à module_columns.
    """
    code_to_module = {m.code: m for m in unique_modules}
    return [code_to_module.get(_col_code(mc)) for mc in module_columns]


def _col_code(mc: ModuleColumn) -> str:
    """Placeholder — la résolution réelle est faite via le mapper."""
    return ""  # sera remplacé dans build_convention_groups


# ---------------------------------------------------------------------------
# Fonction principale
# ---------------------------------------------------------------------------

def build_convention_groups(
    data: ExcelData,
    mapper: ModuleMapper,
) -> list[ConventionGroup]:
    """
    Regroupe les stagiaires en conventions à partir du résultat du parser Excel.

    Args:
        data:   Résultat de parse_formations_excel()
        mapper: Instance de ModuleMapper chargée

    Returns:
        Liste de ConventionGroup triée par premier col_index (ordre Excel).
        Chaque groupe correspond à une convention à générer.
    """
    # --- Étape 1 : Calculer la clé de groupement pour chaque stagiaire ---
    # Clé = frozenset des colonnes où mark ∈ {x, optionnel, TNS}
    GROUPING_MARKS = {MARK_CONFIRMED, MARK_OPTIONAL, MARK_TNS}

    groups_raw: dict[frozenset, list[Stagiaire]] = {}

    for stagiaire in data.stagiaires:
        key = frozenset(
            col_idx
            for col_idx, mark in stagiaire.modules.items()
            if mark in GROUPING_MARKS
        )
        if not key:
            continue  # Stagiaire sans mark pertinent (ex: seulement 'à inviter')
        groups_raw.setdefault(key, []).append(stagiaire)

    # --- Étape 2 : Construire les ConventionGroup ---
    result: list[ConventionGroup] = []

    for col_indexes, stagiaires_list in groups_raw.items():

        # Colonnes Excel correspondantes (dans le bon ordre)
        module_cols = sorted(
            [mc for mc in data.module_columns if mc.col_index in col_indexes],
            key=lambda mc: mc.col_index,
        )

        # Modules uniques (dédupliqués par code, ordre de première apparition)
        seen_codes: set[str] = set()
        unique_modules: list[Module] = []
        for mc in module_cols:
            module = mapper.resolve(mc.type_groupe, mc.domaine)
            if module is not None and module.code not in seen_codes:
                seen_codes.add(module.code)
                unique_modules.append(module)

        # Catégorisation des stagiaires
        convention_stagiaires = [s for s in stagiaires_list if not s.is_tns()]
        tns_stagiaires = [s for s in stagiaires_list if s.is_tns()]

        optional_stagiaires = [
            s for s in convention_stagiaires
            if any(
                mark in (MARK_OPTIONAL, MARK_TO_INVITE)
                for col_idx, mark in s.modules.items()
                if col_idx in col_indexes
            )
        ]

        # Candidats "à inviter" : stagiaires d'AUTRES groupes ayant
        # un mark "à inviter" sur au moins une colonne de CE groupe
        candidats = [
            s for s in data.stagiaires
            if s not in stagiaires_list
            and any(
                col_idx in col_indexes and mark == MARK_TO_INVITE
                for col_idx, mark in s.modules.items()
            )
        ]

        # Modalité (doit être cohérente dans le groupe — on prend la première)
        modalite = module_cols[0].modalite if module_cols else ""

        # nb_demi_journees = nb colonnes Excel = nb demi-journées réelles (modifiable par Estelle)
        nb_demi_journees = len(module_cols)

        # Label suggéré
        label = _suggest_label(module_cols, unique_modules)

        result.append(ConventionGroup(
            col_indexes=col_indexes,
            module_columns=module_cols,
            modules=unique_modules,
            all_stagiaires=stagiaires_list,
            convention_stagiaires=convention_stagiaires,
            tns_stagiaires=tns_stagiaires,
            optional_stagiaires=optional_stagiaires,
            candidats_a_inviter=candidats,
            modalite=modalite,
            label=label,
            nb_demi_journees=nb_demi_journees,
        ))

    # Tri par premier col_index → ordre cohérent avec l'Excel
    result.sort(key=lambda g: min(g.col_indexes))
    return result


# ---------------------------------------------------------------------------
# Génération du label suggéré
# ---------------------------------------------------------------------------

def _suggest_label(
    module_cols: list[ModuleColumn],
    unique_modules: list[Module],
) -> str:
    """
    Génère un label lisible à partir du type de groupe et des modules.

    Exemples :
      "Administrateurs"          (4 modules admin)
      "Utilisateurs — 1-02, 1-0A" (Commerciaux)
      "Utilisateurs — 1-02, 1-10, 2-06" (SAV)
    """
    if not module_cols:
        return "Groupe inconnu"

    type_groupe = module_cols[0].type_groupe
    codes = ", ".join(m.code for m in unique_modules)

    if type_groupe == "Administrateurs":
        return "Administrateurs"
    return f"{type_groupe} — {codes}"
