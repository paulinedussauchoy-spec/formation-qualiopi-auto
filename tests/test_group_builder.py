"""
Tests de group_builder.py — cas GENERFEU (PRD §8).

Attendu : 5 groupes
  1. Administrateurs    cols {5,6,7,8}   → 4 modules uniques, 2 journées, 14h
     - 5 convention + 1 TNS (BERTRAND exclu de la convention)
  2. Commerciaux Gr01   cols {9,10,11,12} → 2 modules uniques, 1 journée, 7h
     - 6 convention, SELOSSE inclus (ses 'à inviter' sur 16-17 ne l'excluent pas)
  3. Commerciaux Gr02   cols {13,14}      → 2 modules uniques, 1 journée, 7h
     - 5 convention
  4. SAV Contrats Gr03  cols {15,16,17}   → 3 modules uniques, 1.5 journée, 10.5h
     - 5 convention (MARQUES, DAHAN, BOUTIER, REHAIMINE = optionnel sur 1 col)
  5. Techniciens Gr04   cols {18,19}      → 1 module unique (1-10), 0.5 journée
     - 17 convention
     ⚠ PRD dit 2 demi-journées : Estelle override dans l'UI (hors scope group_builder)
"""

import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from core.excel_parser import parse_formations_excel
from core.module_mapper import ModuleMapper
from core.group_builder import build_convention_groups, ConventionGroup

EXCEL_PATH = ROOT / "base de travail" / "GENERFEU - Formations et inscrits.xlsx"
MAPPER = ModuleMapper()


def _load_groups() -> list[ConventionGroup]:
    data = parse_formations_excel(EXCEL_PATH)
    return build_convention_groups(data, MAPPER)


# ---------------------------------------------------------------------------
# Structure globale
# ---------------------------------------------------------------------------

def test_five_groups():
    groups = _load_groups()
    assert len(groups) == 5, (
        f"Attendu 5 groupes, obtenu {len(groups)}\n"
        + "\n".join(f"  {g.summary()}" for g in groups)
    )
    print(f"  {len(groups)} groupes détectés — OK")


def test_groups_sorted_by_first_col():
    groups = _load_groups()
    first_cols = [min(g.col_indexes) for g in groups]
    assert first_cols == sorted(first_cols), f"Groupes non triés : {first_cols}"
    print(f"  Ordre des groupes : {first_cols} — OK")


def test_all_39_stagiaires_covered():
    """Chaque stagiaire doit appartenir à exactement un groupe."""
    data = parse_formations_excel(EXCEL_PATH)
    groups = build_convention_groups(data, MAPPER)
    all_in_groups = [s for g in groups for s in g.all_stagiaires]
    assert len(all_in_groups) == 39, (
        f"Attendu 39 stagiaires répartis, obtenu {len(all_in_groups)}"
    )
    # Chaque stagiaire dans exactement un groupe
    names_in_groups = [s.nom_complet for g in groups for s in g.all_stagiaires]
    assert len(names_in_groups) == len(set(names_in_groups)), (
        f"Doublons détectés : {[n for n in names_in_groups if names_in_groups.count(n) > 1]}"
    )
    print(f"  39 stagiaires dans 5 groupes, sans doublon — OK")


# ---------------------------------------------------------------------------
# Groupe 1 : Administrateurs cols {5,6,7,8}
# ---------------------------------------------------------------------------

def test_groupe_admin_col_indexes():
    groups = _load_groups()
    g = groups[0]
    assert g.col_indexes == frozenset({5, 6, 7, 8}), f"Cols attendus {{5,6,7,8}}, obtenu {g.col_indexes}"
    print(f"  Admins cols : {sorted(g.col_indexes)}")


def test_groupe_admin_modules():
    groups = _load_groups()
    g = groups[0]
    assert g.codes_modules == ["1-01", "1-03", "1-09", "1-05"], (
        f"Codes attendus [1-01,1-03,1-09,1-05], obtenu {g.codes_modules}"
    )
    assert g.nb_modules_uniques == 4
    print(f"  Admins modules : {g.codes_modules}")


def test_groupe_admin_demi_journees():
    groups = _load_groups()
    g = groups[0]
    assert g.nb_demi_journees == 4, f"Attendu 4 demi-journées, obtenu {g.nb_demi_journees}"
    assert g.nb_journees == 2.0
    assert g.nb_heures_par_stagiaire == 14.0
    print(f"  Admins : {g.nb_demi_journees} DJ = {g.nb_journees} J = {g.nb_heures_par_stagiaire}h")


def test_groupe_admin_montants():
    groups = _load_groups()
    g = groups[0]
    assert g.montant_ht == 1800.0, f"Attendu 1800€ HT, obtenu {g.montant_ht}"
    assert g.montant_ttc == 2160.0
    print(f"  Admins : {g.montant_ht}€ HT / {g.montant_ttc}€ TTC")


def test_groupe_admin_tns_exclu_convention():
    groups = _load_groups()
    g = groups[0]
    assert len(g.tns_stagiaires) == 1
    assert g.tns_stagiaires[0].nom == "BERTRAND"
    assert len(g.convention_stagiaires) == 5
    noms_convention = {s.nom for s in g.convention_stagiaires}
    assert "BERTRAND" not in noms_convention
    assert noms_convention == {"DUPONT", "FAVRE", "POMARES", "SANCHEZ", "NATALE"}
    print(f"  Admins : BERTRAND TNS exclu, {g.effectif_convention} dans la convention")


def test_groupe_admin_modalite_visio():
    groups = _load_groups()
    g = groups[0]
    assert g.modalite == "VISIO", f"Attendu VISIO, obtenu {g.modalite!r}"
    print(f"  Admins modalité : {g.modalite}")


# ---------------------------------------------------------------------------
# Groupe 2 : Commerciaux BE Gr01 cols {9,10,11,12}
# ---------------------------------------------------------------------------

def test_groupe_gr01_col_indexes():
    groups = _load_groups()
    g = groups[1]
    assert g.col_indexes == frozenset({9, 10, 11, 12}), f"Obtenu {g.col_indexes}"
    print(f"  Gr01 cols : {sorted(g.col_indexes)}")


def test_groupe_gr01_deux_modules_uniques():
    """4 colonnes mais seulement 2 codes uniques : 1-02 et 1-0A."""
    groups = _load_groups()
    g = groups[1]
    assert g.nb_colonnes == 4
    assert g.nb_modules_uniques == 2
    assert set(g.codes_modules) == {"1-02", "1-0A"}
    assert g.nb_demi_journees == 2
    assert g.nb_heures_par_stagiaire == 7.0
    print(f"  Gr01 : {g.nb_colonnes} cols → {g.nb_modules_uniques} modules uniques → {g.nb_demi_journees} DJ")


def test_groupe_gr01_selosse_inclus():
    """SELOSSE est dans Gr01 malgré ses 'à inviter' sur cols 16-17."""
    groups = _load_groups()
    g = groups[1]
    noms = {s.nom for s in g.convention_stagiaires}
    assert "SELOSSE" in noms, f"SELOSSE devrait être dans Gr01, présents : {noms}"
    assert len(g.convention_stagiaires) == 6
    print(f"  Gr01 : {len(g.convention_stagiaires)} stagiaires dont SELOSSE — OK")


def test_groupe_gr01_selosse_candidat_sav():
    """SELOSSE (à inviter sur 16-17) doit apparaître comme candidat dans SAV Gr03."""
    groups = _load_groups()
    sav_group = next(g for g in groups if g.col_indexes == frozenset({15, 16, 17}))
    candidats_noms = {s.nom for s in sav_group.candidats_a_inviter}
    assert "SELOSSE" in candidats_noms, (
        f"SELOSSE devrait être candidat SAV Gr03, candidats : {candidats_noms}"
    )
    print(f"  SAV Gr03 candidats à inviter : {candidats_noms}")


def test_groupe_gr01_pas_de_tns():
    groups = _load_groups()
    g = groups[1]
    assert len(g.tns_stagiaires) == 0
    print(f"  Gr01 : pas de TNS — OK")


# ---------------------------------------------------------------------------
# Groupe 3 : Commerciaux BE Gr02 cols {13,14}
# ---------------------------------------------------------------------------

def test_groupe_gr02():
    groups = _load_groups()
    g = groups[2]
    assert g.col_indexes == frozenset({13, 14})
    assert set(g.codes_modules) == {"1-02", "1-0A"}
    assert g.nb_demi_journees == 2
    noms = {s.nom for s in g.convention_stagiaires}
    assert noms == {"MACHADO", "MACQUET", "CASULA", "BATAILLE", "FROMENT"}
    assert g.montant_ht == 900.0
    print(f"  Gr02 : {len(g.convention_stagiaires)} stagiaires, {g.montant_ht}€ HT")


# ---------------------------------------------------------------------------
# Groupe 4 : SAV Contrats Gr03 cols {15,16,17}
# ---------------------------------------------------------------------------

def test_groupe_sav_modules():
    groups = _load_groups()
    g = groups[3]
    assert g.col_indexes == frozenset({15, 16, 17})
    assert g.codes_modules == ["1-02", "1-10", "2-06"]
    assert g.nb_demi_journees == 3
    assert g.nb_heures_par_stagiaire == 10.5
    assert g.montant_ht == 1350.0
    print(f"  SAV Gr03 : modules {g.codes_modules}, {g.nb_demi_journees} DJ, {g.montant_ht}€ HT")


def test_groupe_sav_stagiaires():
    groups = _load_groups()
    g = groups[3]
    assert len(g.convention_stagiaires) == 5
    noms = {s.nom for s in g.convention_stagiaires}
    assert noms == {"MARQUES", "VANBECELAERE", "DAHAN", "BOUTIER", "REHAIMINE"}
    print(f"  SAV Gr03 : {noms}")


def test_groupe_sav_optionnels():
    """MARQUES, DAHAN, BOUTIER, REHAIMINE ont des marks 'optionnel'."""
    groups = _load_groups()
    g = groups[3]
    assert g.has_optional
    noms_optionnels = {s.nom for s in g.optional_stagiaires}
    assert noms_optionnels == {"MARQUES", "DAHAN", "BOUTIER", "REHAIMINE"}, (
        f"Optionnels attendus : MARQUES, DAHAN, BOUTIER, REHAIMINE\nObtenu : {noms_optionnels}"
    )
    assert len(g.optional_stagiaires) == 4
    print(f"  SAV Gr03 optionnels : {noms_optionnels}")


def test_groupe_sav_modalite_sur_site():
    groups = _load_groups()
    g = groups[3]
    assert g.modalite == "SUR SITE"
    print(f"  SAV Gr03 modalité : {g.modalite}")


# ---------------------------------------------------------------------------
# Groupe 5 : Techniciens Gr04 cols {18,19}
# ---------------------------------------------------------------------------

def test_groupe_techniciens_un_module_deux_colonnes():
    """Les 2 colonnes (18,19) mappent toutes les deux sur 1-10 → 1 module unique."""
    groups = _load_groups()
    g = groups[4]
    assert g.col_indexes == frozenset({18, 19})
    assert g.nb_colonnes == 2
    assert g.nb_modules_uniques == 1
    assert g.codes_modules == ["1-10"]
    # nb_demi_journees = 1 par défaut (Estelle le passera à 2 dans l'UI)
    assert g.nb_demi_journees == 1
    print(f"  Techniciens : 2 cols, 1 module unique (1-10), {g.nb_demi_journees} DJ par défaut")
    print(f"  ⚠ PRD dit 2 DJ — Estelle surcharge dans l'UI (hors scope group_builder)")


def test_groupe_techniciens_17_stagiaires():
    groups = _load_groups()
    g = groups[4]
    assert len(g.convention_stagiaires) == 17, (
        f"Attendu 17 techniciens, obtenu {len(g.convention_stagiaires)}"
    )
    assert len(g.tns_stagiaires) == 0
    print(f"  Techniciens : {len(g.convention_stagiaires)} stagiaires")


# ---------------------------------------------------------------------------
# Propriétés calculées globales
# ---------------------------------------------------------------------------

def test_total_montant_ht():
    """Montant total des 5 conventions GENERFEU."""
    groups = _load_groups()
    # Admins : 4 DJ * 450 = 1800
    # Gr01   : 2 DJ * 450 = 900
    # Gr02   : 2 DJ * 450 = 900
    # SAV    : 3 DJ * 450 = 1350
    # Tech   : 1 DJ * 450 = 450 (avant override)
    total = sum(g.montant_ht for g in groups)
    assert total == 5400.0, f"Total attendu 5400€ HT, obtenu {total}"
    print(f"  Total HT (avant override Techniciens) : {total}€")


def test_summary_affichage():
    groups = _load_groups()
    for g in groups:
        s = g.summary()
        assert "[" in s and "cols=" in s
        print(f"  {s}")


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def run_all():
    tests = [
        test_five_groups,
        test_groups_sorted_by_first_col,
        test_all_39_stagiaires_covered,
        test_groupe_admin_col_indexes,
        test_groupe_admin_modules,
        test_groupe_admin_demi_journees,
        test_groupe_admin_montants,
        test_groupe_admin_tns_exclu_convention,
        test_groupe_admin_modalite_visio,
        test_groupe_gr01_col_indexes,
        test_groupe_gr01_deux_modules_uniques,
        test_groupe_gr01_selosse_inclus,
        test_groupe_gr01_selosse_candidat_sav,
        test_groupe_gr01_pas_de_tns,
        test_groupe_gr02,
        test_groupe_sav_modules,
        test_groupe_sav_stagiaires,
        test_groupe_sav_optionnels,
        test_groupe_sav_modalite_sur_site,
        test_groupe_techniciens_un_module_deux_colonnes,
        test_groupe_techniciens_17_stagiaires,
        test_total_montant_ht,
        test_summary_affichage,
    ]

    passed = failed = 0
    for test in tests:
        name = test.__name__
        try:
            test()
            print(f"PASS  {name}")
            passed += 1
        except AssertionError as e:
            print(f"FAIL  {name}: {e}")
            failed += 1
        except Exception as e:
            import traceback
            print(f"ERROR {name}: {type(e).__name__}: {e}")
            traceback.print_exc()
            failed += 1

    print(f"\n{passed}/{passed + failed} tests passés")
    return failed == 0


if __name__ == "__main__":
    success = run_all()
    sys.exit(0 if success else 1)
