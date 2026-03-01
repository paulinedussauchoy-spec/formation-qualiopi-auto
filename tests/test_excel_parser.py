"""
Tests du parser Excel — fichier GENERFEU.

Vérifications attendues d'après le PRD :
  - 15 colonnes modules (cols 5-19, E-S)
  - 4 cols Administrateurs (modules 1-01, 1-03, 1-09, 1-05)
  - 11 cols Utilisateurs (groupes Gr01/Gr02/Gr03/Gr04)
  - 39 stagiaires (lignes 9-47)
  - BERTRAND Gilles = TNS sur ses 4 colonnes
  - MARQUES Cynthia : optionnel sur col 17 (Facturation)
  - Colonnes Interne Koban (T-X) exclues
"""

import sys
from pathlib import Path

# Ajout du répertoire racine au PYTHONPATH
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from core.excel_parser import (
    parse_formations_excel,
    MARK_TNS, MARK_CONFIRMED, MARK_OPTIONAL, MARK_TO_INVITE,
)

EXCEL_PATH = ROOT / "base de travail" / "GENERFEU - Formations et inscrits.xlsx"


def test_parse_returns_correct_client_name():
    data = parse_formations_excel(EXCEL_PATH)
    assert data.client_name == "GENERFEU", f"Attendu 'GENERFEU', obtenu '{data.client_name}'"
    print(f"  client_name = {data.client_name!r}")


def test_module_columns_count():
    data = parse_formations_excel(EXCEL_PATH)
    count = len(data.module_columns)
    assert count == 15, f"Attendu 15 colonnes modules, obtenu {count}"
    print(f"  {count} colonnes modules")


def test_module_columns_no_interne_koban():
    data = parse_formations_excel(EXCEL_PATH)
    for mc in data.module_columns:
        assert mc.type_groupe != "Interne Koban", (
            f"Colonne {mc.col_index} 'Interne Koban' aurait dû être exclue"
        )
    print("  OK — aucune colonne 'Interne Koban'")


def test_admin_columns():
    data = parse_formations_excel(EXCEL_PATH)
    admin_cols = [mc for mc in data.module_columns if mc.type_groupe == "Administrateurs"]
    assert len(admin_cols) == 4, f"Attendu 4 cols Administrateurs, obtenu {len(admin_cols)}"
    domaines = {mc.domaine for mc in admin_cols}
    expected = {"Socle commun", "CRM", "SAV", "Facturation clients / CONTRATS"}
    assert domaines == expected, f"Domaines admin inattendus: {domaines}"
    print(f"  Administrateurs : {[mc.domaine for mc in admin_cols]}")


def test_utilisateurs_columns():
    data = parse_formations_excel(EXCEL_PATH)
    util_cols = [mc for mc in data.module_columns if mc.type_groupe == "Utilisateurs"]
    assert len(util_cols) == 11, f"Attendu 11 cols Utilisateurs, obtenu {len(util_cols)}"
    print(f"  Utilisateurs : {len(util_cols)} colonnes")


def test_modalites():
    data = parse_formations_excel(EXCEL_PATH)
    admin_cols = [mc for mc in data.module_columns if mc.type_groupe == "Administrateurs"]
    for mc in admin_cols:
        assert mc.modalite == "VISIO", f"Col {mc.col_index}: attendu VISIO, obtenu {mc.modalite!r}"

    util_cols = [mc for mc in data.module_columns if mc.type_groupe == "Utilisateurs"]
    for mc in util_cols:
        assert mc.modalite == "SUR SITE", f"Col {mc.col_index}: attendu SUR SITE, obtenu {mc.modalite!r}"
    print("  Modalités : VISIO (admin) / SUR SITE (utilisateurs) — OK")


def test_duree_always_half_day():
    data = parse_formations_excel(EXCEL_PATH)
    for mc in data.module_columns:
        assert mc.duree == 0.5, f"Col {mc.col_index}: durée attendue 0.5, obtenu {mc.duree}"
    print("  Toutes les durées = 0.5 — OK")


def test_stagiaires_count():
    data = parse_formations_excel(EXCEL_PATH)
    count = len(data.stagiaires)
    assert count == 39, f"Attendu 39 stagiaires, obtenu {count}"
    print(f"  {count} stagiaires")


def test_bertrand_is_tns():
    data = parse_formations_excel(EXCEL_PATH)
    bertrand = next((s for s in data.stagiaires if s.nom == "BERTRAND"), None)
    assert bertrand is not None, "BERTRAND introuvable"
    assert bertrand.is_tns(), f"BERTRAND devrait être TNS, marks = {bertrand.modules}"
    assert len(bertrand.active_col_indexes()) == 0, "TNS ne devrait avoir aucun col actif"
    print(f"  BERTRAND Gilles = TNS sur {len(bertrand.modules)} colonnes — OK")


def test_dupont_admin_all_confirmed():
    data = parse_formations_excel(EXCEL_PATH)
    dupont = next((s for s in data.stagiaires if s.nom == "DUPONT"), None)
    assert dupont is not None, "DUPONT introuvable"
    admin_cols = [mc.col_index for mc in data.module_columns if mc.type_groupe == "Administrateurs"]
    for col_idx in admin_cols:
        mark = dupont.modules.get(col_idx)
        assert mark == MARK_CONFIRMED, (
            f"DUPONT col {col_idx}: attendu 'x', obtenu {mark!r}"
        )
    print(f"  DUPONT Christelle : {len(dupont.active_col_indexes())} cols actifs — OK")


def test_marques_has_optionnel():
    data = parse_formations_excel(EXCEL_PATH)
    marques = next((s for s in data.stagiaires if s.nom == "MARQUES"), None)
    assert marques is not None, "MARQUES introuvable"
    optional_marks = [col for col, mark in marques.modules.items() if mark == MARK_OPTIONAL]
    assert len(optional_marks) >= 1, (
        f"MARQUES devrait avoir au moins un 'optionnel', marks = {marques.modules}"
    )
    print(f"  MARQUES : {marques.modules} — optionnel sur cols {optional_marks} — OK")


def test_selosse_has_a_inviter():
    data = parse_formations_excel(EXCEL_PATH)
    selosse = next((s for s in data.stagiaires if s.nom == "SELOSSE"), None)
    assert selosse is not None, "SELOSSE introuvable"
    invite_marks = [col for col, mark in selosse.modules.items() if mark == MARK_TO_INVITE]
    assert len(invite_marks) >= 1, (
        f"SELOSSE devrait avoir au moins un 'à inviter', marks = {selosse.modules}"
    )
    print(f"  SELOSSE : 'à inviter' sur cols {invite_marks} — OK")


def test_techniciens_only_sav():
    """Les techniciens (cols 18-19) ne participent qu'aux modules SAV.
    NB : CASTILLE existe en double (Guy = commercial, Lucas = technicien).
    On identifie les techniciens par leurs colonnes actives (18-19).
    """
    data = parse_formations_excel(EXCEL_PATH)
    # Les techniciens sont ceux qui ont UNIQUEMENT des marks sur cols 18 et/ou 19
    techniciens = [
        s for s in data.stagiaires
        if s.modules and all(c in {18, 19} for c in s.modules.keys())
    ]
    assert len(techniciens) == 17, f"Attendu 17 techniciens (cols 18-19 uniquement), obtenu {len(techniciens)}: {[s.nom for s in techniciens]}"
    for s in techniciens:
        active = s.active_col_indexes()
        assert all(c in {18, 19} for c in active), (
            f"{s.nom}: attendu uniquement cols 18-19, obtenu {active}"
        )
    print(f"  {len(techniciens)} techniciens sur cols SAV (18-19) uniquement — OK")


def run_all():
    tests = [
        test_parse_returns_correct_client_name,
        test_module_columns_count,
        test_module_columns_no_interne_koban,
        test_admin_columns,
        test_utilisateurs_columns,
        test_modalites,
        test_duree_always_half_day,
        test_stagiaires_count,
        test_bertrand_is_tns,
        test_dupont_admin_all_confirmed,
        test_marques_has_optionnel,
        test_selosse_has_a_inviter,
        test_techniciens_only_sav,
    ]

    passed = 0
    failed = 0
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
            print(f"ERROR {name}: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{passed}/{passed + failed} tests passés")
    return failed == 0


if __name__ == "__main__":
    success = run_all()
    sys.exit(0 if success else 1)
