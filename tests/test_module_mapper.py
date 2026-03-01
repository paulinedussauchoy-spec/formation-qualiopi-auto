"""
Tests de module_mapper.py — PRD sections 6 et 7.

Vérifie :
  - Les 8 correspondances de mapping GENERFEU (PRD section 6)
  - Le catalogue complet (16 modules, PRD section 7)
  - La robustesse de la normalisation (casse, espaces, variantes)
  - Les cas non mappés (retour None)
  - L'intégration avec les colonnes réelles de l'Excel GENERFEU
"""

import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from core.module_mapper import ModuleMapper, get_mapper

MAPPER = ModuleMapper()


# ---------------------------------------------------------------------------
# Catalogue des modules
# ---------------------------------------------------------------------------

def test_catalogue_has_16_modules():
    modules = MAPPER.all_modules()
    assert len(modules) == 16, f"Attendu 16 modules, obtenu {len(modules)}: {[m.code for m in modules]}"
    print(f"  {len(modules)} modules dans le catalogue")


def test_all_module_codes_present():
    expected_codes = {
        "1-01", "1-02", "1-03", "1-0A", "1-05", "1-06", "1-07",
        "1-09", "1-10", "2-05", "2-06", "2-07", "2-08", "2-09",
        "2-10", "2-11",
    }
    known = MAPPER.known_codes()
    missing = expected_codes - known
    assert not missing, f"Codes manquants : {missing}"
    print(f"  Codes présents : {sorted(known)}")


def test_each_module_has_required_fields():
    for module in MAPPER.all_modules():
        assert module.code, f"Code vide pour {module}"
        assert module.intitule, f"Intitulé vide pour {module.code}"
        assert module.objectif_operationnel, f"Objectif opérationnel vide pour {module.code}"
        assert len(module.objectifs) >= 1, f"Aucun objectif pour {module.code}"
    print("  Tous les modules ont les champs requis")


def test_objectifs_texte_format():
    m = MAPPER.get_module("1-01")
    texte = m.objectifs_texte()
    lines = texte.strip().split("\n")
    assert all(line.startswith("- ") for line in lines), f"Format inattendu : {texte}"
    print(f"  objectifs_texte() → {len(lines)} lignes formatées")


def test_intitule_court():
    m = MAPPER.get_module("1-01")
    assert m.intitule_court() == "Administration socle commun", m.intitule_court()
    m2 = MAPPER.get_module("1-10")
    assert m2.intitule_court() == "Utilisation service client SAV", m2.intitule_court()
    print("  intitule_court() OK")


# ---------------------------------------------------------------------------
# Mapping GENERFEU (PRD section 6) — 8 règles
# ---------------------------------------------------------------------------

def test_mapping_admin_socle_commun():
    m = MAPPER.resolve("Administrateurs", "Socle commun")
    assert m is not None
    assert m.code == "1-01", f"Attendu 1-01, obtenu {m.code}"
    print(f"  Administrateurs + Socle commun → {m.code} ({m.intitule})")


def test_mapping_admin_crm():
    m = MAPPER.resolve("Administrateurs", "CRM")
    assert m is not None and m.code == "1-03"
    print(f"  Administrateurs + CRM → {m.code}")


def test_mapping_admin_sav():
    m = MAPPER.resolve("Administrateurs", "SAV")
    assert m is not None and m.code == "1-09"
    print(f"  Administrateurs + SAV → {m.code}")


def test_mapping_admin_facturation():
    m = MAPPER.resolve("Administrateurs", "Facturation clients / CONTRATS")
    assert m is not None and m.code == "1-05"
    print(f"  Administrateurs + Facturation clients / CONTRATS → {m.code}")


def test_mapping_utilisateurs_socle_commun():
    m = MAPPER.resolve("Utilisateurs", "Socle commun")
    assert m is not None and m.code == "1-02"
    print(f"  Utilisateurs + Socle commun → {m.code}")


def test_mapping_utilisateurs_crm():
    m = MAPPER.resolve("Utilisateurs", "CRM")
    assert m is not None and m.code == "1-0A"
    print(f"  Utilisateurs + CRM → {m.code}")


def test_mapping_utilisateurs_sav():
    m = MAPPER.resolve("Utilisateurs", "SAV")
    assert m is not None and m.code == "1-10"
    print(f"  Utilisateurs + SAV → {m.code}")


def test_mapping_utilisateurs_sav_triple_plus():
    """'SAV +++' est une variante de SAV pour Utilisateurs → même code 1-10."""
    m = MAPPER.resolve("Utilisateurs", "SAV +++")
    assert m is not None and m.code == "1-10", f"SAV +++ devrait mapper sur 1-10, obtenu {m}"
    print(f"  Utilisateurs + SAV +++ → {m.code} (variante OK)")


def test_mapping_utilisateurs_facturation():
    m = MAPPER.resolve("Utilisateurs", "Facturation clients / CONTRATS")
    assert m is not None and m.code == "2-06"
    print(f"  Utilisateurs + Facturation clients / CONTRATS → {m.code}")


# ---------------------------------------------------------------------------
# Robustesse de la normalisation
# ---------------------------------------------------------------------------

def test_normalisation_casse():
    """La casse ne doit pas affecter le résultat."""
    m1 = MAPPER.resolve("Administrateurs", "Socle commun")
    m2 = MAPPER.resolve("administrateurs", "socle commun")
    m3 = MAPPER.resolve("ADMINISTRATEURS", "SOCLE COMMUN")
    assert m1 == m2 == m3, "La normalisation de casse échoue"
    print("  Normalisation casse — OK")


def test_normalisation_espaces():
    """Les espaces superflus ne doivent pas affecter le résultat."""
    m1 = MAPPER.resolve("Administrateurs", "Socle commun")
    m2 = MAPPER.resolve("  Administrateurs  ", "  Socle commun  ")
    assert m1 == m2, "La normalisation des espaces échoue"
    print("  Normalisation espaces — OK")


def test_normalisation_slash_avec_espaces():
    """'Facturation clients / CONTRATS' avec espaces autour du slash."""
    m1 = MAPPER.resolve("Utilisateurs", "Facturation clients / CONTRATS")
    m2 = MAPPER.resolve("Utilisateurs", "Facturation clients/CONTRATS")
    assert m1 is not None and m1.code == "2-06"
    assert m2 is not None and m2.code == "2-06"
    print(f"  Variantes slash → {m1.code} — OK")


# ---------------------------------------------------------------------------
# Cas non mappés
# ---------------------------------------------------------------------------

def test_unknown_type_groupe_returns_none():
    m = MAPPER.resolve("Interne Koban", "Evaluation finale /5")
    assert m is None, f"'Interne Koban' devrait retourner None, obtenu {m}"
    print("  Interne Koban → None — OK")


def test_unknown_domaine_returns_none():
    m = MAPPER.resolve("Utilisateurs", "Module inconnu XYZ")
    assert m is None
    print("  Domaine inconnu → None — OK")


def test_resolve_or_raise_raises_on_unknown():
    try:
        MAPPER.resolve_or_raise("Utilisateurs", "Module inconnu")
        assert False, "Aurait dû lever ValueError"
    except ValueError as e:
        print(f"  resolve_or_raise lève ValueError — OK ({e})")


# ---------------------------------------------------------------------------
# Intégration avec l'Excel GENERFEU réel
# ---------------------------------------------------------------------------

def test_integration_all_generfeu_module_columns():
    """
    Résout les 15 colonnes modules de GENERFEU depuis l'Excel réel.
    Toutes les colonnes formation (pas Interne Koban) doivent mapper.
    """
    from core.excel_parser import parse_formations_excel

    excel_path = ROOT / "base de travail" / "GENERFEU - Formations et inscrits.xlsx"
    data = parse_formations_excel(excel_path)

    unresolved = []
    resolved = []
    for mc in data.module_columns:
        module = MAPPER.resolve(mc.type_groupe, mc.domaine)
        if module is None:
            unresolved.append((mc.col_index, mc.type_groupe, mc.domaine))
        else:
            resolved.append((mc.col_index, module.code))

    assert not unresolved, (
        f"Colonnes non résolues : {unresolved}"
    )
    print(f"  {len(resolved)} colonnes résolues :")
    for col_idx, code in sorted(resolved):
        print(f"    col {col_idx} → {code}")


def test_singleton_get_mapper():
    m1 = get_mapper()
    m2 = get_mapper()
    assert m1 is m2, "get_mapper() doit retourner la même instance"
    print("  get_mapper() singleton — OK")


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def run_all():
    tests = [
        test_catalogue_has_16_modules,
        test_all_module_codes_present,
        test_each_module_has_required_fields,
        test_objectifs_texte_format,
        test_intitule_court,
        test_mapping_admin_socle_commun,
        test_mapping_admin_crm,
        test_mapping_admin_sav,
        test_mapping_admin_facturation,
        test_mapping_utilisateurs_socle_commun,
        test_mapping_utilisateurs_crm,
        test_mapping_utilisateurs_sav,
        test_mapping_utilisateurs_sav_triple_plus,
        test_mapping_utilisateurs_facturation,
        test_normalisation_casse,
        test_normalisation_espaces,
        test_normalisation_slash_avec_espaces,
        test_unknown_type_groupe_returns_none,
        test_unknown_domaine_returns_none,
        test_resolve_or_raise_raises_on_unknown,
        test_integration_all_generfeu_module_columns,
        test_singleton_get_mapper,
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
            print(f"ERROR {name}: {type(e).__name__}: {e}")
            failed += 1

    print(f"\n{passed}/{passed + failed} tests passés")
    return failed == 0


if __name__ == "__main__":
    success = run_all()
    sys.exit(0 if success else 1)
