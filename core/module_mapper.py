"""
module_mapper.py — Correspondance colonnes Excel → modules Koban

Charge config/modules.json et expose :
  - resolve(type_groupe, domaine) → Module
  - get_module(code) → Module
  - all_modules() → list[Module]

La normalisation des clés de mapping gère les variantes rencontrées dans
les fichiers Excel d'Estelle (casse, espaces, "SAV +++" vs "SAV", etc.).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

# Chemin du fichier de config (relatif à ce fichier)
_CONFIG_PATH = Path(__file__).parent.parent / "config" / "modules.json"


# ---------------------------------------------------------------------------
# Dataclass Module
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Module:
    code: str
    intitule: str
    objectif_operationnel: str
    objectifs: tuple[str, ...]

    def objectifs_texte(self, separator: str = "\n") -> str:
        """Objectifs formatés en liste à puces."""
        return separator.join(f"- {o}" for o in self.objectifs)

    def intitule_court(self) -> str:
        """Partie après le premier ' - ' (ex: 'Administration socle commun')."""
        parts = self.intitule.split(" - ", 1)
        return parts[1] if len(parts) > 1 else self.intitule


# ---------------------------------------------------------------------------
# Normalisation des clés de mapping
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    """Normalise une chaîne pour la comparaison : minuscules, espaces uniques, sans ponctuation finale."""
    text = text.strip().lower()
    text = re.sub(r"\s+", " ", text)          # espaces multiples → un seul
    text = re.sub(r"\s*/\s*", "/", text)       # "clients / CONTRATS" → "clients/CONTRATS"
    return text


# ---------------------------------------------------------------------------
# ModuleMapper
# ---------------------------------------------------------------------------

class ModuleMapper:
    """
    Charge le référentiel depuis modules.json et résout les correspondances
    (type_groupe + domaine) → Module.

    Usage :
        mapper = ModuleMapper()
        module = mapper.resolve("Administrateurs", "Socle commun")
        # → Module(code="1-01", intitule="Socle commun - Administration socle commun", …)
    """

    def __init__(self, config_path: str | Path = _CONFIG_PATH):
        path = Path(config_path)
        if not path.exists():
            raise FileNotFoundError(f"Fichier de config introuvable : {path}")

        with path.open(encoding="utf-8") as f:
            data = json.load(f)

        # Catalogue des modules : code → Module
        self._modules: dict[str, Module] = {}
        for code, entry in data["modules"].items():
            self._modules[code] = Module(
                code=code,
                intitule=entry["intitule"],
                objectif_operationnel=entry["objectif_operationnel"],
                objectifs=tuple(entry["objectifs"]),
            )

        # Table de mapping : (type_groupe_norm, domaine_norm) → code
        # Construite à partir du tableau "mapping" du JSON
        self._mapping: dict[tuple[str, str], str] = {}
        for rule in data["mapping"]:
            type_norm = _normalize(rule["type_groupe"])
            code = rule["code"]
            for domaine in rule["domaines"]:
                key = (type_norm, _normalize(domaine))
                self._mapping[key] = code

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------

    def resolve(self, type_groupe: str, domaine: str) -> Optional[Module]:
        """
        Retourne le Module correspondant à (type_groupe, domaine).
        Retourne None si la combinaison est inconnue.

        Args:
            type_groupe: Valeur de la ligne 1 Excel (ex: "Administrateurs")
            domaine:     Valeur de la ligne 2 Excel (ex: "Socle commun")
        """
        key = (_normalize(type_groupe), _normalize(domaine))
        code = self._mapping.get(key)
        if code is None:
            return None
        return self._modules.get(code)

    def resolve_or_raise(self, type_groupe: str, domaine: str) -> Module:
        """
        Comme resolve() mais lève ValueError si la combinaison est inconnue.
        Utile en génération de documents où un mapping manquant est bloquant.
        """
        module = self.resolve(type_groupe, domaine)
        if module is None:
            raise ValueError(
                f"Aucun module trouvé pour type_groupe={type_groupe!r}, domaine={domaine!r}. "
                f"Vérifiez la table de mapping dans config/modules.json."
            )
        return module

    def get_module(self, code: str) -> Optional[Module]:
        """Retourne un module par son code (ex: '1-01'). None si inconnu."""
        return self._modules.get(code)

    def all_modules(self) -> list[Module]:
        """Liste tous les modules du catalogue, triés par code."""
        return sorted(self._modules.values(), key=lambda m: m.code)

    def known_codes(self) -> set[str]:
        """Ensemble des codes modules connus."""
        return set(self._modules.keys())


# ---------------------------------------------------------------------------
# Instance partagée (singleton léger — chargé une seule fois)
# ---------------------------------------------------------------------------

_default_mapper: Optional[ModuleMapper] = None


def get_mapper() -> ModuleMapper:
    """Retourne l'instance partagée du mapper (chargée à la première utilisation)."""
    global _default_mapper
    if _default_mapper is None:
        _default_mapper = ModuleMapper()
    return _default_mapper
