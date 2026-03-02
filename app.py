"""
app.py — Interface Streamlit : générateur de documents Qualiopi eBazten.

Flux utilisateur (voir ARCHITECTURE.md) :
  Phase A (avant formation) : upload Excel → infos client → aperçu groupes → conventions ZIP
  Phase B (après formation) : même Excel   → certificats + feuilles de présence ZIP
"""

from __future__ import annotations

import io
import re
import subprocess
import sys
import tempfile
import zipfile
from copy import deepcopy
from datetime import date as date_cls
from pathlib import Path
from typing import Optional

import streamlit as st

# ---------------------------------------------------------------------------
# Chemins & imports core
# ---------------------------------------------------------------------------

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from core.excel_parser import parse_formations_excel
from core.group_builder import ConventionGroup, build_convention_groups
from core.module_mapper import ModuleMapper
from core.document_gen import (
    CertificatRenderer,
    ClientInfo,
    ConventionRenderer,
    FeuillePresenceRenderer,
)

# ---------------------------------------------------------------------------
# Configuration de la page
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Koban Qualiopi — Documents",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# Helpers — LibreOffice
# ---------------------------------------------------------------------------

def _find_libreoffice() -> Optional[str]:
    """Cherche l'exécutable LibreOffice sur le système."""
    candidates = [
        "soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        try:
            r = subprocess.run([c, "--version"], capture_output=True, timeout=5)
            if r.returncode == 0:
                return c
        except Exception:
            pass
    return None


def _convert_to_pdf(docx_path: Path, out_dir: Path, soffice: str) -> Optional[Path]:
    """Convertit un .docx en .pdf via LibreOffice headless. Retourne None en cas d'échec."""
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(docx_path)],
            capture_output=True, timeout=60,
        )
        pdf = out_dir / (docx_path.stem + ".pdf")
        return pdf if pdf.exists() else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Helpers — ZIP
# ---------------------------------------------------------------------------

def _build_zip(folders: dict[str, list[Path]], soffice: Optional[str] = None) -> bytes:
    """
    Construit un ZIP en mémoire.

    Args:
        folders: dict { nom_dossier_dans_zip → [chemins .docx] }
        soffice: chemin soffice si PDF souhaité, None sinon
    Returns:
        Bytes du ZIP
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for folder, paths in folders.items():
            for p in paths:
                zf.write(p, arcname=f"{folder}/{p.name}")
                if soffice:
                    pdf = _convert_to_pdf(p, p.parent, soffice)
                    if pdf:
                        zf.write(pdf, arcname=f"{folder}/{pdf.name}")
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Helpers — Session state & clés de widgets
# ---------------------------------------------------------------------------

def _uid() -> int:
    """Identifiant de l'upload courant (change à chaque nouveau fichier)."""
    return st.session_state.get("upload_count", 0)


def _wkey(*parts) -> str:
    """Génère une clé de widget préfixée par l'upload_id pour éviter les conflits."""
    return f"u{_uid()}_" + "_".join(str(p) for p in parts)


def _slug(text: str) -> str:
    text = re.sub(r"[^\w]", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


# ---------------------------------------------------------------------------
# Helpers — Filtrage des stagiaires optionnels
# ---------------------------------------------------------------------------

def _get_filtered_groups(groups: list[ConventionGroup]) -> list[ConventionGroup]:
    """
    Retourne des copies des groupes avec les stagiaires optionnels
    décochés par l'utilisateur exclus de convention_stagiaires.
    """
    result = []
    for i, group in enumerate(groups):
        g = deepcopy(group)
        excluded = set()
        for s in g.optional_stagiaires:
            key = _wkey("stag", i, s.nom, s.prenom)
            if not st.session_state.get(key, True):
                excluded.add((s.nom, s.prenom))
        if excluded:
            g.convention_stagiaires = [
                s for s in g.convention_stagiaires
                if (s.nom, s.prenom) not in excluded
            ]
        result.append(g)
    return result


# ---------------------------------------------------------------------------
# Section aperçu : 1 expander par groupe
# ---------------------------------------------------------------------------

def _render_group(i: int, group: ConventionGroup) -> None:
    """Affiche un expander pour un groupe avec édition des DJ et stagiaires."""
    montant = group.nb_demi_journees * group.tarif_demi_journee_ht
    ttc     = montant * 1.20

    with st.expander(
        f"**{group.label}** — "
        f"{group.effectif_convention} stagiaires · "
        f"{group.nb_demi_journees} DJ · "
        f"{montant:,.0f} € HT",
        expanded=True,
    ):
        col_l, col_r = st.columns([1, 1])

        # ── Colonne gauche : stats + modules
        with col_l:
            new_dj = st.number_input(
                "Demi-journées facturées",
                min_value=1, max_value=30,
                value=group.nb_demi_journees,
                step=1,
                key=_wkey("dj", i),
                help="Valeur déduite automatiquement de l'Excel. Modifiable.",
            )
            group.nb_demi_journees = int(new_dj)

            m_ht  = group.nb_demi_journees * group.tarif_demi_journee_ht
            m_ttc = m_ht * 1.20
            col_m1, col_m2 = st.columns(2)
            col_m1.metric("Montant HT",  f"{m_ht:,.0f} €")
            col_m2.metric("Montant TTC", f"{m_ttc:,.0f} €")

            st.write(
                f"**Durée :** {group.nb_journees} j · "
                f"{group.nb_heures_par_stagiaire:.0f} h / stagiaire · "
                f"**Modalité :** {group.modalite}"
            )

            st.write("**Modules :**")
            for m in group.modules:
                st.markdown(f"- `{m.code}` — {m.intitule}")

        # ── Colonne droite : stagiaires
        with col_r:
            st.write("**Stagiaires (convention) :**")
            for s in group.convention_stagiaires:
                is_opt = s in group.optional_stagiaires
                label  = f"{s.nom} {s.prenom}"
                if is_opt:
                    label += " *(optionnel)*"
                st.checkbox(
                    label,
                    value=True,
                    key=_wkey("stag", i, s.nom, s.prenom),
                )

            if group.tns_stagiaires:
                st.write("**TNS** *(hors convention, présents aux sessions)* :")
                for s in group.tns_stagiaires:
                    st.markdown(f"- {s.nom} {s.prenom}")

            if group.candidats_a_inviter:
                st.write("**Candidats à inviter** *(suggestion)* :")
                for s in group.candidats_a_inviter:
                    st.markdown(f"- {s.nom} {s.prenom}")

        # ── Dates des sessions (feuilles de présence) ────────────────────────
        st.markdown("---")
        st.write("**Dates des sessions (feuilles de présence) :**")
        n_dj = len(group.module_columns)
        dj_cols = st.columns(min(n_dj, 4))
        for j, mc in enumerate(group.module_columns):
            dj_label = f"DJ{j + 1:02d}"
            if mc.domaine:
                dj_label += f" — {mc.domaine}"
            # Valeur par défaut : date issue de l'Excel si disponible
            default_date = None
            if mc.date:
                try:
                    from datetime import datetime as _dt
                    default_date = _dt.strptime(mc.date, "%d/%m/%Y").date()
                except Exception:
                    pass
            dj_cols[j % 4].date_input(
                dj_label,
                value=default_date,
                format="DD/MM/YYYY",
                key=_wkey("date_dj", i, j),
            )


# ---------------------------------------------------------------------------
# Génération — Phase A : Conventions
# ---------------------------------------------------------------------------

def _do_generate_conventions(
    groups: list[ConventionGroup],
    client: ClientInfo,
    dates_prev: str,
    date_doc: str,
    ref: str,
    soffice: Optional[str],
) -> None:
    """Génère les conventions et stocke le ZIP dans session_state."""
    try:
        renderer = ConventionRenderer()
    except FileNotFoundError as e:
        st.error(str(e))
        return

    filtered = _get_filtered_groups(groups)

    with st.spinner("Génération des conventions..."):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                out = Path(tmp)
                paths = renderer.generate_all(
                    groups=filtered,
                    client=client,
                    dates_previsionnelles=dates_prev,
                    date_document=date_doc,
                    output_dir=out,
                )
                zip_bytes = _build_zip({"conventions": paths}, soffice)

            st.session_state["convention_zip"]   = zip_bytes
            st.session_state["convention_count"] = len(paths)
            st.success(f"{len(paths)} convention(s) générée(s).")

        except Exception as exc:
            st.error(f"Erreur lors de la génération : {exc}")
            st.code(__import__("traceback").format_exc())


# ---------------------------------------------------------------------------
# Génération — Phase B : Certificats + Feuilles de présence
# ---------------------------------------------------------------------------

def _apply_dj_dates(
    groups: list[ConventionGroup],
    group_offset: int = 0,
) -> list[ConventionGroup]:
    """
    Retourne des copies des groupes avec mc.date mis à jour
    depuis les champs de date saisis dans l'UI (session_state).
    """
    from datetime import date as _date_t
    result = []
    for i, group in enumerate(groups):
        g = deepcopy(group)
        for j, mc in enumerate(g.module_columns):
            key = _wkey("date_dj", i + group_offset, j)
            val = st.session_state.get(key)
            if isinstance(val, _date_t):
                mc.date = val.strftime("%d/%m/%Y")
        result.append(g)
    return result


def _get_dates_realisees(groups: list[ConventionGroup]) -> list[Optional[str]]:
    """
    Construit la liste des dates effectives de session par groupe,
    triées chronologiquement et formatées pour le certificat de réalisation.
    Ex: ["01/04/2026, 08/04/2026, 15/04/2026", None, "15/05/2026, 22/05/2026", ...]

    None si aucune date n'a été saisie pour ce groupe (fallback sur dates_previsionnelles).
    """
    from datetime import date as _date_t, datetime as _dt
    result = []
    for i, group in enumerate(groups):
        dates = []
        for j in range(len(group.module_columns)):
            val = st.session_state.get(_wkey("date_dj", i, j))
            if isinstance(val, _date_t):
                dates.append(val)
        if dates:
            dates_sorted = sorted(set(dates))
            result.append(", ".join(d.strftime("%d/%m/%Y") for d in dates_sorted))
        else:
            result.append(None)
    return result


def _do_generate_post(
    groups: list[ConventionGroup],
    client: ClientInfo,
    dates_prev: str,
    date_doc: str,
    ref: str,
    soffice: Optional[str],
) -> None:
    """Génère certificats + feuilles de présence et stocke le ZIP dans session_state."""
    try:
        certif_r   = CertificatRenderer()
        presence_r = FeuillePresenceRenderer()
    except FileNotFoundError as e:
        st.error(str(e))
        return

    # Dates effectives de session (saisies dans l'UI) — pour certificats et présences
    groups_with_dates     = _apply_dj_dates(groups)
    dates_realisees_list  = _get_dates_realisees(groups)

    with st.spinner("Génération des certificats et feuilles de présence..."):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                out = Path(tmp)
                certif_out   = out / "certificats"
                presence_out = out / "feuilles_presence"
                certif_out.mkdir(parents=True, exist_ok=True)
                presence_out.mkdir(parents=True, exist_ok=True)

                certif_paths = certif_r.generate_all(
                    groups=groups,
                    client=client,
                    dates_previsionnelles=dates_prev,
                    date_document=date_doc,
                    ref_dossier=ref,
                    output_dir=certif_out,
                    dates_realisees_per_group=dates_realisees_list,
                )
                presence_paths = presence_r.generate_all(
                    groups=groups_with_dates,
                    client=client,
                    dates_previsionnelles=dates_prev,
                    date_document=date_doc,
                    ref_dossier=ref,
                    output_dir=presence_out,
                )
                zip_bytes = _build_zip(
                    {
                        "certificats":       certif_paths,
                        "feuilles_presence": presence_paths,
                    },
                    soffice,
                )

            st.session_state["post_zip"]        = zip_bytes
            st.session_state["certif_count"]    = len(certif_paths)
            st.session_state["presence_count"]  = len(presence_paths)
            st.success(
                f"✅ {len(certif_paths)} certificat(s) + "
                f"{len(presence_paths)} feuille(s) de présence générés. "
                f"Cliquez sur **Télécharger** ci-dessous."
            )

        except Exception as exc:
            st.error(f"Erreur lors de la génération : {exc}")
            st.code(__import__("traceback").format_exc())


# ---------------------------------------------------------------------------
# Application principale
# ---------------------------------------------------------------------------

def main() -> None:

    # ── En-tête ────────────────────────────────────────────────────────────
    st.title("Générateur de documents Qualiopi")
    st.caption("eBazten · Koban CRM — Estelle Lecanu")

    # ── LibreOffice (détecté une seule fois par session) ───────────────────
    if "libreoffice" not in st.session_state:
        st.session_state["libreoffice"] = _find_libreoffice()

    soffice: Optional[str] = st.session_state["libreoffice"]
    if not soffice:
        st.info(
            "LibreOffice non détecté — les documents seront générés en .docx uniquement. "
            "Sur Streamlit Cloud (Linux), le PDF est automatiquement disponible."
        )

    # ── Étape 1 — Upload ────────────────────────────────────────────────────
    st.header("1. Fichier Excel")

    uploaded = st.file_uploader(
        "Tableau « Formations et inscrits »",
        type=["xlsx"],
        help="Fichier Excel d'Estelle : groupes de formation, modules et liste des stagiaires.",
    )

    if uploaded is None:
        st.info("Déposez votre fichier Excel pour commencer.")
        return

    # Détection d'un nouveau fichier → reset complet de la session
    if st.session_state.get("upload_name") != uploaded.name:
        old_count = st.session_state.get("upload_count", 0)
        soffice_saved = st.session_state.get("libreoffice")
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.session_state["upload_count"] = old_count + 1
        st.session_state["libreoffice"]  = soffice_saved
        soffice = soffice_saved

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded.read())
            tmp_path = Path(tmp.name)

        try:
            with st.spinner("Analyse du fichier Excel..."):
                excel_data = parse_formations_excel(tmp_path)
                mapper     = ModuleMapper()
                groups     = build_convention_groups(excel_data, mapper)
        except Exception as exc:
            st.error(f"Impossible de lire le fichier : {exc}")
            return

        st.session_state["upload_name"] = uploaded.name
        st.session_state["excel_data"]  = excel_data
        st.session_state["groups"]      = groups

    groups: list[ConventionGroup] = st.session_state.get("groups", [])
    excel_data = st.session_state.get("excel_data")

    if not groups:
        st.warning("Aucun groupe de formation détecté dans ce fichier.")
        return

    nb_stag  = sum(len(g.all_stagiaires) for g in groups)
    nb_conv  = len(groups)
    st.success(
        f"**{nb_conv} groupes détectés** — "
        f"{nb_stag} stagiaires — "
        f"Client : **{excel_data.client_name}**"
    )

    # ── Étape 2 — Informations client ───────────────────────────────────────
    st.header("2. Informations client")

    # Pré-remplir depuis les valeurs déjà saisies (si formulaire déjà validé)
    _saved = st.session_state.get("client")
    _default_ref = (
        st.session_state.get("ref")
        or f"{date_cls.today().year}-{excel_data.client_name}"
    )

    with st.form("client_form"):
        col1, col2 = st.columns(2)
        with col1:
            nom          = st.text_input(
                "Nom de la société *",
                value=_saved.nom if _saved else (excel_data.client_name or ""),
            )
            representant = st.text_input(
                "Représentant légal *",
                value=_saved.representant if _saved else "Gilles BERTRAND",
            )
            fonction     = st.text_input(
                "Fonction *",
                value=_saved.fonction if _saved else "Dirigeant",
            )
        with col2:
            adresse = st.text_area(
                "Adresse complète *", height=100,
                value=_saved.adresse if _saved else (
                    "PARC D'ACTIVITES DES ECLAPONS\n"
                    "3 CHEMIN DES ECLAPONS\n"
                    "69390 VOURLES"
                ),
            )
            frais = st.number_input(
                "Frais de mission HT (€)",
                min_value=0.0,
                value=_saved.frais_mission_ht if _saved else 0.0,
                step=50.0, format="%.2f",
            )

        st.divider()
        col3, col4 = st.columns(2)
        with col3:
            dates_prev = st.text_input(
                "Dates prévisionnelles *",
                value=st.session_state.get("dates_prev", ""),
                placeholder="Ex : Du 01/04/2026 au 31/10/2026",
            )
        with col4:
            date_doc = st.text_input(
                "Date du document",
                value=st.session_state.get("date_doc", date_cls.today().strftime("%d/%m/%Y")),
            )
        ref = st.text_input(
            "Référence dossier",
            value=_default_ref,
            placeholder="2026-GENERFEU",
        )

        submitted = st.form_submit_button("Valider", type="primary", use_container_width=True)

    if submitted:
        missing = [
            f for f, v in [
                ("Nom de la société", nom),
                ("Adresse", adresse),
                ("Représentant légal", representant),
                ("Fonction", fonction),
                ("Dates prévisionnelles", dates_prev),
            ] if not v.strip()
        ]
        if missing:
            st.error(f"Champs obligatoires manquants : {', '.join(missing)}")
        else:
            st.session_state["client"]     = ClientInfo(
                nom=nom, adresse=adresse,
                representant=representant, fonction=fonction,
                frais_mission_ht=frais,
            )
            st.session_state["dates_prev"] = dates_prev
            st.session_state["date_doc"]   = date_doc
            st.session_state["ref"]        = ref
            st.success("Informations enregistrées.")

    if "client" not in st.session_state:
        st.info("Renseignez et validez les informations client pour continuer.")
        return

    client:     ClientInfo = st.session_state["client"]
    dates_prev: str        = st.session_state["dates_prev"]
    date_doc:   str        = st.session_state["date_doc"]
    ref:        str        = st.session_state["ref"]

    # ── Étape 3 — Aperçu des groupes ────────────────────────────────────────
    st.header("3. Aperçu des groupes détectés")
    st.caption(
        "Vérifiez les groupes, ajustez les demi-journées facturées si nécessaire "
        "et décochez les stagiaires optionnels à exclure."
    )

    for i, group in enumerate(groups):
        _render_group(i, group)

    # Résumé financier global
    total_ht  = sum(g.nb_demi_journees * g.tarif_demi_journee_ht for g in groups)
    total_ttc = total_ht * 1.20
    st.info(
        f"**Total toutes conventions :** {total_ht:,.0f} € HT — "
        f"{total_ttc:,.0f} € TTC"
    )

    st.divider()

    # ── Phase A — Conventions ───────────────────────────────────────────────
    st.header("Phase A — Conventions")
    st.caption(
        f"À générer **avant la formation**. "
        f"{nb_conv} convention(s) — une par groupe de stagiaires."
    )

    if st.button(
        "Générer les conventions",
        type="primary",
        use_container_width=True,
        key="btn_conv",
    ):
        _do_generate_conventions(groups, client, dates_prev, date_doc, ref, soffice)

    if "convention_zip" in st.session_state:
        n    = st.session_state["convention_count"]
        fmt  = "(+ PDF)" if soffice else "(.docx uniquement)"
        name = f"conventions_{_slug(client.nom)}_{date_doc.replace('/', '-')}.zip"
        st.download_button(
            label=f"Télécharger — {n} convention(s) {fmt}",
            data=st.session_state["convention_zip"],
            file_name=name,
            mime="application/zip",
            use_container_width=True,
        )

    st.divider()

    # ── Phase B — Certificats + Feuilles de présence ───────────────────────
    nb_certifs  = sum(len(g.all_stagiaires) for g in groups)
    nb_feuilles = sum(len(g.module_columns) for g in groups)

    st.header("Phase B — Certificats + Feuilles de présence")
    st.caption(
        f"À générer **après la formation**. "
        f"~{nb_certifs} certificat(s) individuels + {nb_feuilles} feuille(s) de présence."
    )

    if st.button(
        "Générer certificats et feuilles de présence",
        type="primary",
        use_container_width=True,
        key="btn_post",
    ):
        _do_generate_post(groups, client, dates_prev, date_doc, ref, soffice)

    if "post_zip" in st.session_state:
        n_c  = st.session_state["certif_count"]
        n_f  = st.session_state["presence_count"]
        fmt  = "(+ PDF)" if soffice else "(.docx uniquement)"
        name = f"post_formation_{_slug(client.nom)}_{date_doc.replace('/', '-')}.zip"
        st.download_button(
            label=f"⬇️  Télécharger le ZIP — {n_c} certificat(s) + {n_f} feuille(s) {fmt}",
            data=st.session_state["post_zip"],
            file_name=name,
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )
        st.caption(
            f"Le ZIP contient deux dossiers : **certificats/** ({n_c} fichiers) "
            f"et **feuilles_presence/** ({n_f} fichiers)."
        )


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()
