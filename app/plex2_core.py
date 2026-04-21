#!/usr/bin/env python3
"""
PLEX² – moteur DOE Explorer
Basé sur le script source utilisateur DoE_Explorer_PhAm-ForApp.py.
Version Windows / GUI avec LHS automatique et résumé Excel ordonné.
"""

from __future__ import annotations

import argparse
import itertools
import math
import re
import sys
from collections import Counter, defaultdict
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import pyDOE3
except ImportError as exc:
    raise ImportError("pyDOE3 n'est pas installe. Installer avec : pip install pyDOE3") from exc


# ============================================================
# COULEURS PAR CATEGORIE
# ============================================================
CATEGORY_COLORS: dict[str, tuple[str, str]] = {
    "resume":       ("2E2E2E", "FFFFFF"),
    "complet":      ("1F4E79", "FFFFFF"),
    "fraction":     ("833C00", "FFFFFF"),
    "plackett":     ("7030A0", "FFFFFF"),
    "frac2L":       ("BF5F00", "FFFFFF"),
    "dsd":          ("0A6B6B", "FFFFFF"),
    "ofat":         ("3A3A3A", "FFFFFF"),
    "taguchi":      ("375623", "FFFFFF"),
    "gsd":          ("1D6B55", "FFFFFF"),
    "ccd":          ("006B6B", "FFFFFF"),
    "ccd_coded":    ("4BACC6", "FFFFFF"),
    "box-behnken":  ("C00000", "FFFFFF"),
    "bb_coded":     ("FF7B7B", "000000"),
    "doehlert":     ("4A235A", "FFFFFF"),
    "doehlert_cod": ("9B59B6", "FFFFFF"),
    "lhs":          ("7F6000", "FFFFFF"),
}


def _category(type_label: str) -> str:
    t = type_label.lower()
    if "resume" in t or "summary" in t:
        return "resume"
    if "min/max" in t or "reformat" in t:
        return "frac2L"
    if "dsd" in t or "definitive screening" in t:
        return "dsd"
    if "ofat" in t or "one-factor" in t:
        return "ofat"
    if "complet" in t:
        return "complet"
    if "fraction" in t or "frac" in t:
        return "fraction"
    if "plackett" in t:
        return "plackett"
    if "taguchi" in t:
        return "taguchi"
    if "gsd" in t:
        return "gsd"
    if "ccd" in t or "composite" in t:
        return "ccd_coded" if "coded" in t else "ccd"
    if "box" in t:
        return "bb_coded" if "coded" in t else "box-behnken"
    if "doehlert" in t:
        return "doehlert_cod" if "coded" in t else "doehlert"
    if "lhs" in t or "hypercube" in t:
        return "lhs"
    return "resume"


# ============================================================
# HELPERS NOMS DE FEUILLE
# ============================================================
_INVALID = re.compile(r"[\\/:*?\[\]]")


def _safe_name(name: str, used: set[str]) -> str:
    name = _INVALID.sub("_", name).strip()[:31] or "Sheet"
    base, i = name, 2
    while name in used:
        suffix = f"_{i}"
        name = (base[: 31 - len(suffix)] + suffix)[:31]
        i += 1
    used.add(name)
    return name


# ============================================================
# CONVERSIONS CODE -> VALEURS REELLES
# ============================================================

def _df_full_factorial(factors: Dict[str, List[Any]]) -> pd.DataFrame:
    names = list(factors.keys())
    grid = list(itertools.product(*(factors[n] for n in names)))
    df = pd.DataFrame(grid, columns=names)
    df.insert(0, "Run", range(1, len(df) + 1))
    return df



def _df_from_0indexed(mat: np.ndarray, factors: Dict[str, List[Any]]) -> pd.DataFrame:
    names = list(factors.keys())
    data = {n: [factors[n][int(c)] for c in mat[:, j]] for j, n in enumerate(names)}
    df = pd.DataFrame(data)
    df.insert(0, "Run", range(1, len(df) + 1))
    return df



def _df_from_pm1_2level(mat: np.ndarray, factors: Dict[str, List[Any]]) -> pd.DataFrame:
    names = list(factors.keys())
    data = {
        n: [factors[n][0] if v < 0 else factors[n][1] for v in mat[:, j]]
        for j, n in enumerate(names)
    }
    df = pd.DataFrame(data)
    df.insert(0, "Run", range(1, len(df) + 1))
    return df



def _df_from_rsm(mat: np.ndarray, factors: Dict[str, List[Any]]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    names = list(factors.keys())
    df_coded = pd.DataFrame(mat, columns=[f"{n}_coded" for n in names])
    df_coded.insert(0, "Run", range(1, len(df_coded) + 1))

    def _is_num(lvls: list[Any]) -> bool:
        return all(isinstance(x, (int, float, np.number)) for x in lvls)

    if not all(_is_num(factors[n]) for n in names):
        return df_coded, df_coded

    mins = np.array([min(factors[n]) for n in names], dtype=float)
    maxs = np.array([max(factors[n]) for n in names], dtype=float)
    centers = (mins + maxs) / 2.0
    scales = (maxs - mins) / 2.0
    scales[scales == 0] = 1.0

    real = centers + mat * scales
    df_real = pd.DataFrame(real, columns=names)
    df_real.insert(0, "Run", range(1, len(df_real) + 1))
    return df_real, df_coded


# ============================================================
# HELPERS DSD / LHS / TRI
# ============================================================

def _is_numeric_factor(lvls: List[Any]) -> bool:
    return all(isinstance(v, (int, float, np.number)) for v in lvls)



def _hadamard_sylvester(n: int) -> np.ndarray:
    H = np.array([[1.0]])
    while H.shape[0] < n:
        H = np.block([[H, H], [H, -H]])
    return H



def _dsd_matrix(k: int) -> np.ndarray:
    n = 1
    while n <= k:
        n *= 2
    H = _hadamard_sylvester(n)
    sub = H[1 : k + 1, 1 : k + 1].copy()
    np.fill_diagonal(sub, 0.0)
    return np.vstack([sub, -sub, np.zeros((1, k))])



def _compute_auto_lhs_samples(factors: Dict[str, List[Any]]) -> int:
    """Calcule une taille LHS bornée pour éviter les volumes excessifs."""
    numeric_profile = [len(v) for v in factors.values() if _is_numeric_factor(v)]
    if not numeric_profile:
        return 0

    k_num = len(numeric_profile)
    avg_l = sum(numeric_profile) / k_num
    max_l = max(numeric_profile)

    full_runs = 1
    for n in numeric_profile:
        full_runs *= n
        if full_runs >= 100_000:
            full_runs = 100_000
            break

    base = max(12, 4 * k_num)
    coverage = math.ceil(1.5 * k_num * avg_l)
    spread = math.ceil(2.25 * max_l * math.sqrt(k_num))
    root_full = math.ceil(math.sqrt(full_runs))
    ceiling = max(24, min(300, 20 * k_num))

    candidate = max(base, coverage, spread, root_full)
    candidate = min(candidate, ceiling)
    if candidate % 2:
        candidate += 1
    return max(12, candidate)



def _normalize_warning_note(type_label: str, notes: str, is_generated: bool) -> str:
    text = (notes or '').strip()
    lower = f"{type_label} {text}".lower()
    if not is_generated:
        if not text.lower().startswith("⚠ omis"):
            text = f"⚠ OMIS – {text.removeprefix('OMIS – ').strip()}" if text else "⚠ OMIS"
        return text
    if "latin hypercube" in lower and not text.lower().startswith("⚠ lhs"):
        text = f"⚠ LHS : {text}" if text else "⚠ LHS"
    elif ("reformat" in lower or "min/max" in lower or "exclus" in lower) and not text.startswith("⚠"):
        text = f"⚠ REFORMATAGE : {text}" if text else "⚠ REFORMATAGE"
    return text



def _classify_priority(type_label: str, notes: str, is_generated: bool) -> int:
    lower = f"{type_label} {notes}".lower()
    if not is_generated or lower.startswith("⚠ omis") or "⚠ omis" in lower:
        return 2
    if "⚠" in notes or "reformat" in lower or "min/max" in lower or " latin hypercube" in lower:
        return 1
    return 0



def _validate_plan_df(df: pd.DataFrame, type_label: str) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Le plan '{type_label}' n'a pas produit de DataFrame pandas valide.")
    if df.empty:
        raise ValueError(f"Le plan '{type_label}' a produit un DataFrame vide.")
    if df.columns.duplicated().any():
        raise ValueError(f"Le plan '{type_label}' contient des colonnes dupliquées.")
    return df.copy()



def _build_short_label(seed: str) -> str:
    seed = seed.upper()
    seed = seed.replace(" ", "_")
    seed = seed.replace("-", "_")
    seed = re.sub(r"[^A-Z0-9_]+", "", seed)
    seed = re.sub(r"_+", "_", seed).strip("_")
    return seed[:24] or "PLAN"


# ============================================================
# TAGUCHI : AUTO-DECOUVERTE COMPLETE
# ============================================================

def _taguchi_all_compatible(factors: Dict[str, List[Any]]) -> List[Tuple[str, pd.DataFrame]]:
    try:
        from pyDOE3 import get_orthogonal_array, list_orthogonal_arrays
    except ImportError:
        return []

    needed = Counter(len(lvls) for lvls in factors.values())
    factors_by_level: dict[int, list[str]] = defaultdict(list)
    for name, lvls in factors.items():
        factors_by_level[len(lvls)].append(name)

    results: List[Tuple[str, pd.DataFrame]] = []
    for oa_name in list_orthogonal_arrays():
        try:
            oa = get_orthogonal_array(oa_name)
        except Exception:
            continue

        cols_by_level: dict[int, list[int]] = defaultdict(list)
        for j in range(oa.shape[1]):
            cols_by_level[len(np.unique(oa[:, j]))].append(j)

        if not all(len(cols_by_level.get(lvl, [])) >= req for lvl, req in needed.items()):
            continue

        data: dict[str, list[Any]] = {}
        for lvl, fnames in factors_by_level.items():
            for fname, col_idx in zip(fnames, cols_by_level[lvl]):
                lvls = factors[fname]
                data[fname] = [lvls[int(c)] for c in oa[:, col_idx].astype(int)]

        df = pd.DataFrame(data)
        df.insert(0, "Run", range(1, len(df) + 1))
        results.append((oa_name, df))

    results.sort(key=lambda x: (len(x[1]), x[0]))
    return results


# ============================================================
# FORMATAGE EXCEL
# ============================================================

def _format_workbook(outfile: str, sheet_meta: dict[str, str]) -> None:
    wb = load_workbook(outfile)

    row_even = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    row_odd = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center")
    thin = Border(
        left=Side(style="thin", color="BBBBBB"),
        right=Side(style="thin", color="BBBBBB"),
        top=Side(style="thin", color="BBBBBB"),
        bottom=Side(style="thin", color="BBBBBB"),
    )

    def _fmt(ws, hdr_bg: str, hdr_fg: str, wide_last: bool = False) -> None:
        hdr_fill = PatternFill("solid", start_color=hdr_bg, end_color=hdr_bg)
        hdr_font = Font(name="Arial", bold=True, color=hdr_fg, size=10)
        ws.row_dimensions[1].height = 22
        for cell in ws[1]:
            if cell.value is not None:
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = center
                cell.border = thin
        for ri, row in enumerate(ws.iter_rows(min_row=2), start=2):
            fill = row_even if ri % 2 == 0 else row_odd
            for cell in row:
                cell.font = data_font
                cell.fill = fill
                cell.alignment = center
                cell.border = thin
        for col in ws.columns:
            mw = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(mw + 3, 12)
        if wide_last and ws.max_column > 1:
            lc = get_column_letter(ws.max_column)
            ws.column_dimensions[lc].width = max(ws.column_dimensions[lc].width, 60)
        ws.freeze_panes = "B2"

    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for sname in wb.sheetnames:
        ws = wb[sname]
        tl = sheet_meta.get(sname, "")
        cat = _category(tl)
        bg, fg = CATEGORY_COLORS.get(cat, ("808080", "FFFFFF"))
        _fmt(ws, bg, fg, wide_last=(cat == "resume"))
        ws.sheet_properties.tabColor = bg

        if sname == "00_RESUME":
            notes_col = None
            for cell in ws[1]:
                if cell.value is not None and str(cell.value).strip().lower() == "notes":
                    notes_col = cell.column
                    break
            if notes_col is not None:
                for row in ws.iter_rows(min_row=2, min_col=notes_col, max_col=notes_col):
                    for cell in row:
                        cell.alignment = left_align

    wb.save(outfile)


# ============================================================
# FONCTION PRINCIPALE
# ============================================================

def build_doe_explorer(
    factors: Dict[str, List[Any]],
    outfile: str = "PLEX2_plans.xlsx",
    max_full_factorial_runs: int = 30_000,
    gsd_reductions: Tuple[int, ...] = (2, 3, 4, 5),
    include_lhs: bool = True,
    lhs_samples: Optional[int] = None,
) -> list[dict[str, Any]]:
    if not factors:
        raise ValueError("Aucun facteur fourni.")
    for name, lvls in factors.items():
        if not isinstance(lvls, (list, tuple)) or len(lvls) < 2:
            raise ValueError(f"Facteur '{name}' : il faut au moins 2 niveaux.")

    names = list(factors.keys())
    k = len(names)
    level_profile = [len(factors[n]) for n in names]
    all_2level = all(L == 2 for L in level_profile)

    plans: list[dict[str, Any]] = []
    omitted: list[dict[str, Any]] = []

    def add(short_seed: str, df: pd.DataFrame, type_label: str, notes: str = "") -> None:
        safe_df = _validate_plan_df(df, type_label)
        safe_notes = _normalize_warning_note(type_label, notes, True)
        plans.append(
            {
                "short": _build_short_label(short_seed),
                "type": type_label,
                "essais": len(safe_df),
                "colonnes": len(safe_df.columns) - 1,
                "notes": safe_notes,
                "df": safe_df,
                "priority": _classify_priority(type_label, safe_notes, True),
                "order": len(plans),
            }
        )

    def skip(short_seed: str, type_label: str, reason: str, runs: int = 0) -> None:
        notes = _normalize_warning_note(type_label, str(reason), False)
        omitted.append(
            {
                "short": _build_short_label(short_seed),
                "type": type_label,
                "essais": runs,
                "colonnes": k,
                "notes": notes,
                "df": None,
                "priority": 2,
                "order": len(omitted),
            }
        )

    # 1. Factoriel complet
    n_full = 1
    for L in level_profile:
        n_full *= L

    if n_full <= max_full_factorial_runs:
        add(
            "FACT_COMPLET",
            _df_full_factorial(factors),
            "Factoriel Complet",
            f"{'x'.join(str(L) for L in level_profile)} = {n_full} essais",
        )
    else:
        skip("FACT_COMPLET", "Factoriel Complet", f"{n_full} essais > seuil {max_full_factorial_runs:,}", n_full)

    if not all_2level:
        factors_2L: Dict[str, List[Any]] = {n: [lvls[0], lvls[-1]] for n, lvls in factors.items()}
        reduced_names = [n for n in names if len(factors[n]) > 2]
        if reduced_names:
            reformat_label = (
                f"⚠ REFORMATAGE min/max : {len(reduced_names)} facteur(s) réduit(s) à 2 niveaux "
                f"({', '.join(reduced_names)}). Les niveaux intermédiaires ne sont pas utilisés dans ce plan."
            )
        else:
            reformat_label = "⚠ REFORMATAGE min/max : niveaux extrêmes utilisés."
    else:
        factors_2L = factors.copy()
        reformat_label = ""

    # 1b. Factoriel 2^k reformaté
    if not all_2level:
        n_full_2L = 2 ** k
        if n_full_2L <= max_full_factorial_runs:
            add(
                "FACT_2L_MM",
                _df_full_factorial(factors_2L),
                "Factoriel 2^k min/max reformaté",
                f"2^{k} = {n_full_2L} essais. {reformat_label}",
            )
        else:
            skip(
                "FACT_2L_MM",
                "Factoriel 2^k min/max reformaté",
                f"2^{k} = {n_full_2L} > seuil {max_full_factorial_runs:,}. {reformat_label}",
                n_full_2L,
            )

    # 2. Fractionnaires 2 niveaux
    if all_2level:
        try:
            from pyDOE3 import fracfact_by_res

            max_res = min(k + 1, 10)
            seen_sizes: set[int] = set()
            for res in range(max_res, 2, -1):
                try:
                    mat = fracfact_by_res(k, res)
                except Exception:
                    continue
                n = len(mat)
                if n in seen_sizes:
                    continue
                seen_sizes.add(n)
                add(
                    f"FRAC_RES{res}",
                    _df_from_pm1_2level(mat, factors),
                    f"Fractionnaire 2^({k}-p) Res.{res}",
                    f"Resolution {res}, {n} essais",
                )
        except ImportError:
            skip("FRAC_2L", "Fractionnaire 2-niveaux", "fracfact_by_res non disponible (pyDOE3 < 1.2)")

    # 2b. Fractionnaires reformatés
    if not all_2level:
        try:
            from pyDOE3 import fracfact_by_res as _fbr_mm

            seen_2L: set[int] = set()
            for res in range(min(k + 1, 10), 2, -1):
                try:
                    mat_fr = _fbr_mm(k, res)
                except Exception:
                    continue
                n_fr = len(mat_fr)
                if n_fr in seen_2L:
                    continue
                seen_2L.add(n_fr)
                add(
                    f"FRAC2L_RES{res}",
                    _df_from_pm1_2level(mat_fr, factors_2L),
                    f"Fractionnaire 2^(k-p) Res.{res} min/max reformaté",
                    f"Résolution {res}, {n_fr} essais. {reformat_label}",
                )
        except ImportError:
            skip(
                "FRAC2L",
                "Fractionnaire 2-niveaux min/max reformaté",
                "fracfact_by_res non disponible (pyDOE3 < 1.2)",
            )

    # 3. PB
    if all_2level:
        try:
            from pyDOE3 import pbdesign

            mat = pbdesign(k)
            add("PLACKETT_BURMAN", _df_from_pm1_2level(mat, factors), "Plackett-Burman", f"Criblage 2-niveaux, {len(mat)} essais")
        except Exception as e:
            skip("PLACKETT_BURMAN", "Plackett-Burman", str(e))

    # 3b. PB reformaté
    if not all_2level:
        try:
            from pyDOE3 import pbdesign as _pb_mm

            mat_pb = _pb_mm(k)
            add(
                "PB_MINMAX",
                _df_from_pm1_2level(mat_pb, factors_2L),
                "Plackett-Burman min/max reformaté",
                f"Criblage 2-niveaux, {len(mat_pb)} essais. {reformat_label}",
            )
        except Exception as e:
            skip("PB_MINMAX", "Plackett-Burman min/max reformaté", str(e))

    # 4. DSD
    quant_names = [n for n in names if _is_numeric_factor(factors[n])]
    cat_names = [n for n in names if not _is_numeric_factor(factors[n])]
    k_quant = len(quant_names)

    if k_quant >= 2:
        try:
            dsd_mat = _dsd_matrix(k_quant)
            dsd_data: Dict[str, list[Any]] = {}
            for j, n in enumerate(quant_names):
                lvls_q = factors[n]
                lo = float(min(lvls_q))
                hi = float(max(lvls_q))
                ctr = (lo + hi) / 2.0
                col_vals: list[Any] = []
                for v in dsd_mat[:, j]:
                    if abs(v) < 0.5:
                        col_vals.append(ctr)
                    elif v > 0:
                        col_vals.append(hi)
                    else:
                        col_vals.append(lo)
                dsd_data[n] = col_vals

            df_dsd = pd.DataFrame(dsd_data)
            df_dsd.insert(0, "Run", range(1, len(df_dsd) + 1))

            dsd_note = (
                f"⚠ REFORMATAGE : 3 niveaux (min, (min+max)/2, max) dérivés des valeurs min et max de chaque facteur. "
                f"2k+1 = {len(df_dsd)} essais. Construction Hadamard–Sylvester (approx. DSD). "
            )
            if cat_names:
                dsd_note += f"Facteurs non-numériques exclus : {', '.join(cat_names)} ({len(cat_names)}/{k} facteur(s))."
            if k_quant != k:
                dsd_note += f" Plan partiel sur {k_quant}/{k} facteurs quantitatifs."

            add("DSD", df_dsd, "Definitive Screening Design (DSD)", dsd_note)
        except Exception as e:
            skip("DSD", "DSD", str(e))
    else:
        skip("DSD", "DSD", f"k_quant = {k_quant} < 2 – DSD requiert ≥ 2 facteurs numériques")

    # 5. OFAT
    base_run = {n: factors[n][0] for n in names}
    ofat_runs: list[dict[str, Any]] = [base_run.copy()]
    for n_ofat in names:
        for lvl_ofat in factors[n_ofat][1:]:
            run_ofat = base_run.copy()
            run_ofat[n_ofat] = lvl_ofat
            ofat_runs.append(run_ofat)

    df_ofat = pd.DataFrame(ofat_runs)
    df_ofat.insert(0, "Run", range(1, len(df_ofat) + 1))
    n_ofat_total = len(df_ofat)
    lsum = " + ".join(f"({len(factors[n])}-1)" for n in names)
    add(
        "OFAT",
        df_ofat,
        "OFAT (One-Factor-At-a-Time)",
        f"Plan de référence. Essais = 1 + [{lsum}] = {n_ofat_total}. Base = 1er niveau de chaque facteur. Chaque facteur varie un par un, les autres restent au niveau de base.",
    )

    # 6. Taguchi
    taguchi_list = _taguchi_all_compatible(factors)
    if taguchi_list:
        for oa_name, df_oa in taguchi_list:
            short = re.sub(r"[ ()^*]", "", oa_name)
            add(f"TAGUCHI_{short}", df_oa, f"Taguchi {oa_name}", f"{oa_name} — {len(df_oa)} essais, {k} facteurs extraits")
    else:
        skip("TAGUCHI", "Taguchi", "Aucun OA compatible pour ce profil de niveaux")

    # 7. GSD
    try:
        from pyDOE3 import gsd

        for red in gsd_reductions:
            if red <= 1:
                continue
            try:
                mat = gsd(level_profile, red)
                add(
                    f"GSD_RED{red}",
                    _df_from_0indexed(mat, factors),
                    f"GSD reduction={red}",
                    f"Criblage multi-niveaux, reduction={red}, {len(mat)} essais",
                )
            except Exception as e:
                skip(f"GSD_RED{red}", f"GSD red={red}", str(e))
    except ImportError:
        skip("GSD", "GSD", "gsd() non disponible dans cette version de pyDOE3")

    # 8. CCD / 9. Box-Behnken / 10. Doehlert
    numeric_factors = {n: factors[n] for n in quant_names}
    numeric_note = ""
    if cat_names:
        numeric_note = f" ⚠ REFORMATAGE : facteurs non-numériques exclus : {', '.join(cat_names)}."

    if k_quant >= 2:
        try:
            from pyDOE3 import ccdesign

            for face, label in [("ccc", "Circonscrit"), ("ccf", "Face-centre"), ("cci", "Inscrit")]:
                try:
                    mat = np.array(ccdesign(k_quant, center=(1, 1), face=face))
                    df_real, _ = _df_from_rsm(mat, numeric_factors)
                    add(
                        f"CCD_{face.upper()}",
                        df_real,
                        f"CCD {label} (valeurs reelles)",
                        f"face={face}, {len(df_real)} essais.{numeric_note}",
                    )
                except Exception as e:
                    skip(f"CCD_{face.upper()}", f"CCD {face}", str(e))
        except ImportError:
            skip("CCD", "CCD", "ccdesign() non disponible")
    else:
        skip("CCD", "CCD", "requiert au moins 2 facteurs numériques")

    if k_quant >= 3:
        try:
            from pyDOE3 import bbdesign

            mat = np.array(bbdesign(k_quant))
            df_real, _ = _df_from_rsm(mat, numeric_factors)
            add(
                "BOX_BEHNKEN",
                df_real,
                "Box-Behnken (valeurs reelles)",
                f"{len(df_real)} essais, surface de reponse sans coin.{numeric_note}",
            )
        except Exception as e:
            skip("BOX_BEHNKEN", "Box-Behnken", str(e))
    else:
        skip("BOX_BEHNKEN", "Box-Behnken", "requiert au moins 3 facteurs numériques")

    if k_quant >= 2:
        doehlert_tasks = [
            ("doehlert_shell_design", "DOEHLERT_SHELL", "Doehlert Shell", {"num_center_points": 1}),
            ("doehlert_simplex_design", "DOEHLERT_SIMPLEX", "Doehlert Simplex", {}),
        ]
        for func_name, short_label, label, kw in doehlert_tasks:
            try:
                func = getattr(pyDOE3, func_name, None)
                if func is None:
                    exec(f"from pyDOE3 import {func_name} as func")
                mat = np.array(func(k_quant, **kw))
                df_real, _ = _df_from_rsm(mat, numeric_factors)
                add(short_label, df_real, f"{label} (valeurs reelles)", f"{len(df_real)} essais, domaine spherique.{numeric_note}")
            except Exception as e:
                skip(short_label, label, str(e))
    else:
        skip("DOEHLERT", "Doehlert", "requiert au moins 2 facteurs numériques")

    # 11. LHS automatique et systématique
    if include_lhs:
        numeric_names = [n for n in names if _is_numeric_factor(factors[n])]
        excluded_names = [n for n in names if n not in numeric_names]
        if numeric_names:
            try:
                from pyDOE3 import lhs

                n_lhs = lhs_samples if lhs_samples else _compute_auto_lhs_samples(factors)
                mat = lhs(len(numeric_names), samples=n_lhs, criterion="center")
                data_lhs: dict[str, list[float]] = {}
                for j, name in enumerate(numeric_names):
                    lvls = factors[name]
                    lo, hi = float(min(lvls)), float(max(lvls))
                    data_lhs[name] = [round(lo + v * (hi - lo), 6) for v in mat[:, j]]
                df_lhs = pd.DataFrame(data_lhs)
                df_lhs.insert(0, "Run", range(1, len(df_lhs) + 1))
                note = f"⚠ LHS : criterion=center, {n_lhs} points (calcul automatique)."
                if excluded_names:
                    note += f" ⚠ REFORMATAGE : facteurs non-numériques exclus : {', '.join(excluded_names)}."
                add("LHS_AUTO", df_lhs, "Latin Hypercube Sampling", note)
            except Exception as e:
                skip("LHS_AUTO", "LHS", str(e))
        else:
            skip("LHS_AUTO", "LHS", "aucun facteur numérique disponible pour le Latin Hypercube")

    # Tri final : nominaux -> reformatage -> omis
    generated_sorted = sorted(plans, key=lambda item: (item["priority"], item["order"], item["type"]))
    omitted_sorted = sorted(omitted, key=lambda item: (item["priority"], item["order"], item["type"]))

    used_names: set[str] = {"00_RESUME"}
    sheets: list[tuple[str, pd.DataFrame, str]] = []
    summary_rows: list[dict[str, Any]] = []
    sheet_meta: dict[str, str] = {"00_RESUME": "Resume"}

    for idx, item in enumerate(generated_sorted, start=1):
        sheet_name = _safe_name(f"{idx:02d}_{item['short']}", used_names)
        item["sheet"] = sheet_name
        sheets.append((sheet_name, item["df"], item["type"]))
        sheet_meta[sheet_name] = item["type"]
        summary_rows.append(
            {
                "Feuille": sheet_name,
                "Type": item["type"],
                "Essais": item["essais"],
                "Colonnes": item["colonnes"],
                "Notes": item["notes"],
            }
        )

    for item in omitted_sorted:
        summary_rows.append(
            {
                "Feuille": "—",
                "Type": item["type"],
                "Essais": item["essais"],
                "Colonnes": item["colonnes"],
                "Notes": item["notes"],
            }
        )

    df_resume = pd.DataFrame(summary_rows)
    with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
        df_resume.to_excel(writer, sheet_name="00_RESUME", index=False)
        for sname, df, _ in sheets:
            df.to_excel(writer, sheet_name=sname, index=False)

    _format_workbook(outfile, sheet_meta)
    return summary_rows


# ============================================================
# CLI (conservé pour debug interne)
# ============================================================

def _cli() -> None:
    p = argparse.ArgumentParser(description="PLEX² – plans exhaustifs dans un fichier Excel")
    p.add_argument("--k", type=int, required=True, help="Nombre de facteurs")
    p.add_argument("--l", type=int, required=True, help="Niveaux (uniforme)")
    p.add_argument("--out", default="PLEX2_plans.xlsx", help="Fichier de sortie")
    p.add_argument("--max-full", type=int, default=30_000, help="Seuil max pour le factoriel complet")
    args = p.parse_args()

    import string

    names = (list(string.ascii_uppercase) + [f"X{i}" for i in range(1, 100)])[: args.k]
    factors = {n: list(range(1, args.l + 1)) for n in names}
    build_doe_explorer(factors, args.out, max_full_factorial_runs=args.max_full, include_lhs=True)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        _cli()
    else:
        factors_level = {
            "Matériau": ["Acier", "Alu"],
            "Température": [20.0, 120.0],
            "Pression": [1.0, 2.0, 3.0],
            "Lubrification": ["Oui", "Non"],
            "Epaisseur": [0.4, 0.6],
            "Dureté": [40.0, 60.0, 70.0],
            "Ra": [0.12, 0.8, 1.6],
        }
        build_doe_explorer(factors_level, outfile="PLEX2_demo.xlsx", max_full_factorial_runs=30_000, include_lhs=True)
