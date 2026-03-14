"""
normalize_names.py – Normalise Nom (MAJUSCULES) et Prénom (Titre) dans MOF_France_FUSIONNÉ.xlsx
  - Nom    → TOUT EN MAJUSCULES (ex: "dupont" → "DUPONT")
  - Prénom → Titre avec gestion des tirets (ex: "jean-philippe" → "Jean-Philippe")
  - Nom complet → rebuild "NOM Prénom"
"""

import pandas as pd
import unicodedata
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

EXCEL_PATH = Path("/Users/nicolas/projets_mof/output/MOF_France_FUSIONNÉ.xlsx")


def title_prenom(s: str) -> str:
    """Met chaque mot en majuscule initiale, gère tirets et apostrophes.
    Ex: 'jean-philippe' → 'Jean-Philippe', 'de bourgogne' → 'De Bourgogne'
    """
    if not s or not isinstance(s, str):
        return s
    s = s.strip()
    # Séparer sur tiret ET espace, recapitaliser chaque partie
    parts = []
    for part in s.replace("-", "—").split():
        sub = part.split("—")
        parts.append("-".join(p.capitalize() for p in sub))
    return " ".join(parts)


def upper_nom(s: str) -> str:
    """Met un nom de famille en MAJUSCULES."""
    if not s or not isinstance(s, str):
        return s
    return s.strip().upper()


def rebuild_nom_complet(nom: str, prenom: str) -> str:
    n = str(nom).strip() if pd.notna(nom) and str(nom).strip() else ""
    p = str(prenom).strip() if pd.notna(prenom) and str(prenom).strip() else ""
    if n and p:
        return f"{n} {p}"
    return n or p


def main():
    print(f"Chargement : {EXCEL_PATH}")
    df = pd.read_excel(EXCEL_PATH)
    n = len(df)
    print(f"  {n} lignes")

    # Avant
    print(f"\nExemples AVANT:")
    sample_idx = df[df['Nom'].notna() & df['Prénom'].notna()].index[:5]
    print(df.loc[sample_idx, ['Nom', 'Prénom', 'Nom complet']].to_string())

    # Normalisation
    df['Nom']    = df['Nom'].apply(lambda x: upper_nom(x) if pd.notna(x) else x)
    df['Prénom'] = df['Prénom'].apply(lambda x: title_prenom(x) if pd.notna(x) else x)

    # Rebuild Nom complet
    df['Nom complet'] = df.apply(
        lambda r: rebuild_nom_complet(r['Nom'], r['Prénom']), axis=1
    )

    # Après
    print(f"\nExemples APRÈS:")
    print(df.loc[sample_idx, ['Nom', 'Prénom', 'Nom complet']].to_string())

    # Vérifier quelques cas problématiques
    print("\nVérif Jean-Philippe :")
    jp = df[df['Prénom'].str.contains('jean|Jean|philippe|Philippe', case=False, na=False)].head(5)
    print(jp[['Nom', 'Prénom', 'Nom complet']].to_string())

    # ----- Sauvegarde Excel -----
    print(f"\nSauvegarde → {EXCEL_PATH}")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MOFs"

    header_fill = PatternFill(start_color="1A2E5A", end_color="1A2E5A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    cols = list(df.columns)
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for ri, row_data in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val if pd.notna(val) else None)

    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Stats métiers
    ws2 = wb.create_sheet("Stats Métiers")
    stats_m = df.groupby("Métier").size().sort_values(ascending=False).reset_index(name="Nb")
    ws2.append(["Métier", "Nb"])
    for _, r in stats_m.iterrows():
        ws2.append([r["Métier"], r["Nb"]])

    # Stats départements
    ws3 = wb.create_sheet("Stats Départements")
    stats_d = df.groupby(["Num_Dept", "Département", "Région"]).size().sort_values(ascending=False).reset_index(name="Nb")
    ws3.append(["Num_Dept", "Département", "Région", "Nb"])
    for _, r in stats_d.iterrows():
        ws3.append([r["Num_Dept"], r["Département"], r["Région"], r["Nb"]])

    wb.save(EXCEL_PATH)
    print(f"  ✓ Sauvegardé ({n} MOFs)")


if __name__ == "__main__":
    main()
