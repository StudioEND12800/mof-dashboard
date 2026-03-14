"""
merge_complet.py  –  Fusionne MOF_France_FUSIONNÉ.xlsx avec MOF_France_complet.xlsx
Actions :
  1. Ajoute colonnes contact (Téléphone, Email, Site web, Adresse, Code postal, URL source)
     en enrichissant les lignes déjà présentes dans FUSIONNÉ
  2. Corrige doublons "Prénom manquant" (même Nom + Métier + Promo, un Prénom vide)
  3. Ajoute les lignes genuinement nouvelles depuis COMPLET
  4. Re-normalise Nom/Prénom/Nom complet
  5. Sauvegarde
"""

import pandas as pd
import unicodedata
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

FUSE_PATH   = Path("/Users/nicolas/projets_mof/output/MOF_France_FUSIONNÉ.xlsx")
COMPL_PATH  = Path("/Users/nicolas/projets_mof/output/MOF_France_complet.xlsx")

# ── helpers ──────────────────────────────────────────────────────────────────
def norm(s) -> str:
    if not s or isinstance(s, float):
        return ""
    return "".join(c for c in unicodedata.normalize("NFD", str(s).strip().upper())
                   if unicodedata.category(c) != "Mn")

def upper_nom(s) -> str:
    return str(s).strip().upper() if pd.notna(s) and str(s).strip() else s

def title_prenom(s) -> str:
    if not s or not isinstance(s, str) or not str(s).strip():
        return s
    parts = []
    for part in str(s).strip().replace("-", "—").split():
        sub = part.split("—")
        parts.append("-".join(p.capitalize() for p in sub))
    return " ".join(parts)

def rebuild_nc(nom, prenom) -> str:
    n = str(nom).strip()   if pd.notna(nom)    and str(nom).strip()    else ""
    p = str(prenom).strip() if pd.notna(prenom) and str(prenom).strip() else ""
    if n and p:
        return f"{n} {p}"
    return n or p

def count_fill(row) -> int:
    return sum(1 for v in row if pd.notna(v) and str(v).strip() != "")

# ── Clé large : Nom seul + Métier + Promo (pour cas Prénom manquant) ─────────
def broad_key(nom, metier, promo) -> tuple:
    promo_str = ""
    if pd.notna(promo) and str(promo).strip() not in ("", "nan"):
        try:
            promo_str = str(int(float(str(promo).strip())))
        except (ValueError, TypeError):
            promo_str = ""
    return (norm(nom), norm(str(metier)) if pd.notna(metier) else "", promo_str)

# ── Clé stricte : sorted(Nom, Prénom) + Métier ───────────────────────────────
def strict_key(nom, prenom, metier) -> tuple:
    return (tuple(sorted([norm(nom), norm(prenom)])),
            norm(str(metier)) if pd.notna(metier) else "")

# ═══════════════════════════════════════════════════════════════════════════════
def main():
    print("Chargement des fichiers…")
    fuse  = pd.read_excel(FUSE_PATH)
    compl = pd.read_excel(COMPL_PATH)
    print(f"  FUSIONNÉ : {len(fuse)} lignes")
    print(f"  COMPLET  : {len(compl)} lignes")

    # ── 1. Ajouter colonnes contact à FUSIONNÉ ────────────────────────────────
    new_cols = ["Téléphone", "Email", "Site web", "Adresse", "Code postal", "URL source"]
    for col in new_cols:
        if col not in fuse.columns:
            fuse[col] = pd.NA

    # Renommer 'Entreprise / Établissement' → 'Entreprise' dans COMPLET
    compl = compl.rename(columns={"Entreprise / Établissement": "Entreprise"})

    # ── 2. Construire index de COMPLET (strict + large) ───────────────────────
    compl["_skey"] = compl.apply(lambda r: strict_key(r["Nom"], r["Prénom"], r["Métier"]), axis=1)
    compl["_bkey"] = compl.apply(lambda r: broad_key(r["Nom"], r["Métier"], r["Promotion"]), axis=1)

    compl_skey = {row["_skey"]: row for _, row in compl.iterrows()}
    compl_bkey = {row["_bkey"]: row for _, row in compl.iterrows()
                  if row["_bkey"][2] != ""}  # seulement si promo renseignée

    # ── 3. Enrichir FUSIONNÉ + corriger Prénom manquant ──────────────────────
    enriched = 0
    prenom_fixed = 0

    fuse["_skey"] = fuse.apply(lambda r: strict_key(r["Nom"], r["Prénom"], r["Métier"]), axis=1)
    fuse["_bkey"] = fuse.apply(lambda r: broad_key(r["Nom"], r["Métier"], r["Promotion"]), axis=1)

    for idx, row in fuse.iterrows():
        # Match strict d'abord
        crow = compl_skey.get(row["_skey"])
        # Si pas de match strict ET Prénom vide → essai match large
        if crow is None and (not row["Prénom"] or str(row["Prénom"]).strip() in ("", "nan")):
            crow = compl_bkey.get(row["_bkey"])
            if crow is not None:
                # Récupérer le Prénom manquant
                cp = crow.get("Prénom", "")
                if pd.notna(cp) and str(cp).strip() not in ("", "nan"):
                    fuse.at[idx, "Prénom"] = title_prenom(str(cp))
                    fuse.at[idx, "Nom complet"] = rebuild_nc(fuse.at[idx, "Nom"], fuse.at[idx, "Prénom"])
                    prenom_fixed += 1

        if crow is not None:
            # Enrichir les colonnes contact
            for col in new_cols:
                if col in crow.index:
                    val = crow[col]
                    if pd.notna(val) and str(val).strip() not in ("", "nan"):
                        cur = fuse.at[idx, col]
                        if not pd.notna(cur) or str(cur).strip() in ("", "nan"):
                            fuse.at[idx, col] = val
                            enriched += 1

    print(f"\n  Prénoms récupérés (cas Prénom manquant) : {prenom_fixed}")
    print(f"  Champs contact enrichis : {enriched}")

    # ── 4. Supprimer doublons résiduels (Prénom vide / rempli même ligne) ────
    # Recalculer les clés strictes après correction des prénoms
    fuse["_skey"] = fuse.apply(lambda r: strict_key(r["Nom"], r["Prénom"], r["Métier"]), axis=1)
    fuse["_bkey"] = fuse.apply(lambda r: broad_key(r["Nom"], r["Métier"], r["Promotion"]), axis=1)

    # Détecter les doublons (même bkey, promo renseignée, lignes multiples)
    bkey_counts = fuse[fuse["_bkey"].apply(lambda k: k[2] != "")]["_bkey"].value_counts()
    dup_bkeys = set(bkey_counts[bkey_counts > 1].index)
    dup_mask = fuse["_bkey"].isin(dup_bkeys) & (fuse["_bkey"].apply(lambda k: k[2] != ""))

    n_before = len(fuse)
    if dup_mask.sum() > 0:
        print(f"\n  Doublons bkey résiduels : {dup_mask.sum()} lignes")
        to_drop = []
        for bkey, grp in fuse[dup_mask].groupby("_bkey"):
            # Garder la plus complète
            grp2 = grp.copy()
            grp2["_fill"] = grp2.apply(count_fill, axis=1)
            keeper = grp2.sort_values("_fill", ascending=False).index[0]
            drop_idx = [i for i in grp2.index if i != keeper]
            to_drop.extend(drop_idx)
            print(f"    Gardé idx={keeper}, supprimé: {drop_idx}")
        fuse = fuse.drop(index=to_drop)
    print(f"  Lignes supprimées (doublons résiduels) : {n_before - len(fuse)}")

    # ── 5. Ajouter les nouvelles lignes de COMPLET ────────────────────────────
    fuse_skeys = set(fuse["_skey"].tolist())
    fuse_bkeys_with_promo = set(
        fuse[fuse["_bkey"].apply(lambda k: k[2] != "")]["_bkey"].tolist()
    )

    truly_new = []
    for _, crow in compl.iterrows():
        sk = crow["_skey"]
        bk = crow["_bkey"]
        # Exclure si déjà dans FUSIONNÉ (strict ou large)
        if sk in fuse_skeys:
            continue
        if bk[2] != "" and bk in fuse_bkeys_with_promo:
            continue
        # Exclure les inversions Nom/Prénom déjà traitées
        # (si skey inversé existe dans fusionné) - géré par sorted() dans strict_key
        truly_new.append(crow)

    if truly_new:
        new_df = pd.DataFrame(truly_new)
        # Aligner colonnes
        for col in fuse.columns:
            if col not in new_df.columns:
                new_df[col] = pd.NA
        new_df = new_df.reindex(columns=fuse.columns)
        fuse = pd.concat([fuse, new_df], ignore_index=True)

    print(f"  Nouvelles lignes ajoutées depuis COMPLET : {len(truly_new)}")

    # ── 6. Normaliser Nom/Prénom/Nom complet ─────────────────────────────────
    fuse["Nom"]    = fuse["Nom"].apply(lambda x: upper_nom(x) if pd.notna(x) else x)
    fuse["Prénom"] = fuse["Prénom"].apply(lambda x: title_prenom(x) if pd.notna(x) and str(x).strip() not in ("nan",) else x)
    fuse["Nom complet"] = fuse.apply(lambda r: rebuild_nc(r["Nom"], r["Prénom"]), axis=1)

    # ── 7. Nettoyage final ────────────────────────────────────────────────────
    fuse = fuse.drop(columns=["_skey", "_bkey", "_fill"], errors="ignore")

    # Réordonner colonnes (contact à la fin)
    base_cols = ["Nom", "Prénom", "Nom complet", "Métier", "Promotion", "Classe",
                 "Région", "Num_Dept", "Département", "Ville", "Statut",
                 "Entreprise", "Source", "Fiabilité"]
    extra_cols = ["Téléphone", "Email", "Site web", "Adresse", "Code postal", "URL source"]
    all_cols = base_cols + extra_cols
    # Ajouter colonnes manquantes
    for col in all_cols:
        if col not in fuse.columns:
            fuse[col] = pd.NA
    fuse = fuse[all_cols]

    # Trier
    fuse = fuse.sort_values(
        ["Num_Dept", "Promotion", "Nom"],
        key=lambda c: c.fillna("").astype(str)
    ).reset_index(drop=True)

    print(f"\n  ══ RÉSULTAT : {len(fuse)} MOFs ══")
    print(f"  Colonnes contact renseignées :")
    for col in extra_cols:
        n_filled = fuse[col].notna().sum()
        print(f"    {col:20s}: {n_filled}")

    # ── 8. Sauvegarde Excel ───────────────────────────────────────────────────
    print(f"\nSauvegarde → {FUSE_PATH}")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MOFs"

    header_fill = PatternFill(start_color="1A2E5A", end_color="1A2E5A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    cols = list(fuse.columns)
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for ri, row_data in enumerate(fuse.itertuples(index=False), 2):
        for ci, val in enumerate(row_data, 1):
            ws.cell(row=ri, column=ci, value=val if pd.notna(val) else None)

    for col_cells in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("Stats Métiers")
    stats_m = fuse.groupby("Métier").size().sort_values(ascending=False).reset_index(name="Nb")
    ws2.append(["Métier", "Nb"])
    for _, r in stats_m.iterrows():
        ws2.append([r["Métier"], r["Nb"]])

    ws3 = wb.create_sheet("Stats Départements")
    stats_d = fuse.groupby(["Num_Dept", "Département", "Région"]).size().sort_values(ascending=False).reset_index(name="Nb")
    ws3.append(["Num_Dept", "Département", "Région", "Nb"])
    for _, r in stats_d.iterrows():
        ws3.append([r["Num_Dept"], r["Département"], r["Région"], r["Nb"]])

    wb.save(FUSE_PATH)
    print(f"  ✓ Fichier sauvegardé")

    # ── Vérif SPIESER ─────────────────────────────────────────────────────────
    print("\n=== Vérif SPIESER ===")
    sp = fuse[fuse["Nom"].apply(norm).str.contains("SPIESER", na=False)]
    print(sp[["Nom", "Prénom", "Métier", "Promotion", "Téléphone", "Email", "Source"]].to_string())


if __name__ == "__main__":
    main()
