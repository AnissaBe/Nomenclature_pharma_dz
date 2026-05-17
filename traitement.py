import pandas as pd
import re


def traiter_positionnement(uploaded_file):
    """
    Prend un fichier CSV uploadé et retourne :
    - par_labo   : une ligne par laboratoire (format collègue)
    - par_code   : une ligne par code produit (notre format)
    - anomalies  : codes avec plusieurs DCI différentes
    """

    # ── Chargement ───────────────────────────────────────────────
    try:
        df = pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python",
            encoding="utf-8"
        )
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df = pd.read_csv(
            uploaded_file,
            sep=None,
            engine="python",
            encoding="latin-1"
        )

    # ── Supprimer colonnes sans nom (Unnamed) ────────────────────
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    # ── Supprimer première colonne N° si elle existe ──────────────
    if df.columns[0] in ["N°", "N°ENREGISTREMENT", "N\u00b0", "N\u00b0ENREGISTREMENT"]:
        df = df.drop(columns=df.columns[0])

    df = df.reset_index(drop=True)

    # ── Noms des colonnes clés ───────────────────────────────────
    COL_CODE   = "CODE"
    COL_DCI    = "DENOMINATION COMMUNE INTERNATIONALE"
    COL_FORME  = "FORME"
    COL_DOSAGE = "DOSAGE"
    COL_COND   = "CONDITIONNEMENT"
    COL_LABO   = "LABORATOIRES DETENTEUR DE LA DECISION D'ENREGISTREMENT"
    COL_PAYS   = "PAYS DU LABORATOIRE DETENTEUR DE LA DECISION D'ENREGISTREMENT"
    COL_STATUT = "STATUT"

    # ── Nettoyage dosages ────────────────────────────────────────
    df[COL_DOSAGE] = df[COL_DOSAGE].apply(clean_dosage)

    # ════════════════════════════════════════════════════════════
    # ONGLET 1 — Par laboratoire (format collègue)
    # ════════════════════════════════════════════════════════════
    tableau = df[[
        COL_CODE, COL_DCI, COL_FORME, COL_DOSAGE,
        COL_COND, COL_LABO, COL_PAYS, COL_STATUT
    ]].copy()

    tableau["FABRIQUE"] = tableau[COL_STATUT].apply(lambda x: "R1" if str(x).strip().upper() == "F" else "")
    tableau["IMPORTE"]  = tableau[COL_STATUT].apply(lambda x: "R1" if str(x).strip().upper() == "I" else "")
    tableau = tableau.drop(columns=[COL_STATUT])

    tableau = tableau.rename(columns={
        COL_DCI:   "DCI",
        COL_COND:  "COND.",
        COL_LABO:  "LABORATOIRE",
        COL_PAYS:  "PAYS"
    })

    # Lignes TOTAL par CODE
    totaux = []
    for code, group in tableau.groupby(COL_CODE, sort=False):
        nb_f = (group["FABRIQUE"] == "R1").sum()
        nb_i = (group["IMPORTE"]  == "R1").sum()
        total = nb_f + nb_i
        totaux.append({
            COL_CODE:      f"Total {code}",
            "DCI":         "",
            COL_FORME:     "",
            COL_DOSAGE:    "",
            "COND.":       "",
            "LABORATOIRE": "",
            "PAYS":        "",
            "FABRIQUE":    f"R{nb_f}"  if nb_f > 0 else "",
            "IMPORTE":     f"R{nb_i}"  if nb_i > 0 else "",
            "TOTAL":       f"R{total}"
        })

    tableau["TOTAL"] = ""
    totaux_df = pd.DataFrame(totaux)

    result = []
    for code, group in tableau.groupby(COL_CODE, sort=False):
        result.append(group)
        total_row = totaux_df[totaux_df[COL_CODE] == f"Total {code}"]
        result.append(total_row)

    # Ligne TOTAL général
    nb_f_total = (tableau["FABRIQUE"] == "R1").sum()
    nb_i_total = (tableau["IMPORTE"]  == "R1").sum()
    total_general = pd.DataFrame([{
        COL_CODE:      "TOTAL",
        "DCI":         "",
        COL_FORME:     "",
        COL_DOSAGE:    "",
        "COND.":       "",
        "LABORATOIRE": "",
        "PAYS":        "",
        "FABRIQUE":    f"R{nb_f_total}",
        "IMPORTE":     f"R{nb_i_total}",
        "TOTAL":       f"R{nb_f_total + nb_i_total}"
    }])
    result.append(total_general)

    par_labo = pd.concat(result, ignore_index=True)

    # ════════════════════════════════════════════════════════════
    # ONGLET 2 — Par code (notre format)
    # ════════════════════════════════════════════════════════════
    par_code = df.groupby(COL_CODE).agg(
        DCI=(COL_DCI, "first"),
        FORME=(COL_FORME, "first"),
        DOSAGE=(COL_DOSAGE, "first"),
        **{"COND.": (COL_COND, "first")},
        LABORATOIRES=(COL_LABO, lambda x: ", ".join(x.dropna().unique())),
        PAYS=(COL_PAYS, lambda x: ", ".join(x.dropna().unique())),
        FABRIQUE=(COL_STATUT, lambda x: f"R{(x.str.strip().str.upper()=='F').sum()}"
                  if (x.str.strip().str.upper()=='F').sum() > 0 else ""),
        IMPORTE=(COL_STATUT, lambda x: f"R{(x.str.strip().str.upper()=='I').sum()}"
                 if (x.str.strip().str.upper()=='I').sum() > 0 else ""),
        TOTAL=(COL_STATUT, lambda x: f"R{len(x)}")
    ).reset_index()

    par_code = par_code.sort_values(COL_CODE).reset_index(drop=True)

    # ════════════════════════════════════════════════════════════
    # ONGLET 3 — Anomalies codes
    # ════════════════════════════════════════════════════════════
    anomalies_check = df.groupby(COL_CODE)[COL_DCI].nunique()
    anomalies_check = anomalies_check[anomalies_check > 1]

    rows = []
    for code in anomalies_check.index:
        dcis = df[df[COL_CODE] == code][COL_DCI].unique()
        row = {COL_CODE: code}
        for i, dci in enumerate(dcis):
            row[f"DCI_{i+1}"] = dci
        rows.append(row)

    anomalies = pd.DataFrame(rows).fillna("") if rows else pd.DataFrame()

    return par_labo, par_code, anomalies


# ── Nettoyage dosages ─────────────────────────────────────────────
def clean_dosage(s):
    if pd.isna(s):
        return s
    s = str(s).strip().upper()
    unites = r'(MG/ML|G/ML|MG/G|MG/L|UI/ML|UI/MG|MCG/ML|MG/CM2|MMOL/ML|MEQ/ML|MG|MCG|UI|GR|ML|G|%)'
    s = re.sub(r'(\d[\d,\.]*)\s+' + unites, r'\1\2', s)
    s = re.sub(r'\(\s+', '(', s)
    s = re.sub(r'\s+\)', ')', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s