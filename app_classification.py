import streamlit as st
import pandas as pd

st.set_page_config(page_title="Classification Plano / Segmentation", layout="wide")
st.title("Classification et éligibilité des Plano / Segmentation")

st.markdown(
    """
Cette application reproduit le pipeline suivant :

1. **Enrichissement de la liste d'articles** (Hiérarchisation des produits + Nouvelle catégorisation + Qualification DVNI)
2. **Détermination de la Qualification / Éligibilité majoritaire** par couple `Plano grouping desc` + `Segmentation`
3. **Report de cette classification sur le fichier "Plan de sol"** (par `ID/DBkey`)

**Sortie finale** : `ID/DBkey`, `Plano grouping desc`, `Segmentation`, `Code qualification`, `Qualification`, `Éligibilité`
"""
)


# ============================================================
# Fonctions utilitaires
# ============================================================
def read_excel_sheet(uploaded_file, sheet_hint=None, skiprows=0, allow_skiprows_adjust=False):
    """Lit un fichier Excel uploadé et laisse choisir la feuille si plusieurs existent."""
    if uploaded_file is None:
        return None
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    default_idx = 0
    if sheet_hint:
        for i, s in enumerate(sheets):
            if sheet_hint.lower() in s.lower():
                default_idx = i
                break
    sheet = st.selectbox(
        f"Feuille à utiliser ({uploaded_file.name})", sheets, index=default_idx, key=uploaded_file.name + "_sheet"
    )

    if allow_skiprows_adjust:
        raw_preview = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, nrows=5)
        with st.expander(f"Aperçu brut de « {uploaded_file.name} » (pour régler le nombre de lignes à sauter)"):
            st.dataframe(raw_preview, use_container_width=True)
        skiprows = st.number_input(
            f"Nombre de lignes à sauter avant l'en-tête ({uploaded_file.name})",
            min_value=0, max_value=10, value=skiprows, step=1, key=uploaded_file.name + "_skiprows",
        )

    df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=skiprows)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def check_columns(df, required_cols, label):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Colonnes manquantes dans « {label} » : {missing}")
        st.write("Colonnes disponibles :", list(df.columns))
        return False
    return True


def normalize_key(series):
    """
    Normalise une colonne servant de clé de jointure : force en texte,
    supprime les espaces superflus, retire un éventuel '.0' issu d'une
    lecture Excel qui a interprété un ID numérique comme un float.
    Évite les jointures silencieusement vides à cause d'un type ou
    d'un espace différent entre deux fichiers.
    """
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    return s


# ============================================================
# ÉTAPE 1 — Chargement des fichiers articles
# ============================================================
st.header("1. Fichiers sources — Articles")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Articles (année)")
    f_articles = st.file_uploader("Fichier des articles bruts (ex: articles 2022)", type=["xlsx"], key="articles")
    df_articles = read_excel_sheet(f_articles, sheet_hint="articles") if f_articles else None

    st.subheader("Hiérarchisation des produits")
    f_hier = st.file_uploader("Fichier de hiérarchisation des produits", type=["xlsx"], key="hier")
    df_hier = read_excel_sheet(f_hier, sheet_hint="feuil", skiprows=1, allow_skiprows_adjust=True) if f_hier else None

with col2:
    st.subheader("Nouvelle catégorisation")
    f_categ = st.file_uploader("Fichier « Catégorisation nouvelle structure de gamme »", type=["xlsx"], key="categ")
    df_categ = read_excel_sheet(f_categ, sheet_hint="catégorisation") if f_categ else None

    st.subheader("Qualification DVNI")
    f_qualif = st.file_uploader("Fichier « Qualification DVNI 19 à 21 »", type=["xlsx"], key="qualif")
    df_qualif = read_excel_sheet(f_qualif, sheet_hint="qualification") if f_qualif else None

articles_ready = all(x is not None for x in [df_articles, df_hier, df_categ, df_qualif])

# ============================================================
# ÉTAPE 2 — Enrichissement des articles
# ============================================================
df_enrichi = None

if articles_ready:
    st.header("2. Enrichissement des articles")

    ok = True
    ok &= check_columns(df_articles, ["SAP ID", "Plano grouping desc", "Segmentation"], "Articles")
    ok &= check_columns(
        df_hier,
        [
            "Sap ID/ Article", "Groupe Marché BDFR", "Id Categ BDFR", "Categ BDFR",
            "Id Sous Categ BDFR", "Sous Categ BDFR", "Id Sous Sous Categ BDFR", "Sous Sous Categ BDFR",
            "Subcategory", "Libellé Brick BDFR", "Libellé secteur", "Code Rayon", "Libellé rayon",
            "code famille", "Libellé Famille", "Grpe march.", "Groupe de marchandises",
            "Code sous-famille", "Libellé sous-famille",
        ],
        "Hiérarchisation",
    )
    ok &= check_columns(df_categ, ["Id Brick", "code_qualification", "libelle_qualification", "Marché", "Marché2"], "Nouvelle catégorisation")
    ok &= check_columns(df_qualif, ["CODE_FAM", "CODE_SFAM", "FAM_SSFAM", "LIB_FAM", "LIB_SFAM", "ÉLIGIBILITÉ_O_N"], "Qualification DVNI")

    if ok:
        df_articles = df_articles.copy()
        df_hier = df_hier.copy()
        df_categ = df_categ.copy()
        df_qualif = df_qualif.copy()

        # --- Jointure 1 : articles <- Hiérarchisation (sur SAP ID) ---
        # Clés normalisées : évite les NaN silencieux si les types/espaces diffèrent
        df_articles["_SAP_ID_key"] = normalize_key(df_articles["SAP ID"])
        df_hier["_SAP_ID_key"] = normalize_key(df_hier["Sap ID/ Article"])

        cols_hier = [
            "_SAP_ID_key", "Groupe Marché BDFR", "Id Categ BDFR", "Categ BDFR",
            "Id Sous Categ BDFR", "Sous Categ BDFR", "Id Sous Sous Categ BDFR", "Sous Sous Categ BDFR",
            "Subcategory", "Libellé Brick BDFR", "Libellé secteur", "Code Rayon", "Libellé rayon",
            "code famille", "Libellé Famille", "Grpe march.", "Groupe de marchandises",
            "Code sous-famille", "Libellé sous-famille",
        ]
        df = df_articles.merge(df_hier[cols_hier], on="_SAP_ID_key", how="left")
        df = df.rename(columns={"Subcategory": "ID Brick_easier", "Libellé Brick BDFR": "Brick_easier"})

        nb_non_matches_1 = df["Groupe Marché BDFR"].isna().sum()
        if nb_non_matches_1 > 0:
            st.warning(f"Jointure Articles ↔ Hiérarchisation : {nb_non_matches_1} article(s) sans correspondance.")

        # --- Jointure 2 : + Nouvelle catégorisation (sur ID Brick_easier <-> Id Brick) ---
        df["_ID_BRICK_key"] = normalize_key(df["ID Brick_easier"])
        df_categ["_ID_BRICK_key"] = normalize_key(df_categ["Id Brick"])

        df = df.merge(
            df_categ[["_ID_BRICK_key", "code_qualification", "libelle_qualification", "Marché", "Marché2"]],
            on="_ID_BRICK_key", how="left",
        )

        nb_non_matches_2 = df["code_qualification"].isna().sum()
        if nb_non_matches_2 > 0:
            st.warning(f"Jointure Hiérarchisation ↔ Nouvelle catégorisation : {nb_non_matches_2} article(s) sans correspondance.")

        # --- Jointure 3 : + Qualification DVNI (sur code famille + Code sous-famille) ---
        df["_FAM_key"] = normalize_key(df["code famille"])
        df["_SFAM_key"] = normalize_key(df["Code sous-famille"])
        df_qualif["_FAM_key"] = normalize_key(df_qualif["CODE_FAM"])
        df_qualif["_SFAM_key"] = normalize_key(df_qualif["CODE_SFAM"])

        df = df.merge(
            df_qualif[["_FAM_key", "_SFAM_key", "FAM_SSFAM", "LIB_FAM", "LIB_SFAM", "ÉLIGIBILITÉ_O_N"]],
            on=["_FAM_key", "_SFAM_key"], how="left",
        )

        nb_non_matches_3 = df["ÉLIGIBILITÉ_O_N"].isna().sum()
        if nb_non_matches_3 > 0:
            st.warning(f"Jointure Hiérarchisation ↔ Qualification DVNI : {nb_non_matches_3} article(s) sans correspondance.")

        # --- Nettoyage : colonnes techniques ---
        drop_cols = [
            "_SAP_ID_key", "_ID_BRICK_key", "_FAM_key", "_SFAM_key",
            "Id Brick", "code famille", "Code sous-famille",
            "Libellé Famille", "Groupe Marché BDFR", "Libellé sous-famille",
        ]
        df = df.drop(columns=[c for c in drop_cols if c in df.columns])

        df_enrichi = df
        st.success(f"Articles enrichis : {len(df_enrichi)} lignes")
        st.dataframe(df_enrichi.head(20), use_container_width=True)
        st.download_button(
            "Télécharger les articles enrichis (CSV)",
            df_enrichi.to_csv(index=False).encode("utf-8-sig"),
            file_name="articles_enrichis.csv",
            mime="text/csv",
        )

# ============================================================
# ÉTAPE 3 — Classification majoritaire par Plano + Segmentation
# ============================================================
df_majoritaire = None

if df_enrichi is not None:
    st.header("3. Classification majoritaire par Plano + Segmentation")

    df_c = df_enrichi.copy()
    df_c["libelle_qualification"] = df_c["libelle_qualification"].fillna("Non qualifié")
    df_c["code_qualification"] = df_c["code_qualification"].fillna("NA")
    df_c["ÉLIGIBILITÉ_O_N"] = df_c["ÉLIGIBILITÉ_O_N"].fillna("Non défini")

    group_cols = ["Plano grouping desc", "Segmentation"]

    # 1. Comptage par groupe + code_qualification + libellé + éligibilité
    comptage = (
        df_c.groupby(group_cols + ["code_qualification", "libelle_qualification", "ÉLIGIBILITÉ_O_N"], dropna=False)
        .size()
        .reset_index(name="Nombre")
    )

    # 2. Somme par (groupe + code_qualification + libellé), éligibilité = celle du
    #    sous-groupe le plus fréquent. Remplace l'ancien .apply(lambda g: pd.Series(...))
    #    qui déclenche des avertissements/erreurs selon la version de pandas :
    #    on trie par Nombre décroissant puis on prend la 1ère valeur d'éligibilité
    #    rencontrée par groupe (= celle du sous-groupe majoritaire), de façon vectorisée.
    comptage_trie = comptage.sort_values("Nombre", ascending=False)
    somme_qualif = comptage_trie.groupby(
        group_cols + ["code_qualification", "libelle_qualification"], dropna=False, as_index=False
    ).agg(**{"sum": ("Nombre", "sum"), "ÉLIGIBILITÉ_O_N": ("ÉLIGIBILITÉ_O_N", "first")})

    if somme_qualif.empty:
        st.warning("Aucune ligne à classifier (vérifiez les jointures de l'étape 2).")
    else:
        # 3. Pour chaque groupe (Plano+Segmentation), garder la ligne avec le plus grand "sum"
        #    dropna=False ajouté ici aussi : sans ça, un couple Plano/Segmentation contenant
        #    un NaN disparaissait silencieusement à cette dernière étape alors qu'il était
        #    bien présent dans les étapes précédentes (incohérence avec le reste du groupby).
        idx_max = somme_qualif.groupby(group_cols, dropna=False)["sum"].idxmax()
        df_majoritaire = somme_qualif.loc[idx_max].reset_index(drop=True)
        df_majoritaire = df_majoritaire.rename(
            columns={
                "code_qualification": "Code qualification",
                "libelle_qualification": "Qualification",
                "ÉLIGIBILITÉ_O_N": "Éligibilité",
            }
        ).drop(columns=["sum"])

        st.success(f"{len(df_majoritaire)} couples Plano/Segmentation classifiés")
        st.dataframe(df_majoritaire, use_container_width=True)
        st.download_button(
            "Télécharger la classification Plano/Segmentation (CSV)",
            df_majoritaire.to_csv(index=False).encode("utf-8-sig"),
            file_name="classification_plano_segmentation.csv",
            mime="text/csv",
        )

# ============================================================
# ÉTAPE 4 — Report sur le Plan de sol
# ============================================================
if df_majoritaire is not None:
    st.header("4. Report sur le Plan de sol")

    st.info(
        "Les libellés Plano/Segmentation du plan de sol brut sont souvent écrits différemment "
        "de ceux des articles (ex. « ACCESSOIRES PLACO » vs « ALARA EXPO »). "
        "Si c'est votre cas, fournissez aussi un fichier de correspondance "
        "(brut → normalisé), sinon la jointure se fera directement sur les libellés."
    )

    f_plan = st.file_uploader("Fichier Plan de sol (ID/DBkey, Plano, Segmentation)", type=["xlsx"], key="plan")
    f_correspondance = st.file_uploader(
        "Fichier de correspondance Plano/Segmentation brut → normalisé (optionnel)", type=["xlsx"], key="correspondance"
    )

    if f_plan is not None:
        df_plan = read_excel_sheet(f_plan, sheet_hint="plan")

        with st.expander("Aperçu du plan de sol"):
            st.dataframe(df_plan.head(20), use_container_width=True)

        id_col = st.selectbox("Colonne ID/DBkey", df_plan.columns, key="id_col")
        plano_col = st.selectbox("Colonne Plano (brut)", df_plan.columns, key="plano_col")
        seg_col = st.selectbox("Colonne Segmentation (brut)", df_plan.columns, key="seg_col")

        df_plan_work = df_plan[[id_col, plano_col, seg_col]].rename(
            columns={id_col: "ID/DBkey", plano_col: "Plano", seg_col: "Segmentation"}
        )

        if f_correspondance is not None:
            df_corr = read_excel_sheet(f_correspondance, sheet_hint="correspondance")
            with st.expander("Aperçu de la table de correspondance"):
                st.dataframe(df_corr.head(20), use_container_width=True)

            corr_plano_brut = st.selectbox("Colonne Plano brut (correspondance)", df_corr.columns, key="corr_plano_brut")
            corr_seg_brut = st.selectbox("Colonne Segmentation brut (correspondance)", df_corr.columns, key="corr_seg_brut")
            corr_plano_norm = st.selectbox("Colonne Plano normalisé (correspondance)", df_corr.columns, key="corr_plano_norm")
            corr_seg_norm = st.selectbox("Colonne Segmentation normalisée (correspondance)", df_corr.columns, key="corr_seg_norm")

            df_corr_work = df_corr[[corr_plano_brut, corr_seg_brut, corr_plano_norm, corr_seg_norm]].rename(
                columns={
                    corr_plano_brut: "Plano",
                    corr_seg_brut: "Segmentation",
                    corr_plano_norm: "Plano grouping desc",
                    corr_seg_norm: "Segmentation.1",
                }
            )
            df_corr_work = df_corr_work.drop_duplicates(subset=["Plano", "Segmentation"])

            # Normalisation des clés brutes avant jointure (espaces/casse fréquents entre exports)
            df_plan_work["_Plano_key"] = normalize_key(df_plan_work["Plano"])
            df_plan_work["_Seg_key"] = normalize_key(df_plan_work["Segmentation"])
            df_corr_work["_Plano_key"] = normalize_key(df_corr_work["Plano"])
            df_corr_work["_Seg_key"] = normalize_key(df_corr_work["Segmentation"])

            df_plan_norm = df_plan_work.merge(
                df_corr_work[["_Plano_key", "_Seg_key", "Plano grouping desc", "Segmentation.1"]],
                on=["_Plano_key", "_Seg_key"], how="left",
            ).drop(columns=["_Plano_key", "_Seg_key"])

            nb_non_normalises = df_plan_norm["Plano grouping desc"].isna().sum()
            if nb_non_normalises > 0:
                st.warning(
                    f"{nb_non_normalises} ligne(s) du plan de sol sans correspondance dans la table "
                    "de normalisation Plano/Segmentation."
                )
        else:
            df_plan_norm = df_plan_work.copy()
            df_plan_norm["Plano grouping desc"] = df_plan_norm["Plano"]
            df_plan_norm["Segmentation.1"] = df_plan_norm["Segmentation"]

        # Jointure finale avec la classification majoritaire (clés normalisées)
        df_plan_norm["_Plano_key"] = normalize_key(df_plan_norm["Plano grouping desc"])
        df_plan_norm["_Seg_key"] = normalize_key(df_plan_norm["Segmentation.1"])

        df_majoritaire_key = df_majoritaire.copy()
        df_majoritaire_key["_Plano_key"] = normalize_key(df_majoritaire_key["Plano grouping desc"])
        df_majoritaire_key["_Seg_key"] = normalize_key(df_majoritaire_key["Segmentation"])

        df_final = df_plan_norm.merge(
            df_majoritaire_key[["_Plano_key", "_Seg_key", "Code qualification", "Qualification", "Éligibilité"]],
            on=["_Plano_key", "_Seg_key"],
            how="left",
        )

        df_final_out = df_final[
            ["ID/DBkey", "Plano", "Segmentation", "Code qualification", "Qualification", "Éligibilité"]
        ]

        nb_non_trouves = df_final_out["Code qualification"].isna().sum()
        if nb_non_trouves > 0:
            st.warning(
                f"{nb_non_trouves} ligne(s) du plan de sol n'ont pas trouvé de classification "
                "(Plano/Segmentation absent de la classification des articles)."
            )

        st.success(f"Résultat final : {len(df_final_out)} lignes")
        st.dataframe(df_final_out, use_container_width=True)
        st.download_button(
            "Télécharger le résultat final (CSV)",
            df_final_out.to_csv(index=False).encode("utf-8-sig"),
            file_name="plan_de_sol_classifie.csv",
            mime="text/csv",
        )