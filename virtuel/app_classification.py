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
        # --- Jointure 1 : articles <- Hiérarchisation (sur SAP ID) ---
        cols_hier = [
            "Sap ID/ Article", "Groupe Marché BDFR", "Id Categ BDFR", "Categ BDFR",
            "Id Sous Categ BDFR", "Sous Categ BDFR", "Id Sous Sous Categ BDFR", "Sous Sous Categ BDFR",
            "Subcategory", "Libellé Brick BDFR", "Libellé secteur", "Code Rayon", "Libellé rayon",
            "code famille", "Libellé Famille", "Grpe march.", "Groupe de marchandises",
            "Code sous-famille", "Libellé sous-famille",
        ]
        df = df_articles.merge(
            df_hier[cols_hier], left_on="SAP ID", right_on="Sap ID/ Article", how="left"
        )
        df = df.rename(columns={"Subcategory": "ID Brick_easier", "Libellé Brick BDFR": "Brick_easier"})

        # --- Jointure 2 : + Nouvelle catégorisation (sur ID Brick_easier <-> Id Brick) ---
        df = df.merge(
            df_categ[["Id Brick", "code_qualification", "libelle_qualification", "Marché", "Marché2"]],
            left_on="ID Brick_easier", right_on="Id Brick", how="left",
        )

        # --- Jointure 3 : + Qualification DVNI (sur code famille + Code sous-famille) ---
        df = df.merge(
            df_qualif[["CODE_FAM", "CODE_SFAM", "FAM_SSFAM", "LIB_FAM", "LIB_SFAM", "ÉLIGIBILITÉ_O_N"]],
            left_on=["code famille", "Code sous-famille"], right_on=["CODE_FAM", "CODE_SFAM"], how="left",
        )

        # --- Nettoyage : colonnes techniques ---
        drop_cols = [
            "Sap ID/ Article", "Id Brick", "code famille", "Code sous-famille",
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

    group_cols = ["Plano grouping desc", "Segmentation"]

    # 1. Comptage par groupe + code_qualification + libellé + éligibilité
    comptage = (
        df_c.groupby(group_cols + ["code_qualification", "libelle_qualification", "ÉLIGIBILITÉ_O_N"], dropna=False)
        .size()
        .reset_index(name="Nombre")
    )

    # 2. Somme par (groupe + code_qualification + libellé), éligibilité = celle du sous-groupe le plus fréquent
    def eligibilite_majoritaire(sous_groupe):
        return sous_groupe.loc[sous_groupe["Nombre"].idxmax(), "ÉLIGIBILITÉ_O_N"]

    somme_qualif = (
        comptage.groupby(group_cols + ["code_qualification", "libelle_qualification"], dropna=False)
        .apply(lambda g: pd.Series({"sum": g["Nombre"].sum(), "ÉLIGIBILITÉ_O_N": eligibilite_majoritaire(g)}))
        .reset_index()
    )

    # 3. Pour chaque groupe (Plano+Segmentation), garder la ligne avec le plus grand "sum"
    idx_max = somme_qualif.groupby(group_cols)["sum"].idxmax()
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

            df_plan_norm = df_plan_work.merge(df_corr_work, on=["Plano", "Segmentation"], how="left")
        else:
            df_plan_norm = df_plan_work.copy()
            df_plan_norm["Plano grouping desc"] = df_plan_norm["Plano"]
            df_plan_norm["Segmentation.1"] = df_plan_norm["Segmentation"]

        # Jointure finale avec la classification majoritaire
        df_final = df_plan_norm.merge(
            df_majoritaire,
            left_on=["Plano grouping desc", "Segmentation.1"],
            right_on=["Plano grouping desc", "Segmentation"],
            how="left",
            suffixes=("", "_classif"),
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