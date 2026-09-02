# Classification Plano / Segmentation

Application **Streamlit** de classification et de détermination de l'éligibilité des couples *Plano grouping desc* / *Segmentation* pour **Brico Dépôt (BDFR)**, à partir des référentiels articles, de la hiérarchisation produit, de la nouvelle catégorisation et de la qualification DVNI.

🔗 **Démo en ligne** : [abdkar95-classification-plano-segment-app-classification-mjonnk.streamlit.app](https://abdkar95-classification-plano-segment-app-classification-mjonnk.streamlit.app/)

## Contexte

Chez Brico Dépôt, chaque article est rattaché à un *Plano grouping desc* et une *Segmentation* utilisés pour structurer l'implantation en magasin (plan de sol). Cet outil permet de déterminer, pour chaque couple Plano/Segmentation, sa **qualification** et son **éligibilité** majoritaires à partir des articles qui le composent, puis de reporter cette classification sur le fichier de plan de sol.

## Pipeline

L'application enchaîne 4 étapes, chacune avec ses propres imports et exports CSV intermédiaires :

1. **Enrichissement de la liste d'articles**
   Fusion du fichier d'articles bruts avec :
   - la **hiérarchisation des produits** (rattachement rayon / famille / sous-famille / brick)
   - la **nouvelle catégorisation** (structure de gamme → qualification)
   - la **qualification DVNI** (éligibilité par famille / sous-famille)

2. **Détermination de la classification majoritaire**
   Pour chaque couple `Plano grouping desc` + `Segmentation`, calcul de la qualification et de l'éligibilité les plus représentées parmi les articles qui le composent (vote majoritaire par volume d'articles).

3. **Report sur le plan de sol**
   Jointure de cette classification sur le fichier « Plan de sol » (par `ID/DBkey`), avec gestion optionnelle d'une table de correspondance lorsque les libellés Plano/Segmentation du plan de sol diffèrent de ceux des articles.

4. **Export final**
   Génération d'un fichier consolidé : `ID/DBkey`, `Plano grouping desc`, `Segmentation`, `Code qualification`, `Qualification`, `Éligibilité`.

## Fichiers d'entrée attendus

| Fichier | Colonnes clés utilisées |
|---|---|
| Articles bruts (ex. articles 2022) | `SAP ID`, `Plano grouping desc`, `Segmentation` |
| Hiérarchisation des produits | `Sap ID/ Article`, `Subcategory`, `code famille`, `Code sous-famille`, … |
| Nouvelle catégorisation | `Id Brick`, `code_qualification`, `libelle_qualification`, `Marché`, `Marché2` |
| Qualification DVNI 19-21 | `CODE_FAM`, `CODE_SFAM`, `ÉLIGIBILITÉ_O_N`, … |
| Plan de sol | `ID/DBkey`, Plano, Segmentation (brut) |
| Correspondance (optionnel) | Plano/Segmentation brut → normalisé |

Chaque fichier est fourni au format `.xlsx`. L'application permet de choisir la feuille et, pour la hiérarchisation, d'ajuster le nombre de lignes à sauter avant l'en-tête.

## Installation

```bash
git clone https://github.com/Abdkar95/Classification-Plano-Segmentation.git
cd Classification-Plano-Segmentation
pip install -r requirements.txt
```

## Lancement en local

```bash
streamlit run app_classification.py
```

L'application s'ouvre ensuite sur `http://localhost:8501`.

## Stack technique

- **Streamlit** — interface web interactive
- **Pandas** — traitement et jointures des données
- **openpyxl** — lecture des fichiers Excel

## Sorties

Chaque étape du pipeline propose un export CSV (encodage `utf-8-sig`, compatible Excel) :
- `articles_enrichis.csv`
- `classification_plano_segmentation.csv`
- `plan_de_sol_classifie.csv`

## Auteur

Développé par [Abdoul-Karym Traoré](https://github.com/Abdkar95).
