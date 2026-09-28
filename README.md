# Détection de faux billets : ONCFM

> Modèle de classification vrai/faux billet à partir de 6 mesures physiques scannées, livré sous forme de script en ligne de commande.

**Stack** · Python · pandas · scikit-learn · Matplotlib · Seaborn · Jupyter

---

## Contexte

Mission pour l'Organisation nationale de lutte contre le faux-monnayage (ONCFM) : détecter automatiquement les faux billets à partir de leurs dimensions et marges, avec un script autonome qui doit fonctionner sur des billets jamais vus.

## Méthodologie

1. **Exploration** de 1 500 billets (1 000 vrais, 500 faux), 6 mesures par billet
2. **Imputation** des 37 valeurs manquantes de `margin_low`, après comparaison de deux méthodes (régression linéaire et KNN) par validation croisée à 5 plis
3. **Standardisation** et découpage train/test stratifié
4. **Comparaison de 4 algorithmes** : K-means, régression logistique, KNN, Random Forest
5. **Sélection du modèle** sur le critère prioritaire : le recall de la classe « faux », pour laisser passer le moins de faux billets possible
6. **Validation croisée** à 5 plis et comparaison par coût d'erreur
7. **Script de prédiction** en ligne de commande

## Résultats clés

- Modèle retenu : **régression logistique**, environ **99,4 % d'accuracy** et **98,8 % de recall** sur la classe faux
- KNN, Random Forest et K-means atteignent entre 98,6 % et 99,0 % d'accuracy
- Recall moyen de la classe faux en validation croisée : 0,980 ± 0,014

## Limites

- Le jeu de test est modeste (495 billets) : la variance sur des billets réels plus hétérogènes n'est pas mesurée
- Le modèle n'a pas été éprouvé sur un flux de billets scannés en conditions réelles

## Contenu du repo

| Fichier | Description |
|---------|-------------|
| `Elices-Diez_Josy_1_Notebook_analyse_062026.ipynb` | Notebook d'analyse complet : imputation, 4 modèles, validation croisée, coût |
| `Elices-Diez_Josy_2_Notebook_exploration_062026.ipynb` | Version exploratoire antérieure, sans validation croisée ni comparaison par coût |
| `src/Elices-Diez_Josy_2_Script_application_062026.py` | Script de prédiction en ligne de commande |
| `Elices-Diez_Josy_3_Support_presentation_062026.pptx` | Support de soutenance |
| `Cahier+des+charges+détection+faux+billets_P12_DAS.pdf` | Cahier des charges de la mission |
| `data_raw/billets.csv` | Données brutes |
| `data_clean/` | Données nettoyées et artefacts du modèle (modèle, scaler, imputer) |
| `figures/` | Visualisations exportées |
| `build_pptx.py` | Script qui génère le support de soutenance |

## Utiliser le script

```bash
uv sync
uv run python src/Elices-Diez_Josy_2_Script_application_062026.py --csv billets_production.csv
uv run python src/Elices-Diez_Josy_2_Script_application_062026.py --values 172.1 104.0 103.8 4.5 3.1 113.2
```

Le script charge le modèle, le scaler et l'imputer depuis `data_clean/`, impute `margin_low` s'il manque, puis renvoie VRAI ou FAUX pour chaque billet.
L'ordre des valeurs de `--values` suit les colonnes `diagonal`, `height_left`, `height_right`, `margin_low`, `margin_up`, `length`.

---

*Formation Data Analyst OpenClassrooms × ENSAE*
