# PLAN — PROJET-12 : Détection de faux billets (ONCFM)

## Objectif

Construire un modèle de classification vrai/faux billet à partir de caractéristiques physiques scannées, le packager en application autonome, et présenter les résultats à l'ONCFM.

## Périmètre

- **In scope** : EDA, preprocessing, entraînement et comparaison de 4 algorithmes (K-means, régression logistique, KNN, Random Forest), sélection du modèle final, script applicatif, support de présentation.
- **Out of scope** : acquisition de nouvelles données, déploiement en production, interface web.

## Données (`billets.csv`)

- Séparateur `;` — lire avec `pd.read_csv(..., sep=";")`
- Colonnes : `is_genuine`, `diagonal`, `height_left`, `height_right`, `margin_low`, `margin_up`, `length`
- `is_genuine` : pandas lit automatiquement en bool → convertir avec `.map({True: 1, False: 0, 'True': 1, 'False': 0})`
- Valeurs manquantes : uniquement dans `margin_low` (~20 lignes)
- Distribution : 1 000 vrais / 500 faux (ratio 2:1)

## Contraintes techniques (cahier des charges)

- **K-means** : utiliser les **centroïdes** pour prédire sur de nouvelles données (distance euclidienne au centroïde le plus proche). Ce n'est pas juste une visualisation.
- **Script** : accepter deux types d'input : (1) valeurs directes en argument CLI, (2) chemin vers un CSV multi-billets.
- **Output** : binaire simple — `VRAI` ou `FAUX` par billet.
- **Évaluation** : matrice de confusion obligatoire. Critère de sélection prioritaire = **recall sur la classe False** (maximiser la détection des faux billets).
- L'application sera testée en direct avec Marie sur `billets_production.csv` (même format que `billets.csv`, sans colonne `is_genuine`).
- Livraison : notebook + script + support.

## Résultats (notebook exécuté)

| Modèle | Accuracy | Recall(Faux) | F1(Faux) | AUC-ROC |
|---|---|---|---|---|
| K-means | 0.987 | 0.984 | 0.980 | — |
| Rég. logistique | 0.992 | 0.984 | 0.988 | 0.9996 |
| KNN | 0.992 | 0.984 | 0.988 | **0.9997** |
| Random Forest | 0.992 | 0.984 | 0.988 | 0.9994 |

**Modèle final : KNN** (meilleur AUC-ROC parmi les modèles à égalité sur Recall et F1).
Artefacts : `data_clean/model_final.pkl`, `data_clean/scaler.pkl`, `data_clean/imputer.pkl`.

## Livrables et nommage

| # | Fichier | Nom attendu |
|---|---|---|
| 1 | Notebook d'analyse | `Elices-Diez_Josy_1_Notebook_analyse_062026.ipynb` |
| 2 | Script applicatif | `src/Elices-Diez_Josy_2_Script_application_062026.py` |
| 3 | Support de présentation | `Elices-Diez_Josy_3_Support_presentation_062026.pptx` |

Dossier zip : `Detection_faux_billets_Elices-Diez_Josy.zip` — contient les 3 livrables.
Support : 20 slides maximum, format PPT.

## Soutenance (30 min)

- **15 min — Présentation** : cheminement complet, traitements amont, pistes explorées, modèle final + justification.
- **10 min — Discussion** : test en direct du script sur `billets_production.csv` fourni pendant la soutenance (format identique aux données d'entraînement) + Q&R sur les choix techniques.
  - Questions types : split train/test, alternatives algorithmiques, justification du modèle final.
- **5 min — Débrief** hors rôle avec l'évaluateur.

**Point critique** : le script doit tourner en autonomie sur un CSV inconnu, en direct. Tester ce cas avant la soutenance.

---

## Plan étape par étape

### Étape 1 — Setup du projet
- [x] Créer la structure de répertoires (`data_raw/`, `data_clean/`, `figures/`, `src/`)
- [x] Initialiser l'environnement uv (`pyproject.toml`)
- [x] Ajouter les dépendances : pandas, scikit-learn, matplotlib, seaborn, jupyter, openpyxl, joblib
- [x] Placer `billets.csv` dans `data_raw/`

### Étape 2 — Exploration des données (EDA)
- [x] Charger le fichier et vérifier le schéma
- [x] Statistiques descriptives par classe
- [x] Visualisation des distributions (histogrammes, boxplots)
- [x] Matrice de corrélation
- [x] Valeurs manquantes identifiées (`margin_low`, ~20 lignes)

### Étape 3 — Preprocessing
- [x] Convertir `is_genuine` → 0/1
- [x] Imputer `margin_low` par régression linéaire
- [x] Standardisation (StandardScaler)
- [x] Split train/test stratifié 75/25

### Étape 4 — Modélisation : K-means
- [x] Entraîner K-means (k=2) sur données standardisées
- [x] Aligner clusters avec classes réelles
- [x] Prédiction par distance euclidienne aux centroïdes
- [x] Métriques calculées (matrice de confusion, recall)

### Étape 5 — Modélisation : Régression logistique
- [x] Entraîner sur train set standardisé
- [x] Analyser les coefficients
- [x] Métriques calculées sur test set

### Étape 6 — Modélisation : KNN
- [x] Optimiser k par GridSearchCV (scoring=recall, cv=5)
- [x] Métriques calculées sur test set

### Étape 7 — Modélisation : Random Forest
- [x] Entraîner (200 estimateurs)
- [x] Feature importances visualisées
- [x] Métriques calculées sur test set

### Étape 8 — Comparaison et sélection du modèle final
- [x] Tableau comparatif des 4 algorithmes
- [x] KNN retenu (meilleur AUC-ROC parmi égalité sur Recall/F1)
- [x] Modèle sauvegardé (`data_clean/model_final.pkl`, `scaler.pkl`, `imputer.pkl`)

### Étape 9 — Script applicatif
- [x] Créer `src/Elices-Diez_Josy_2_Script_application_062026.py`
- [x] Mode 1 : `python script.py --csv billets_production.csv` → prédit chaque ligne
- [x] Mode 2 : `python script.py --values 172.1 104.0 103.8 4.5 3.1 113.2` → prédit un billet
- [x] Output : une ligne par billet — `VRAI` ou `FAUX`
- [x] Testé sur vrais et faux billets, résultats corrects

### Étape 10 — Support de présentation
- [ ] Slide 1 : contexte et objectif
- [ ] Slides 2-3 : données et EDA (chiffres clés, distributions)
- [ ] Slides 4-7 : un slide par algorithme (méthode + résultats)
- [ ] Slide 8 : tableau comparatif + choix du modèle final justifié
- [ ] Slide 9 : démonstration de l'application (live avec Marie)
- [ ] Format : python-pptx (selon standards CLAUDE.md)

---

## Statut

- Étape courante : **10 — Support de présentation**
- Étapes 1 à 9 terminées
- Restant : créer `Elices-Diez_Josy_3_Support_presentation_062026.pptx`
