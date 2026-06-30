#!/usr/bin/env python3
"""
Détection de faux billets — ONCFM
Usage:
    python script.py --csv billets_production.csv
    python script.py --values 172.1 104.0 103.8 4.5 3.1 113.2
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import joblib

MODEL_DIR = Path(__file__).parent.parent / "data_clean"
FEATURES = [
    "diagonal",
    "height_left",
    "height_right",
    "margin_low",
    "margin_up",
    "length",
]
IMPUTER_FEATURES = ["diagonal", "height_left", "height_right", "margin_up", "length"]


def load_artifacts():
    try:
        model = joblib.load(MODEL_DIR / "model_final.pkl")
        scaler = joblib.load(MODEL_DIR / "scaler.pkl")
        imputer = joblib.load(MODEL_DIR / "imputer.pkl")
    except FileNotFoundError as e:
        sys.exit(f"Erreur : artefact manquant — {e}")
    return model, scaler, imputer


def preprocess(df, imputer, scaler):
    df = df[FEATURES].copy()
    missing = df["margin_low"].isnull()
    if missing.any():
        df.loc[missing, "margin_low"] = imputer.predict(
            df.loc[missing, IMPUTER_FEATURES]
        )
    return scaler.transform(df)


def predict_labels(X_scaled, model):
    return ["VRAI" if p == 1 else "FAUX" for p in model.predict(X_scaled)]


def run_csv(path, model, scaler, imputer):
    try:
        df = pd.read_csv(path, sep=";")
    except FileNotFoundError:
        sys.exit(f"Erreur : fichier introuvable — {path}")

    missing_cols = [c for c in FEATURES if c not in df.columns]
    if missing_cols:
        sys.exit(f"Erreur : colonnes manquantes dans le CSV : {missing_cols}")

    X_scaled = preprocess(df, imputer, scaler)
    for label in predict_labels(X_scaled, model):
        print(label)


def run_values(values, model, scaler, imputer):
    df = pd.DataFrame([values], columns=FEATURES)
    X_scaled = preprocess(df, imputer, scaler)
    print(predict_labels(X_scaled, model)[0])


def main():
    parser = argparse.ArgumentParser(
        description="Détection de faux billets — ONCFM",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Exemples :\n"
            "  python script.py --csv billets_production.csv\n"
            "  python script.py --values 172.1 104.0 103.8 4.5 3.1 113.2"
        ),
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--csv",
        metavar="FICHIER",
        help="CSV de billets à analyser (colonnes : diagonal height_left height_right margin_low margin_up length)",
    )
    group.add_argument(
        "--values",
        nargs=6,
        type=float,
        metavar=(
            "diagonal",
            "height_left",
            "height_right",
            "margin_low",
            "margin_up",
            "length",
        ),
        help="6 mesures physiques d'un seul billet",
    )
    args = parser.parse_args()

    model, scaler, imputer = load_artifacts()

    if args.csv:
        run_csv(args.csv, model, scaler, imputer)
    else:
        run_values(args.values, model, scaler, imputer)


if __name__ == "__main__":
    main()
