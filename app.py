import os
import json
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "perfcnc_secret_key"

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max

COLONNES_REQUISES = [
    "Date",
    "Équipe",
    "Code CRMX",
    "Référence pièce",
    "Famille",
    "Cause arrêt",
    "Heures produites",
    "Pièces KO",
]


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def process_excel(filepath):
    df = pd.read_excel(filepath)

    # Normalize column names (strip spaces)
    df.columns = [c.strip() for c in df.columns]

    # Check required columns
    missing = [c for c in COLONNES_REQUISES if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(missing)}")

    # Keep only required columns
    df = df[COLONNES_REQUISES].copy()

    # Convert numeric columns
    df["Heures produites"] = pd.to_numeric(df["Heures produites"], errors="coerce").fillna(0)
    df["Pièces KO"] = pd.to_numeric(df["Pièces KO"], errors="coerce").fillna(0)

    # Format date
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d/%m/%Y")
    df["Date"] = df["Date"].fillna("")

    # --- KPIs ---
    total_saisies = len(df)
    moyenne_heures = round(df["Heures produites"].mean(), 2)
    total_rebuts = int(df["Pièces KO"].sum())

    # --- Causes d'arrêt (bar chart) ---
    causes = (
        df["Cause arrêt"]
        .fillna("Non renseigné")
        .value_counts()
        .sort_values(ascending=False)
    )
    causes_labels = causes.index.tolist()
    causes_values = causes.values.tolist()

    # --- Tableau par famille ---
    famille_group = (
        df.groupby("Famille", dropna=False)
        .agg(
            Saisies=("Famille", "count"),
            Heures_produites=("Heures produites", "sum"),
            Pieces_KO=("Pièces KO", "sum"),
        )
        .reset_index()
    )
    famille_group["Heures_produites"] = famille_group["Heures_produites"].round(2)
    famille_group["Pieces_KO"] = famille_group["Pieces_KO"].astype(int)
    famille_group = famille_group.rename(
        columns={
            "Famille": "Famille",
            "Saisies": "Saisies",
            "Heures_produites": "Heures produites",
            "Pieces_KO": "Pièces KO",
        }
    )
    famille_table = famille_group.to_dict(orient="records")

    # --- Dernières saisies (preview) ---
    preview = df.head(50).to_dict(orient="records")

    return {
        "total_saisies": total_saisies,
        "moyenne_heures": moyenne_heures,
        "total_rebuts": total_rebuts,
        "causes_labels": json.dumps(causes_labels),
        "causes_values": json.dumps(causes_values),
        "famille_table": famille_table,
        "preview": preview,
        "colonnes": COLONNES_REQUISES,
    }


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for("index"))

    file = request.files["file"]

    if file.filename == "":
        flash("Aucun fichier sélectionné.", "danger")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Format non supporté. Veuillez uploader un fichier .xlsx ou .xls", "danger")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(filepath)

    try:
        data = process_excel(filepath)
    except ValueError as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Erreur lors de la lecture du fichier : {e}", "danger")
        return redirect(url_for("index"))
    finally:
        # Remove file after processing
        if os.path.exists(filepath):
            os.remove(filepath)

    return render_template("dashboard.html", **data, filename=filename)


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, port=5000)
