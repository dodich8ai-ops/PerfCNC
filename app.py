import os
import json
import re
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

# ─────────────────────────────────────────────
#  SWITCH IA : passer à True + définir
#  ANTHROPIC_API_KEY en variable d'env pour
#  basculer sur l'API Claude (1 ligne à changer)
# ─────────────────────────────────────────────
USE_CLAUDE_API = False

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

# ─────────────────────────────────────────────
#  PROMPT IA
# ─────────────────────────────────────────────
PROMPT_TEMPLATE = (
    "Tu es un expert en performance d'atelier d'usinage. "
    "Voici les données du mois : {stats}. "
    "Donne exactement 3 insights en JSON uniquement, sans texte autour : "
    '{"probleme_principal": "...", "tendance": "...", "recommandation": "..."}'
)


def _build_prompt(stats: dict) -> str:
    stats_str = (
        f"total_saisies={stats['total_saisies']}, "
        f"moyenne_heures={stats['moyenne_heures']}, "
        f"total_rebuts={stats['total_rebuts']}, "
        f"causes={stats['causes']}"
    )
    return PROMPT_TEMPLATE.format(stats=stats_str)


def _parse_ia_response(text: str) -> dict | None:
    """Extract first JSON object found in the response text."""
    match = re.search(r"\{[^{}]+\}", text, re.DOTALL)
    if not match:
        return None
    try:
        data = json.loads(match.group())
        required = {"probleme_principal", "tendance", "recommandation"}
        if not required.issubset(data.keys()):
            return None
        return data
    except (json.JSONDecodeError, KeyError):
        return None


# ─────────────────────────────────────────────
#  BACKEND OLLAMA
# ─────────────────────────────────────────────
def _analyze_ollama(stats: dict) -> dict | None:
    try:
        import ollama  # type: ignore
        prompt = _build_prompt(stats)
        for model in ("mistral", "llama3", "llama3.2"):
            try:
                response = ollama.chat(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                )
                text = response["message"]["content"]
                result = _parse_ia_response(text)
                if result:
                    return result
            except ollama.ResponseError:
                continue  # model not installed, try next
        return None
    except Exception:
        return None


# ─────────────────────────────────────────────
#  BACKEND CLAUDE API
# ─────────────────────────────────────────────
def _analyze_claude(stats: dict) -> dict | None:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return None
    try:
        import anthropic  # type: ignore
        client = anthropic.Anthropic(api_key=api_key)
        prompt = _build_prompt(stats)
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        text = message.content[0].text
        return _parse_ia_response(text)
    except Exception:
        return None


# ─────────────────────────────────────────────
#  POINT D'ENTRÉE IA (switch automatique)
# ─────────────────────────────────────────────
def analyze_with_ollama(stats: dict) -> dict | None:
    """
    Analyse IA des stats atelier.
    Utilise Claude API si USE_CLAUDE_API=True et ANTHROPIC_API_KEY défini,
    sinon Ollama local.
    Retourne None sans crasher l'app si le service est indisponible.
    """
    if USE_CLAUDE_API and os.environ.get("ANTHROPIC_API_KEY"):
        return _analyze_claude(stats)
    return _analyze_ollama(stats)


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ─────────────────────────────────────────────
#  TRAITEMENT EXCEL
# ─────────────────────────────────────────────
def process_excel(filepath):
    df = pd.read_excel(filepath)
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in COLONNES_REQUISES if c not in df.columns]
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(missing)}")

    df = df[COLONNES_REQUISES].copy()

    df["Heures produites"] = pd.to_numeric(df["Heures produites"], errors="coerce").fillna(0)
    df["Pièces KO"] = pd.to_numeric(df["Pièces KO"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    # ── KPIs ──
    total_saisies = len(df)
    moyenne_heures = round(df["Heures produites"].mean(), 2)
    total_rebuts = int(df["Pièces KO"].sum())

    # ── Causes d'arrêt ──
    causes_series = (
        df["Cause arrêt"].fillna("Non renseigné").value_counts().sort_values(ascending=False)
    )
    causes_labels = causes_series.index.tolist()
    causes_values = causes_series.values.tolist()
    causes_dict = dict(zip(causes_labels, causes_values))

    # ── Familles (camembert + tableau) ──
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
    famille_table = famille_group.rename(
        columns={"Heures_produites": "Heures produites", "Pieces_KO": "Pièces KO"}
    ).to_dict(orient="records")

    famille_pie_labels = famille_group["Famille"].fillna("Inconnue").tolist()
    famille_pie_values = famille_group["Saisies"].tolist()

    # ── Aperçu ──
    preview = df.head(50).to_dict(orient="records")

    # ── Analyse IA ──
    stats_for_ia = {
        "total_saisies": total_saisies,
        "moyenne_heures": moyenne_heures,
        "total_rebuts": total_rebuts,
        "causes": causes_dict,
    }
    insights = analyze_with_ollama(stats_for_ia)

    return {
        "total_saisies": total_saisies,
        "moyenne_heures": moyenne_heures,
        "total_rebuts": total_rebuts,
        "causes_labels": json.dumps(causes_labels),
        "causes_values": json.dumps(causes_values),
        "famille_table": famille_table,
        "famille_pie_labels": json.dumps(famille_pie_labels),
        "famille_pie_values": json.dumps(famille_pie_values),
        "preview": preview,
        "colonnes": COLONNES_REQUISES,
        "insights": insights,
    }


# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────
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
        if os.path.exists(filepath):
            os.remove(filepath)

    return render_template("dashboard.html", **data, filename=filename)


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, port=5000)
