import os
import json
import re
import hashlib
import datetime
import pandas as pd
from flask import (Flask, render_template, request, redirect, url_for,
                   flash, make_response, session, abort)
from werkzeug.utils import secure_filename
from flask_login import (LoginManager, login_required,
                         login_user, logout_user, current_user)
from flask_wtf.csrf import CSRFProtect
from database import db, User, Analysis, Consent, MAX_FREE_ANALYSES, CGU_VERSION

# ─────────────────────────────────────────────
#  BASE DE CONNAISSANCES TRS
#  Chargée une seule fois au démarrage.
# ─────────────────────────────────────────────
_KNOWLEDGE_PATH = os.path.join(os.path.dirname(__file__), "knowledge", "trs_knowledge.json")
try:
    with open(_KNOWLEDGE_PATH, encoding="utf-8") as _f:
        TRS_KNOWLEDGE = json.load(_f)
    _KNOWLEDGE_INDEX = {ind["id"]: ind for ind in TRS_KNOWLEDGE.get("indicateurs", [])}
except Exception:
    TRS_KNOWLEDGE = {"indicateurs": []}
    _KNOWLEDGE_INDEX = {}

# ─────────────────────────────────────────────
#  SWITCH IA : passer à True + définir
#  ANTHROPIC_API_KEY en variable d'env pour
#  basculer sur l'API Claude (1 ligne à changer)
# ─────────────────────────────────────────────
USE_CLAUDE_API = False

# ─────────────────────────────────────────────
#  PARAMÈTRES CONFIGURABLES
#  HEURES_POSTE    : durée théorique d'un poste (calcul Disponibilité)
#  CADENCE_NOMINALE: pièces/heure à pleine vitesse (calcul Performance)
#                    → None si non connue : Performance affichée "—"
# ─────────────────────────────────────────────
HEURES_POSTE     = 8
CADENCE_NOMINALE = None  # ex. : 120  pour 120 pièces/heure

# ─────────────────────────────────────────────
#  OPTIMISATION COÛT IA
#  Au-delà de SAMPLE_THRESHOLD lignes, on
#  échantillonne SAMPLE_SIZE lignes aléatoires
#  pour les stats (résultat toujours représentatif).
# ─────────────────────────────────────────────
SAMPLE_THRESHOLD = 1000
SAMPLE_SIZE      = 200

# ─────────────────────────────────────────────
#  COLLECTE ANONYMISÉE (version gratuite)
#  Stats agrégées uniquement, jamais de données
#  brutes ni de nom de fichier.
# ─────────────────────────────────────────────
DATA_COLLECT_DIR = os.path.join(os.path.dirname(__file__), "data", "collected")

app = Flask(__name__)
# En production, définir SECRET_KEY dans les variables d'environnement.
app.secret_key = os.environ.get("SECRET_KEY", "perfcnc_dev_secret_change_in_prod")

# ── Base de données SQLite ──
_DB_PATH = os.path.join(os.path.dirname(__file__), "instance", "perfcnc.db")
os.makedirs(os.path.dirname(_DB_PATH), exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", f"sqlite:///{_DB_PATH}")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db.init_app(app)

# ── Flask-Login ──
login_manager = LoginManager(app)
login_manager.login_view  = "login"
login_manager.login_message = "Connectez-vous pour accéder à cette page."
login_manager.login_message_category = "danger"

@login_manager.user_loader
def _load_user(user_id):
    return User.query.get(int(user_id))

# ── CSRF ──
csrf = CSRFProtect(app)

from flask_mail import Mail, Message as MailMessage

app.config["MAIL_SERVER"]   = os.environ.get("MAIL_SERVER",   "smtp.gmail.com")
app.config["MAIL_PORT"]     = int(os.environ.get("MAIL_PORT", "587"))
app.config["MAIL_USE_TLS"]  = True
app.config["MAIL_USERNAME"] = os.environ.get("MAIL_USERNAME", "")
app.config["MAIL_PASSWORD"] = os.environ.get("MAIL_PASSWORD", "")
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "contact@perfcnc.com")
mail = Mail(app)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB max

# Colonne optionnelle pour le Taux qualité
COLONNE_TOTAL_PIECES = "Total pièces"

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


def _get_ia_context(stats: dict) -> str:
    """
    Sélectionne depuis trs_knowledge.json les définitions et leviers
    pertinents selon les seuils des données analysées.
    Injecte toujours les valeurs de référence TRS, plus le contexte
    Qualité si le taux de rebuts dépasse 5 %, et Disponibilité si
    plusieurs causes d'arrêt distinctes sont présentes.
    """
    lines = []
    total_saisies = stats.get("total_saisies", 0)
    total_rebuts  = stats.get("total_rebuts", 0)
    causes        = stats.get("causes", {})

    # Taux de rebuts élevé → contexte Taux qualité
    if total_saisies > 0 and total_rebuts / max(total_saisies, 1) > 0.05:
        ind = _KNOWLEDGE_INDEX.get("Qualité")
        if ind:
            leviers = ", ".join(ind.get("leviers_amelioration", [])[:3])
            lines.append(
                f"[QUALITÉ] {ind['definition_operateur']} "
                f"Formule : {ind['formule_simple']}. "
                f"Leviers : {leviers}."
            )

    # Plusieurs causes d'arrêt → contexte Disponibilité
    if len(causes) >= 3:
        ind = _KNOWLEDGE_INDEX.get("Disponibilité")
        if ind:
            leviers = ", ".join(ind.get("leviers_amelioration", [])[:3])
            lines.append(
                f"[DISPONIBILITÉ] {ind['definition_operateur']} "
                f"Formule : {ind['formule_simple']}. "
                f"Leviers : {leviers}."
            )

    # Toujours : valeurs de référence TRS
    ind_trs = _KNOWLEDGE_INDEX.get("TRS")
    if ind_trs:
        ref = ind_trs.get("valeurs_reference", {})
        lines.append(
            f"[TRS RÉFÉRENCE] Mauvais : {ref.get('mauvais', '<60')} % — "
            f"Acceptable : {ref.get('acceptable', '60-75')} % — "
            f"Bon : {ref.get('bon', '75-85')} % — "
            f"Excellent : {ref.get('excellent', '>85')} %."
        )

    return "\n".join(lines)


def _build_prompt(stats: dict) -> str:
    stats_str = (
        f"total_saisies={stats['total_saisies']}, "
        f"moyenne_heures={stats['moyenne_heures']}, "
        f"total_rebuts={stats['total_rebuts']}, "
        f"causes={stats['causes']}"
    )
    context = _get_ia_context(stats)
    if context:
        return (
            "Tu es un expert en performance d'atelier d'usinage. "
            f"Contexte industriel de référence :\n{context}\n\n"
            f"Voici les données du mois : {stats_str}. "
            "En tenant compte de ce contexte, donne exactement 3 insights "
            "en JSON uniquement, sans texte autour : "
            '{"probleme_principal": "...", "tendance": "...", "recommandation": "..."}'
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
#  ÉVOLUTION TRS — SUIVI TEMPOREL
# ─────────────────────────────────────────────
def compute_trs_timeline(df: "pd.DataFrame") -> dict:
    """
    Calcule le TRS par jour, semaine et mois sur le DataFrame complet
    (avant échantillonnage, pour conserver la fidélité temporelle).

    Colonnes utilisées :
      - Date, Heures produites, Pièces KO (obligatoires)
      - Total pièces                       (optionnel → Qualité)
      - Machine                            (optionnel → ventilation par machine)
      - Famille                            (toujours présente → ventilation par famille)

    Performance calculée uniquement si CADENCE_NOMINALE ≠ None ET
    colonne « Total pièces » présente (pour avoir les pièces réelles).
    """
    wdf = df.copy()
    # Les dates sont déjà des chaînes "%d/%m/%Y" à ce stade
    wdf["_dt"] = pd.to_datetime(wdf["Date"], format="%d/%m/%Y", errors="coerce")
    wdf = wdf.dropna(subset=["_dt"])

    if wdf.empty:
        return {
            "jour": [], "semaine": [], "mois": [],
            "par_machine": {}, "par_famille": {},
            "has_machine": False,
            "has_cadence": CADENCE_NOMINALE is not None,
            "has_total_pieces": False,
            "machines": [], "familles": [],
        }

    wdf["_jour"]    = wdf["_dt"].dt.strftime("%Y-%m-%d")
    wdf["_semaine"] = wdf["_dt"].dt.strftime("%G-S%V")   # semaine ISO
    wdf["_mois"]    = wdf["_dt"].dt.strftime("%Y-%m")

    has_machine   = "Machine" in wdf.columns
    has_total_p   = COLONNE_TOTAL_PIECES in wdf.columns
    has_cadence   = CADENCE_NOMINALE is not None

    def _stats(gdf):
        n      = len(gdf)
        h_ouv  = n * HEURES_POSTE
        h_prod = float(gdf["Heures produites"].sum())
        ko     = float(gdf["Pièces KO"].sum())

        dispo = round(h_prod / h_ouv * 100, 1) if h_ouv > 0 else None

        perf    = None
        qualite = None
        if has_total_p:
            total_p = float(pd.to_numeric(
                gdf[COLONNE_TOTAL_PIECES], errors="coerce").fillna(0).sum())
            if total_p > 0:
                qualite = round((total_p - ko) / total_p * 100, 1)
                if has_cadence and h_prod > 0:
                    pieces_theo = h_prod * CADENCE_NOMINALE
                    perf = round(min(total_p / pieces_theo * 100, 100), 1)

        if dispo is not None and perf is not None and qualite is not None:
            trs = round(dispo * perf * qualite / 10000, 1)
        elif dispo is not None and qualite is not None:
            trs = round(dispo * qualite / 100, 1)
        else:
            trs = dispo

        return {
            "disponibilite": dispo,
            "performance":   perf,
            "qualite":       qualite,
            "trs":           trs,
            "nb_postes":     n,
        }

    def _series(src, col):
        rows = []
        for periode, grp in src.groupby(col, sort=True):
            s = _stats(grp)
            s["periode"] = str(periode)
            rows.append(s)
        return rows

    machines = sorted(wdf["Machine"].dropna().astype(str).unique().tolist()) if has_machine else []
    familles = sorted(wdf["Famille"].dropna().astype(str).unique().tolist())

    result = {
        "has_machine":      has_machine,
        "has_cadence":      has_cadence,
        "has_total_pieces": has_total_p,
        "machines":         machines,
        "familles":         familles,
    }

    for gran, col in [("jour", "_jour"), ("semaine", "_semaine"), ("mois", "_mois")]:
        result[gran] = _series(wdf, col)

    result["par_machine"] = {
        str(m): {g: _series(mdf, c)
                 for g, c in [("jour","_jour"),("semaine","_semaine"),("mois","_mois")]}
        for m, mdf in (wdf.groupby("Machine") if has_machine else [])
    }

    result["par_famille"] = {
        (str(f) if pd.notna(f) else "Inconnue"): {
            g: _series(fdf, c)
            for g, c in [("jour","_jour"),("semaine","_semaine"),("mois","_mois")]
        }
        for f, fdf in wdf.groupby("Famille", dropna=False)
    }

    return result


def _compute_trs_alerte(trs_timeline: dict) -> dict | None:
    """
    Retourne le niveau d'alerte basé sur le TRS de la dernière période.
    Priorité : mois → semaine → jour.
    """
    for gran in ("mois", "semaine", "jour"):
        data = trs_timeline.get(gran, [])
        if data:
            last = data[-1]
            trs  = last.get("trs")
            if trs is not None:
                if trs < 60:
                    return {
                        "niveau": "rouge", "trs": trs, "periode": last["periode"],
                        "message": (
                            f"TRS critique : {trs} % sur la dernière période "
                            f"({last['periode']}). Analysez les causes d'arrêt en priorité."
                        ),
                    }
                elif trs < 75:
                    return {
                        "niveau": "orange", "trs": trs, "periode": last["periode"],
                        "message": (
                            f"TRS acceptable : {trs} % sur la dernière période "
                            f"({last['periode']}). Des actions ciblées peuvent améliorer ce score."
                        ),
                    }
                else:
                    return {
                        "niveau": "vert", "trs": trs, "periode": last["periode"],
                        "message": (
                            f"Bonne performance : TRS {trs} % sur la dernière période "
                            f"({last['periode']})."
                        ),
                    }
    return None


# ─────────────────────────────────────────────
#  INDICATEURS DE PERFORMANCE
# ─────────────────────────────────────────────
def compute_indicators(df: "pd.DataFrame") -> dict:
    """
    Calcule TRS, Disponibilité et Taux qualité à partir du DataFrame.
    Chaque ligne représente un poste de HEURES_POSTE heures théoriques.

    - Disponibilité = Σ(Heures produites) / (n × HEURES_POSTE) × 100
    - TRS           = Disponibilité × (Taux qualité / 100) si qualité dispo,
                      sinon = Disponibilité (modèle simplifié sans cadence)
    - Taux qualité  = (Σ Total pièces - Σ Pièces KO) / Σ Total pièces × 100
                      uniquement si la colonne "Total pièces" est présente.
    """
    n = len(df)
    heures_ouverture = n * HEURES_POSTE
    heures_produites = df["Heures produites"].sum()

    disponibilite = round((heures_produites / heures_ouverture) * 100, 1) if heures_ouverture > 0 else None

    # Taux qualité — nécessite la colonne optionnelle "Total pièces"
    taux_qualite = None
    if COLONNE_TOTAL_PIECES in df.columns:
        total_pieces = pd.to_numeric(df[COLONNE_TOTAL_PIECES], errors="coerce").fillna(0).sum()
        pieces_ko = df["Pièces KO"].sum()
        if total_pieces > 0:
            taux_qualite = round(((total_pieces - pieces_ko) / total_pieces) * 100, 1)

    # TRS : Disponibilité × Qualité si les deux sont dispo, sinon Disponibilité seule
    if disponibilite is not None and taux_qualite is not None:
        trs = round(disponibilite * (taux_qualite / 100), 1)
    else:
        trs = disponibilite  # modèle simplifié

    return {
        "trs": trs,
        "disponibilite": disponibilite,
        "taux_qualite": taux_qualite,
        "heures_ouverture_theoriques": round(heures_ouverture, 2),
        "heures_poste": HEURES_POSTE,
    }


# ─────────────────────────────────────────────
#  COLLECTE ANONYMISÉE
# ─────────────────────────────────────────────
def collect_anonymous_stats(stats: dict) -> None:
    """
    Sauvegarde les stats agrégées dans data/collected/ sous forme JSON.
    Appelée uniquement si USE_CLAUDE_API=False (version gratuite) et
    après consentement de l'utilisateur.

    Règles strictes :
      - Jamais le nom du fichier original
      - Jamais les dates réelles des saisies
      - Jamais les données brutes ligne par ligne
      - Uniquement des KPIs agrégés et des dicts cause→nb
    """
    if USE_CLAUDE_API:
        return  # version Pro : pas de collecte
    try:
        os.makedirs(DATA_COLLECT_DIR, exist_ok=True)
        ts     = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S_%f")
        salt   = os.urandom(8).hex()
        anon   = hashlib.sha256(f"{ts}{salt}".encode()).hexdigest()[:12]
        fname  = f"{ts}_{anon}.json"
        payload = {
            "total_saisies":  stats.get("total_saisies"),
            "moyenne_heures": stats.get("moyenne_heures"),
            "total_rebuts":   stats.get("total_rebuts"),
            "causes":         stats.get("causes", {}),
            "familles":       stats.get("familles", {}),
            "trs":            stats.get("trs"),
        }
        with open(os.path.join(DATA_COLLECT_DIR, fname), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # Ne jamais crasher l'app pour la collecte


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

    # Conserver les colonnes optionnelles si elles sont présentes
    _opt = [c for c in [COLONNE_TOTAL_PIECES, "Machine"] if c in df.columns]
    df = df[COLONNES_REQUISES + _opt].copy()

    df["Heures produites"] = pd.to_numeric(df["Heures produites"], errors="coerce").fillna(0)
    df["Pièces KO"] = pd.to_numeric(df["Pièces KO"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    # ── Évolution TRS — calculée sur données complètes avant échantillonnage ──
    trs_timeline = compute_trs_timeline(df)
    trs_alerte   = _compute_trs_alerte(trs_timeline)

    # ── Échantillonnage pour fichiers volumineux (IA + KPIs agrégés) ──
    original_size = len(df)
    sampled = original_size > SAMPLE_THRESHOLD
    if sampled:
        df = df.sample(n=SAMPLE_SIZE).reset_index(drop=True)

    # ── KPIs ──
    total_saisies = len(df)
    moyenne_heures = round(df["Heures produites"].mean(), 2)
    total_rebuts = int(df["Pièces KO"].sum())
    total_heures_produites = round(df["Heures produites"].sum(), 2)

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
    famille_dict       = dict(zip(famille_pie_labels, famille_pie_values))

    # ── Indicateurs de performance ──
    indicateurs = compute_indicators(df)

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
        "total_saisies":         total_saisies,
        "moyenne_heures":        moyenne_heures,
        "total_rebuts":          total_rebuts,
        "total_heures_produites": total_heures_produites,
        "causes_labels":      json.dumps(causes_labels),
        "causes_values":      json.dumps(causes_values),
        "causes_dict":        causes_dict,
        "famille_table":      famille_table,
        "famille_pie_labels": json.dumps(famille_pie_labels),
        "famille_pie_values": json.dumps(famille_pie_values),
        "famille_dict":       famille_dict,
        "preview":            preview,
        "colonnes":           COLONNES_REQUISES,
        "insights":           insights,
        "indicateurs":        indicateurs,
        "sampled":            sampled,
        "original_size":      original_size,
        "trs_timeline":       trs_timeline,
        "trs_alerte":         trs_alerte,
    }


# ─────────────────────────────────────────────
#  DÉMO — GÉNÉRATION DU FICHIER EXCEL EXEMPLE
# ─────────────────────────────────────────────
_DEMO_PATH = os.path.join(os.path.dirname(__file__), "data", "demo", "demo_atelier.xlsx")


def _generate_demo_file():
    """Crée data/demo/demo_atelier.xlsx avec 50 lignes réalistes si non existant."""
    import random
    from openpyxl import Workbook  # type: ignore

    os.makedirs(os.path.dirname(_DEMO_PATH), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Données"

    headers = [
        "Date", "Équipe", "Code CRMX", "Référence pièce", "Famille",
        "Cause arrêt", "Heures produites", "Pièces KO", "Total pièces",
    ]
    ws.append(headers)

    equipes      = ["Matin", "Après-midi", "Nuit"]
    codes        = ["CRM001", "CRM002", "CRM003", "CRM004", "CRM005"]
    refs         = [f"PX-{n}" for n in range(100, 320, 10)]
    familles     = ["Usinage", "Fraisage", "Tournage"]
    causes       = ["Panne machine", "Réglage", "Manque matière", "Pause", "Maintenance"]

    today  = datetime.date.today()
    start  = today - datetime.timedelta(days=56)

    rng = random.Random(42)  # reproductible
    for _ in range(50):
        offset    = rng.randint(0, 55)
        date_row  = start + datetime.timedelta(days=offset)
        total_p   = rng.randint(40, 120)
        pieces_ko = rng.randint(0, min(15, total_p))
        ws.append([
            date_row.strftime("%d/%m/%Y"),
            rng.choice(equipes),
            rng.choice(codes),
            rng.choice(refs),
            rng.choice(familles),
            rng.choice(causes),
            round(rng.uniform(4, 8), 1),
            pieces_ko,
            total_p,
        ])

    wb.save(_DEMO_PATH)


# Génère le fichier démo au démarrage si absent
if not os.path.exists(_DEMO_PATH):
    try:
        _generate_demo_file()
    except Exception:
        pass


# ─────────────────────────────────────────────
#  GÉNÉRATION PDF
# ─────────────────────────────────────────────
def generate_pdf(data: dict) -> bytes:
    """
    Génère un rapport PDF A4 à partir des données du dashboard.
    Retourne les bytes du PDF.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer,
        Table, TableStyle, HRFlowable,
    )
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from io import BytesIO
    import datetime

    # ── Palette ──
    NAVY     = colors.HexColor('#1F3864')
    NAVY_LT  = colors.HexColor('#e8eef8')
    ORANGE   = colors.HexColor('#ED7D31')
    ORA_LT   = colors.HexColor('#fdf0e8')
    ORA_DK   = colors.HexColor('#c96520')
    ORA_BD   = colors.HexColor('#f8d5b3')
    GRAY_BG  = colors.HexColor('#f4f6f9')
    GRAY_BD  = colors.HexColor('#dde3ec')
    TEXT     = colors.HexColor('#1a2333')
    MUTED    = colors.HexColor('#5a6a80')
    BLUE     = colors.HexColor('#2980b9')
    GREEN    = colors.HexColor('#1a7a4a')
    GRN_LT   = colors.HexColor('#e8f5ee')
    GRN_BD   = colors.HexColor('#b3dfc7')
    RED      = colors.HexColor('#c0392b')
    RED_LT   = colors.HexColor('#fdf0f0')
    RED_BD   = colors.HexColor('#f5c6c6')
    WHITE    = colors.white

    buffer = BytesIO()
    PAGE_W, _ = A4
    MARGIN   = 1.5 * cm
    USABLE_W = PAGE_W - 2 * MARGIN
    COL3     = USABLE_W / 3

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=2 * cm,
        title="Rapport PerfCNC",
    )

    # ── Styles ──
    def ps(name, **kw):
        defaults = dict(fontName='Helvetica', fontSize=9, textColor=TEXT, leading=13)
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    S_BODY     = ps('body')
    S_MUTED    = ps('muted', fontSize=8, textColor=MUTED)
    S_H2       = ps('h2', fontSize=12, fontName='Helvetica-Bold', textColor=NAVY,
                    spaceBefore=10, spaceAfter=5)
    S_TH       = ps('th', fontSize=8, fontName='Helvetica-Bold', textColor=WHITE)
    S_TH_C     = ps('th_c', fontSize=8, fontName='Helvetica-Bold', textColor=WHITE,
                    alignment=TA_CENTER)
    S_NUM      = ps('num', alignment=TA_CENTER)
    S_FOOTER   = ps('footer', fontSize=7, textColor=MUTED, alignment=TA_CENTER)

    def val_style(name, color):
        return ps(name, fontSize=20, fontName='Helvetica-Bold', textColor=color,
                  alignment=TA_CENTER, leading=24)

    def lbl_style(name):
        return ps(name, fontSize=7, textColor=MUTED, alignment=TA_CENTER)

    def sub_style(name):
        return ps(name, fontSize=7, textColor=MUTED, alignment=TA_CENTER, leading=10)

    story = []
    now = datetime.datetime.now().strftime('%d/%m/%Y à %H:%M')
    filename = data.get('filename', '')

    # ── Helpers ──
    def tbl_pad(extra=None):
        base = [
            ('TOPPADDING',    (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING',   (0,0), (-1,-1), 7),
            ('RIGHTPADDING',  (0,0), (-1,-1), 7),
        ]
        return base + (extra or [])

    def zebra(n_rows, even=WHITE, odd=GRAY_BG):
        return [('BACKGROUND', (0,i), (-1,i), even if i % 2 == 1 else odd)
                for i in range(1, n_rows)]

    # ════════════════════════════════
    #  EN-TÊTE
    # ════════════════════════════════
    hdr = Table([
        [
            Paragraph('PerfCNC', ps('logo', fontSize=22, fontName='Helvetica-Bold',
                                    textColor=ORANGE)),
            Paragraph('Rapport de performance atelier',
                      ps('hrt', fontSize=10, fontName='Helvetica-Bold', textColor=WHITE,
                         alignment=TA_RIGHT)),
        ],
        [
            Paragraph('Tableau de bord production CNC',
                      ps('hsub', fontSize=9, textColor=colors.HexColor('#b8c8e0'))),
            Paragraph(filename,
                      ps('hfn', fontSize=8, textColor=colors.HexColor('#7a9ac8'),
                         alignment=TA_RIGHT)),
        ],
        [
            Paragraph(f'Analyse du {now}',
                      ps('hdate', fontSize=8, textColor=colors.HexColor('#9ab8d8'))),
            Paragraph('', S_BODY),
        ],
    ], colWidths=[USABLE_W * 0.55, USABLE_W * 0.45])
    hdr.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), NAVY),
        ('TOPPADDING',    (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING',   (0,0), (-1,-1), 14),
        ('RIGHTPADDING',  (0,0), (-1,-1), 14),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.5 * cm))

    # ════════════════════════════════
    #  KPI CARDS
    # ════════════════════════════════
    story.append(Paragraph('Indicateurs clés', S_H2))

    kpi_specs = [
        (data.get('total_saisies', '—'),  'Total saisies',               NAVY_LT, RED_BD,  NAVY),
        (data.get('moyenne_heures', '—'), 'Heures produites (moy.)',      ORA_LT,  ORA_BD,  ORANGE),
        (data.get('total_rebuts', '—'),   'Total rebuts (Pièces KO)',     RED_LT,  RED_BD,  RED),
    ]
    kpi_cells = []
    for i, (val, lbl, bg, bd, color) in enumerate(kpi_specs):
        cell = Table(
            [[Paragraph(str(val), val_style(f'kv{i}', color))],
             [Paragraph(lbl,      lbl_style(f'kl{i}'))]],
            colWidths=[COL3 - 8],
        )
        cell.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), bg),
            ('BOX',        (0,0), (-1,-1), 0.5, bd),
            ('TOPPADDING',    (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        kpi_cells.append(cell)

    kpi_tbl = Table([kpi_cells], colWidths=[COL3, COL3, COL3])
    kpi_tbl.setStyle(TableStyle([
        ('LEFTPADDING',  (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING',   (0,0), (-1,-1), 0),
        ('BOTTOMPADDING',(0,0), (-1,-1), 0),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  INDICATEURS DE PERFORMANCE
    # ════════════════════════════════
    indicateurs = data.get('indicateurs') or {}
    if indicateurs:
        story.append(Paragraph('Indicateurs de performance', S_H2))

        def fmt(v):
            return f'{v} %' if v is not None else 'N/D'

        perf_specs = [
            (fmt(indicateurs.get('trs')),          'TRS',          'Taux de Rendement Synthétique',    NAVY),
            (fmt(indicateurs.get('disponibilite')),'Disponibilité','Heures prod. / Heures ouverture',  BLUE),
            (fmt(indicateurs.get('taux_qualite')), 'Taux qualité', '(Total pièces − KO) / Total p.',   GREEN),
        ]
        perf_cells = []
        for i, (val, lbl, sub, color) in enumerate(perf_specs):
            cell = Table(
                [[Paragraph(val, val_style(f'pv{i}', color))],
                 [Paragraph(lbl, lbl_style(f'pl{i}'))],
                 [Paragraph(sub, sub_style(f'ps{i}'))]],
                colWidths=[COL3 - 8],
            )
            cell.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), GRAY_BG),
                ('BOX',        (0,0), (-1,-1), 0.4, GRAY_BD),
                ('TOPPADDING',    (0,0), (-1,-1), 10),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ]))
            perf_cells.append(cell)

        perf_tbl = Table([perf_cells], colWidths=[COL3, COL3, COL3])
        perf_tbl.setStyle(TableStyle([
            ('LEFTPADDING',  (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING',   (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0), (-1,-1), 0),
        ]))
        story.append(perf_tbl)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  CAUSES D'ARRÊT
    # ════════════════════════════════
    causes_labels = data.get('causes_labels_list', [])
    causes_values = data.get('causes_values_list', [])
    if causes_labels:
        story.append(Paragraph("Causes d'arrêt", S_H2))
        total_c = sum(causes_values) or 1
        rows = [[Paragraph("Cause d'arrêt", S_TH),
                 Paragraph('Occurrences', S_TH_C),
                 Paragraph('Part (%)', S_TH_C)]]
        for lbl, val in zip(causes_labels, causes_values):
            rows.append([
                Paragraph(str(lbl), S_BODY),
                Paragraph(str(val), S_NUM),
                Paragraph(f'{round(val / total_c * 100, 1)} %', S_NUM),
            ])
        ct = Table(rows, colWidths=[USABLE_W * 0.62, USABLE_W * 0.19, USABLE_W * 0.19])
        ct.setStyle(TableStyle(
            [('BACKGROUND', (0,0), (-1,0), NAVY),
             ('BOX',        (0,0), (-1,-1), 0.5, GRAY_BD),
             ('INNERGRID',  (0,1), (-1,-1), 0.3, GRAY_BD)]
            + tbl_pad()
            + zebra(len(rows))
        ))
        story.append(ct)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  TABLEAU FAMILLES
    # ════════════════════════════════
    famille_rows = data.get('famille_table', [])
    if famille_rows:
        story.append(Paragraph('Résumé par famille', S_H2))
        rows = [[Paragraph('Famille', S_TH),
                 Paragraph('Saisies', S_TH_C),
                 Paragraph('Heures prod.', S_TH_C),
                 Paragraph('Pièces KO', S_TH_C)]]
        for row in famille_rows:
            rows.append([
                Paragraph(str(row.get('Famille') or '—'), S_BODY),
                Paragraph(str(row.get('Saisies', '')),          S_NUM),
                Paragraph(str(row.get('Heures produites', '')), S_NUM),
                Paragraph(str(row.get('Pièces KO', '')),        S_NUM),
            ])
        ft = Table(rows, colWidths=[USABLE_W*0.4, USABLE_W*0.2,
                                    USABLE_W*0.2, USABLE_W*0.2])
        ft.setStyle(TableStyle(
            [('BACKGROUND', (0,0), (-1,0), NAVY),
             ('BOX',        (0,0), (-1,-1), 0.5, GRAY_BD),
             ('INNERGRID',  (0,1), (-1,-1), 0.3, GRAY_BD)]
            + tbl_pad()
            + zebra(len(rows))
        ))
        story.append(ft)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  INSIGHTS IA
    # ════════════════════════════════
    insights = data.get('insights')
    if isinstance(insights, dict):
        story.append(Paragraph('Analyse IA', S_H2))
        ia_specs = [
            ('Problème principal', insights.get('probleme_principal', ''),
             RED_LT, RED_BD, RED),
            ('Tendance détectée',  insights.get('tendance', ''),
             ORA_LT, ORA_BD, ORA_DK),
            ('Recommandation',     insights.get('recommandation', ''),
             GRN_LT, GRN_BD, GREEN),
        ]
        ia_cells = []
        for i, (title, text, bg, bd, color) in enumerate(ia_specs):
            cell = Table(
                [[Paragraph(title, ps(f'it{i}', fontSize=8, fontName='Helvetica-Bold',
                                      textColor=color))],
                 [Paragraph(text,  ps(f'ib{i}', fontSize=8, leading=12))]],
                colWidths=[COL3 - 8],
            )
            cell.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), bg),
                ('BOX',        (0,0), (-1,-1), 0.7, bd),
                ('TOPPADDING',    (0,0), (-1,-1), 9),
                ('BOTTOMPADDING', (0,0), (-1,-1), 9),
                ('LEFTPADDING',   (0,0), (-1,-1), 7),
                ('RIGHTPADDING',  (0,0), (-1,-1), 7),
                ('VALIGN',        (0,0), (-1,-1), 'TOP'),
            ]))
            ia_cells.append(cell)
        ia_tbl = Table([ia_cells], colWidths=[COL3, COL3, COL3])
        ia_tbl.setStyle(TableStyle([
            ('LEFTPADDING',  (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING',   (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0), (-1,-1), 0),
        ]))
        story.append(ia_tbl)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  TRS PAR SEMAINE
    # ════════════════════════════════
    trs_timeline = data.get("trs_timeline", {})
    semaines = trs_timeline.get("semaine", [])
    if semaines:
        story.append(Paragraph("TRS par semaine", S_H2))
        trs_hdr = [
            Paragraph("Semaine", S_TH_C),
            Paragraph("Disponibilité", S_TH_C),
            Paragraph("Performance", S_TH_C),
            Paragraph("Qualité", S_TH_C),
            Paragraph("TRS", S_TH_C),
            Paragraph("Postes", S_TH_C),
        ]
        trs_rows = [trs_hdr]
        for row in semaines[-12:]:  # last 12 weeks max
            def _fmt_pct(v):
                return f"{v:.1f}%" if v is not None else "—"
            trs_rows.append([
                Paragraph(str(row.get("periode", "")), S_NUM),
                Paragraph(_fmt_pct(row.get("disponibilite")), S_NUM),
                Paragraph(_fmt_pct(row.get("performance")), S_NUM),
                Paragraph(_fmt_pct(row.get("taux_qualite")), S_NUM),
                Paragraph(_fmt_pct(row.get("trs")), S_NUM),
                Paragraph(str(row.get("nb_postes", "")), S_NUM),
            ])
        trs_tbl = Table(trs_rows, colWidths=[USABLE_W*0.22, USABLE_W*0.17, USABLE_W*0.17, USABLE_W*0.13, USABLE_W*0.13, USABLE_W*0.13])
        trs_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,0),  NAVY),
            ('TEXTCOLOR',     (0,0), (-1,0),  WHITE),
            ('GRID',          (0,0), (-1,-1), 0.4, GRAY_BD),
            ('ALIGN',         (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
            *tbl_pad(),
            *zebra(len(trs_rows)),
        ]))
        story.append(trs_tbl)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  COMPARAISON SEMAINE
    # ════════════════════════════════
    week_compare = data.get("week_compare", {})
    tw = week_compare.get("this_week")
    lw = week_compare.get("last_week")
    if tw and lw:
        story.append(Paragraph("Comparaison semaine", S_H2))
        def _fmt_p(v): return f"{v:.1f}%" if v is not None else "—"
        def _delta(a, b):
            if a is None or b is None: return "—"
            d = a - b
            sign = "+" if d > 0 else ""
            return f"{sign}{d:.1f}pts"
        cmp_data = [
            [Paragraph("Indicateur", S_TH), Paragraph(week_compare.get("this_week_label","S. courante"), S_TH_C), Paragraph(week_compare.get("last_week_label","S. précédente"), S_TH_C), Paragraph("Variation", S_TH_C)],
            [Paragraph("TRS", S_BODY), Paragraph(_fmt_p(tw.get("trs")), S_NUM), Paragraph(_fmt_p(lw.get("trs")), S_NUM), Paragraph(_delta(tw.get("trs"), lw.get("trs")), S_NUM)],
            [Paragraph("Disponibilité", S_BODY), Paragraph(_fmt_p(tw.get("disponibilite")), S_NUM), Paragraph(_fmt_p(lw.get("disponibilite")), S_NUM), Paragraph(_delta(tw.get("disponibilite"), lw.get("disponibilite")), S_NUM)],
            [Paragraph("Qualité", S_BODY), Paragraph(_fmt_p(tw.get("taux_qualite")), S_NUM), Paragraph(_fmt_p(lw.get("taux_qualite")), S_NUM), Paragraph(_delta(tw.get("taux_qualite"), lw.get("taux_qualite")), S_NUM)],
        ]
        cmp_tbl = Table(cmp_data, colWidths=[USABLE_W*0.3, USABLE_W*0.23, USABLE_W*0.23, USABLE_W*0.24])
        cmp_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,0),  NAVY),
            ('TEXTCOLOR',     (0,0), (-1,0),  WHITE),
            ('GRID',          (0,0), (-1,-1), 0.4, GRAY_BD),
            ('ALIGN',         (1,0), (-1,-1), 'CENTER'),
            ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
            *tbl_pad(),
            *zebra(len(cmp_data)),
        ]))
        story.append(cmp_tbl)
        story.append(Spacer(1, 0.4 * cm))

    # ════════════════════════════════
    #  PIED DE PAGE
    # ════════════════════════════════
    story.append(HRFlowable(width='100%', thickness=0.5, color=GRAY_BD, spaceAfter=5))
    story.append(Paragraph(
        f'Généré par PerfCNC — perfcnc.com &nbsp;·&nbsp; {now}',
        S_FOOTER,
    ))

    doc.build(story)
    return buffer.getvalue()


# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html",
                           cgu_accepted=session.get("cgu_accepted", False))


@app.route("/cgu")
def cgu():
    return render_template("cgu.html")


@app.route("/glossaire")
def glossaire():
    return render_template("glossaire.html",
                           indicateurs=TRS_KNOWLEDGE.get("indicateurs", []))


@app.route("/exercices")
def exercices():
    return render_template("exercices.html")


@app.route("/comment-ca-marche")
def comment_ca_marche():
    return render_template("how_it_works.html")


@app.route("/download-pdf", methods=["POST"])
def download_pdf():
    try:
        data = {
            "filename":          request.form.get("filename", "rapport"),
            "total_saisies":     request.form.get("total_saisies", "—"),
            "moyenne_heures":    request.form.get("moyenne_heures", "—"),
            "total_rebuts":      request.form.get("total_rebuts", "—"),
            "indicateurs":       json.loads(request.form.get("indicateurs", "{}")),
            "causes_labels_list":json.loads(request.form.get("causes_labels", "[]")),
            "causes_values_list":json.loads(request.form.get("causes_values", "[]")),
            "famille_table":     json.loads(request.form.get("famille_table", "[]")),
            "insights":          json.loads(request.form.get("insights", "null")),
            "trs_timeline":      json.loads(request.form.get("trs_timeline", "{}")),
            "week_compare":      json.loads(request.form.get("week_compare", "{}")),
        }
        pdf_bytes = generate_pdf(data)
    except Exception as e:
        flash(f"Erreur lors de la génération du PDF : {e}", "danger")
        return redirect(url_for("index"))

    safe_name = data["filename"].rsplit(".", 1)[0]
    response = make_response(pdf_bytes)
    response.headers["Content-Type"] = "application/pdf"
    response.headers["Content-Disposition"] = (
        f'attachment; filename="PerfCNC_{safe_name}.pdf"'
    )
    return response


@app.route("/upload", methods=["POST"])
def upload():
    # ── Vérification consentement CGU ──
    if not session.get("cgu_accepted"):
        if request.form.get("cgu_consent") == "on":
            session["cgu_accepted"] = True
        else:
            flash("Vous devez accepter les CGU pour utiliser PerfCNC.", "danger")
            return redirect(url_for("index"))

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
        import traceback
        print(traceback.format_exc())
        flash(f"Erreur lors de la lecture du fichier : {e}", "danger")
        return redirect(url_for("index"))
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

    # ── Collecte anonymisée (version gratuite, utilisateurs non connectés ou plan free) ──
    _is_free = not current_user.is_authenticated or current_user.plan == "free"
    if _is_free:
        collect_anonymous_stats({
            "total_saisies":  data["total_saisies"],
            "moyenne_heures": data["moyenne_heures"],
            "total_rebuts":   data["total_rebuts"],
            "causes":         data["causes_dict"],
            "familles":       data["famille_dict"],
            "trs":            data["indicateurs"].get("trs"),
        })

    # ── Sauvegarde dans l'historique (utilisateurs connectés) ──
    if current_user.is_authenticated:
        if current_user.can_save_analysis:
            try:
                save_data = {k: v for k, v in data.items() if k != "preview"}
                analysis = Analysis(
                    user_id    = current_user.id,
                    filename   = filename,
                    stats_json = json.dumps(save_data, ensure_ascii=False,
                                            default=_json_default),
                )
                db.session.add(analysis)
                db.session.commit()
                rem = current_user.analyses_remaining
                if rem is not None:
                    flash(f"Analyse sauvegardée. ({current_user.analyses_count}/{MAX_FREE_ANALYSES})", "success")
            except Exception:
                pass  # Ne jamais crasher l'app pour la sauvegarde
        else:
            flash(
                f"Limite de {MAX_FREE_ANALYSES} analyses atteinte. "
                "Supprimez une analyse dans votre historique pour en sauvegarder une nouvelle.",
                "danger",
            )

    # ── Week comparison ──
    week_compare = {}
    if current_user.is_authenticated:
        all_analyses = current_user.analyses.all()
        merged = _merge_timelines(all_analyses)
        current_timeline = data.get("trs_timeline", {})
        for gran in ("jour", "semaine", "mois"):
            for entry in current_timeline.get(gran, []):
                p = entry.get("periode", "")
                existing_list = merged.get(gran, [])
                exists = next((e for e in existing_list if e.get("periode") == p), None)
                if not exists:
                    merged.setdefault(gran, []).append(entry)
        week_compare = _week_comparison(merged)
    else:
        week_compare = _week_comparison(data.get("trs_timeline", {}))

    return render_template("dashboard.html", **data, filename=filename, week_compare=week_compare)


# ─────────────────────────────────────────────
#  AUTHENTIFICATION
# ─────────────────────────────────────────────
@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        email     = (request.form.get("email") or "").strip().lower()
        password  = request.form.get("password") or ""
        password2 = request.form.get("password2") or ""
        plan      = request.form.get("plan", "free")
        cgu_ok    = request.form.get("cgu_consent") == "on"

        error = None
        if not email or "@" not in email:
            error = "Adresse e-mail invalide."
        elif len(password) < 8:
            error = "Le mot de passe doit comporter au moins 8 caractères."
        elif password != password2:
            error = "Les deux mots de passe ne correspondent pas."
        elif not cgu_ok:
            error = "Vous devez accepter les CGU pour créer un compte."
        elif User.query.filter_by(email=email).first():
            error = "Cette adresse e-mail est déjà utilisée."

        if error:
            flash(error, "danger")
            return render_template("register.html", email=email)

        user = User(email=email, plan=plan if plan in ("free", "pro") else "free")
        user.set_password(password)
        consent = Consent(version_cgu=CGU_VERSION)
        user.consent = consent
        db.session.add(user)
        db.session.commit()

        login_user(user)
        session["cgu_accepted"] = True
        flash("Compte créé avec succès. Bienvenue sur PerfCNC !", "success")
        return redirect(url_for("index"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        email    = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""

        user = User.query.filter_by(email=email).first()
        if user and user.check_password(password):
            login_user(user, remember=True)
            session["cgu_accepted"] = True
            next_page = request.args.get("next")
            return redirect(next_page or url_for("index"))
        else:
            flash("E-mail ou mot de passe incorrect.", "danger")
            return render_template("login.html", email=email)

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Vous avez été déconnecté.", "success")
    return redirect(url_for("index"))


# ─────────────────────────────────────────────
#  HISTORIQUE
# ─────────────────────────────────────────────
def _merge_timelines(analyses):
    """Merge les trs_timeline de plusieurs analyses par période (garde la dernière)."""
    buckets = {"jour": {}, "semaine": {}, "mois": {}}
    for a in analyses:
        try:
            timeline = a.stats.get("trs_timeline", {})
            for gran in ("jour", "semaine", "mois"):
                for entry in timeline.get(gran, []):
                    p = entry.get("periode", "")
                    existing = buckets[gran].get(p)
                    if existing is None or entry.get("nb_postes", 0) > existing.get("nb_postes", 0):
                        buckets[gran][p] = entry
        except Exception:
            continue
    return {
        gran: sorted(buckets[gran].values(), key=lambda x: x.get("periode", ""))
        for gran in ("jour", "semaine", "mois")
    }


def _week_comparison(merged):
    today     = datetime.date.today()
    this_lbl  = today.strftime("%G-S%V")
    last_lbl  = (today - datetime.timedelta(weeks=1)).strftime("%G-S%V")
    semaines  = merged.get("semaine", [])
    return {
        "this_week":       next((r for r in semaines if r["periode"] == this_lbl), None),
        "last_week":       next((r for r in semaines if r["periode"] == last_lbl), None),
        "this_week_label": this_lbl,
        "last_week_label": last_lbl,
    }


@app.route("/historique")
@login_required
def historique():
    analyses        = current_user.analyses.all()
    merged_timeline = _merge_timelines(analyses) if analyses else {}
    week_compare    = _week_comparison(merged_timeline) if merged_timeline else {}
    return render_template("historique.html",
                           analyses        = analyses,
                           merged_timeline = merged_timeline,
                           week_compare    = week_compare)


@app.route("/analyse/<int:analysis_id>")
@login_required
def view_analysis(analysis_id):
    analysis = Analysis.query.get_or_404(analysis_id)
    if analysis.user_id != current_user.id:
        abort(403)
    stats = analysis.stats
    stats["preview"]      = []
    stats["from_history"] = True
    all_analyses = current_user.analyses.all()
    merged = _merge_timelines(all_analyses)
    week_compare = _week_comparison(merged)
    return render_template("dashboard.html", **stats, filename=analysis.filename, week_compare=week_compare)


@app.route("/analyse/<int:analysis_id>/supprimer", methods=["POST"])
@login_required
def delete_analysis(analysis_id):
    analysis = Analysis.query.get_or_404(analysis_id)
    if analysis.user_id != current_user.id:
        abort(403)
    db.session.delete(analysis)
    db.session.commit()
    flash("Analyse supprimée.", "success")
    return redirect(url_for("historique"))


# ─────────────────────────────────────────────
#  HELPER SÉRIALISATION JSON (numpy → Python)
# ─────────────────────────────────────────────
def _json_default(obj):
    """Convertit les types numpy/pandas non sérialisables."""
    try:
        import numpy as np  # type: ignore
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
    except ImportError:
        pass
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


@app.route("/demo")
def demo():
    session["cgu_accepted"] = True
    session["demo_mode"] = True
    try:
        data = process_excel(_DEMO_PATH)
    except Exception as e:
        flash(f"Erreur lors du chargement de la démo : {e}", "danger")
        return redirect(url_for("index"))

    collect_anonymous_stats({
        "total_saisies":  data["total_saisies"],
        "moyenne_heures": data["moyenne_heures"],
        "total_rebuts":   data["total_rebuts"],
        "causes":         data["causes_dict"],
        "familles":       data["famille_dict"],
        "trs":            data["indicateurs"].get("trs"),
    })

    week_compare = _week_comparison(data.get("trs_timeline", {}))

    return render_template(
        "dashboard.html",
        **data,
        filename="demo_atelier.xlsx",
        demo_mode=True,
        week_compare=week_compare,
    )


@app.route("/contact", methods=["GET", "POST"])
def contact():
    if request.method == "POST":
        name    = (request.form.get("name") or "").strip()
        email   = (request.form.get("email") or "").strip()
        message = (request.form.get("message") or "").strip()
        if not name or not email or not message:
            flash("Tous les champs sont obligatoires.", "danger")
            return render_template("contact.html", name=name, email=email, message=message)
        if "@" not in email:
            flash("Adresse e-mail invalide.", "danger")
            return render_template("contact.html", name=name, email=email, message=message)
        try:
            msg = MailMessage(
                subject   = f"[PerfCNC Contact] {name}",
                sender    = app.config["MAIL_USERNAME"] or CONTACT_EMAIL,
                recipients=[CONTACT_EMAIL],
                body      = f"De : {name} <{email}>\n\n{message}",
                reply_to  = email,
            )
            mail.send(msg)
            flash("Message envoyé avec succès ! Nous vous répondrons rapidement.", "success")
            return redirect(url_for("contact"))
        except Exception as e:
            flash(f"Erreur lors de l'envoi : vérifiez la configuration MAIL. ({e})", "danger")
    return render_template("contact.html")


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    with app.app_context():
        db.create_all()
    if not os.path.exists(_DEMO_PATH):
        _generate_demo_file()
    app.run(debug=True, port=5000)
