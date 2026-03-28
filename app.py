import os
import json
import re
import hashlib
import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, make_response, session
from werkzeug.utils import secure_filename

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
#  PARAMÈTRE CONFIGURABLE
#  Durée théorique d'un poste en heures.
#  Sert au calcul du TRS, de la Disponibilité
#  et des Heures d'ouverture théoriques.
# ─────────────────────────────────────────────
HEURES_POSTE = 8

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

    df = df[COLONNES_REQUISES].copy()

    df["Heures produites"] = pd.to_numeric(df["Heures produites"], errors="coerce").fillna(0)
    df["Pièces KO"] = pd.to_numeric(df["Pièces KO"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    # ── Échantillonnage pour fichiers volumineux ──
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
    }


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
        flash(f"Erreur lors de la lecture du fichier : {e}", "danger")
        return redirect(url_for("index"))
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

    # ── Collecte anonymisée (version gratuite uniquement) ──
    collect_anonymous_stats({
        "total_saisies":  data["total_saisies"],
        "moyenne_heures": data["moyenne_heures"],
        "total_rebuts":   data["total_rebuts"],
        "causes":         data["causes_dict"],
        "familles":       data["famille_dict"],
        "trs":            data["indicateurs"].get("trs"),
    })

    return render_template("dashboard.html", **data, filename=filename)


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, port=5000)
