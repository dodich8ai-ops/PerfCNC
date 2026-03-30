import os, json, datetime
from flask import (Flask, render_template, request, redirect,
                   url_for, flash, session, abort)
from flask_login import (LoginManager, login_required,
                         login_user, logout_user, current_user)
from flask_wtf.csrf import CSRFProtect
from database import db, User, Watch, PriceHistory, MarketPriceCache, MAX_FREE_WATCHES, CGU_VERSION

# ── Switch IA ──
USE_CLAUDE_API = False  # True → API Claude, False → Ollama local

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "watchfolio_dev_secret_change_in_prod")

_DB_PATH = os.path.join(os.path.dirname(__file__), "instance", "watchfolio.db")
os.makedirs(os.path.dirname(_DB_PATH), exist_ok=True)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", f"sqlite:///{_DB_PATH}")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db.init_app(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Connectez-vous pour accéder à cette page."
login_manager.login_message_category = "danger"

@login_manager.user_loader
def _load_user(uid): return User.query.get(int(uid))

csrf = CSRFProtect(app)

# ── Jinja2 filter money ──
@app.template_filter("money")
def money_filter(value):
    try:
        v = float(value)
        return f"{v:,.0f}".replace(",", "\u202f") + " €"
    except (TypeError, ValueError):
        return "— €"

# ── Ollama IA ──
def conseil_ia(montre: Watch):
    """
    Appelle Ollama/Mistral pour générer un conseil marché.
    Retourne dict avec tendance, conseil, fourchette_prix.
    Retourne None si Ollama hors ligne.
    """
    try:
        import ollama
        prompt = (
            f"Tu es un expert en montres de luxe. Analyse cette montre :\n"
            f"- Marque : {montre.marque}\n"
            f"- Référence : {montre.reference}\n"
            f"- Année : {montre.annee or 'inconnue'}\n"
            f"- État : {montre.etat}\n"
            f"- Full set : {'Oui' if montre.full_set else 'Non'}\n"
            f"- Prix d'achat : {montre.prix_achat:.0f} €\n\n"
            f"Réponds UNIQUEMENT en JSON valide, sans texte autour :\n"
            '{"tendance": "hausse/stable/baisse", '
            '"conseil": "conseil de vente en 2 phrases max", '
            '"fourchette_prix": "ex: 8 000 – 12 000 €"}'
        )
        resp = ollama.chat(
            model="mistral",
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp["message"]["content"].strip()
        # Extract JSON even if surrounded by markdown
        import re
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            return json.loads(m.group())
    except Exception:
        pass
    return None


# ── Prix du marché (Chrono24) ──
def get_prix_marche(marque: str, reference: str):
    """
    Scrape les prix Chrono24 pour (marque, référence).
    Cache 24 h dans MarketPriceCache.
    Retourne dict ou None sans jamais crasher.
    """
    import re, statistics

    now    = datetime.datetime.utcnow()
    m_low  = marque.strip().lower()
    r_low  = reference.strip().lower()

    # ── Lecture du cache ──
    cached = MarketPriceCache.query.filter(
        db.func.lower(MarketPriceCache.marque)    == m_low,
        db.func.lower(MarketPriceCache.reference) == r_low,
    ).first()

    if cached and (now - cached.date_cache).total_seconds() < 86_400:
        return {
            "prix_min":    cached.prix_min,
            "prix_median": cached.prix_median,
            "prix_max":    cached.prix_max,
            "nb_annonces": cached.nb_annonces,
            "source":      "Chrono24",
            "date":        cached.date_cache.strftime("%d/%m/%Y"),
            "from_cache":  True,
        }

    # ── Scraping ──
    try:
        import requests
        from bs4 import BeautifulSoup

        query   = f"{marque.strip()} {reference.strip()}"
        url     = (
            "https://www.chrono24.fr/search/index.htm"
            f"?query={requests.utils.quote(query)}&dosearch=true&redirectToSearchIndex=true"
        )
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "fr-FR,fr;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        }

        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        prix_list = []

        # Stratégie 1 — JSON-LD (données structurées)
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                for item in (data if isinstance(data, list) else [data]):
                    offers = item.get("offers", {}) if isinstance(item, dict) else {}
                    if isinstance(offers, dict) and "price" in offers:
                        prix_list.append(float(offers["price"]))
                    elif isinstance(offers, list):
                        prix_list += [float(o["price"]) for o in offers if "price" in o]
            except Exception:
                pass

        # Stratégie 2 — Sélecteurs CSS connus de Chrono24
        if not prix_list:
            selectors = [
                ".js-article-item .price",
                ".article-price",
                "[data-testid='price']",
                ".price-value",
                ".wt-price",
                ".search-result-price",
            ]
            for sel in selectors:
                for el in soup.select(sel)[:10]:
                    text = el.get_text(strip=True)
                    nums = re.findall(r'[\d\s\u202f]+', text)
                    for n in nums:
                        cleaned = re.sub(r'\s|\u202f', '', n)
                        if cleaned and 3 <= len(cleaned) <= 8:
                            try:
                                v = float(cleaned)
                                if 200 <= v <= 10_000_000:
                                    prix_list.append(v)
                            except ValueError:
                                pass
                if prix_list:
                    break

        # Stratégie 3 — Regex sur le HTML brut (ex: "9 500 €")
        if not prix_list:
            pattern = r'(\d{1,3}(?:[\u202f\s]\d{3})*)\s*€'
            for m in re.finditer(pattern, resp.text):
                try:
                    v = float(re.sub(r'\s|\u202f', '', m.group(1)))
                    if 200 <= v <= 10_000_000:
                        prix_list.append(v)
                except ValueError:
                    pass
            prix_list = sorted(set(prix_list))[:30]

        if not prix_list:
            return None

        # Filtrage des valeurs aberrantes (IQR)
        prix_list = sorted(prix_list)
        if len(prix_list) >= 6:
            q1 = prix_list[len(prix_list) // 4]
            q3 = prix_list[3 * len(prix_list) // 4]
            iqr = q3 - q1
            if iqr > 0:
                prix_list = [p for p in prix_list
                             if q1 - 1.5 * iqr <= p <= q3 + 1.5 * iqr]

        if not prix_list:
            return None

        result = {
            "prix_min":    round(min(prix_list)),
            "prix_median": round(statistics.median(prix_list)),
            "prix_max":    round(max(prix_list)),
            "nb_annonces": len(prix_list),
            "source":      "Chrono24",
            "date":        now.strftime("%d/%m/%Y"),
            "from_cache":  False,
        }

        # ── Mise en cache ──
        if cached:
            cached.prix_min    = result["prix_min"]
            cached.prix_median = result["prix_median"]
            cached.prix_max    = result["prix_max"]
            cached.nb_annonces = result["nb_annonces"]
            cached.date_cache  = now
        else:
            db.session.add(MarketPriceCache(
                marque      = marque.strip(),
                reference   = reference.strip(),
                prix_min    = result["prix_min"],
                prix_median = result["prix_median"],
                prix_max    = result["prix_max"],
                nb_annonces = result["nb_annonces"],
            ))
        db.session.commit()
        return result

    except Exception:
        return None


# ── Routes ──
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        pwd   = request.form.get("password") or ""
        pwd2  = request.form.get("password2") or ""
        plan  = request.form.get("plan", "free")
        if not email or "@" not in email:
            flash("Adresse e-mail invalide.", "danger")
        elif len(pwd) < 8:
            flash("Mot de passe trop court (8 caractères minimum).", "danger")
        elif pwd != pwd2:
            flash("Les mots de passe ne correspondent pas.", "danger")
        elif User.query.filter_by(email=email).first():
            flash("Cette adresse est déjà utilisée.", "danger")
        else:
            user = User(email=email, plan=plan if plan in ("free","pro") else "free")
            user.set_password(pwd)
            db.session.add(user)
            db.session.commit()
            login_user(user)
            flash("Compte créé ! Bienvenue sur WatchFolio.", "success")
            return redirect(url_for("dashboard"))
        return render_template("register.html", email=email)
    return render_template("register.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        pwd   = request.form.get("password") or ""
        user  = User.query.filter_by(email=email).first()
        if user and user.check_password(pwd):
            login_user(user, remember=True)
            return redirect(request.args.get("next") or url_for("dashboard"))
        flash("E-mail ou mot de passe incorrect.", "danger")
        return render_template("login.html", email=email)
    return render_template("login.html")

@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("Déconnexion réussie.", "success")
    return redirect(url_for("index"))

@app.route("/dashboard")
@login_required
def dashboard():
    watches = current_user.watches.all()
    valeur_totale    = sum(w.prix_actuel for w in watches)
    cout_total       = sum(w.prix_achat for w in watches)
    plus_value_totale= valeur_totale - cout_total
    meilleure = max(watches, key=lambda w: w.plus_value_pct, default=None)
    pire      = min(watches, key=lambda w: w.plus_value_pct, default=None)
    # Distribution par marque
    marques = {}
    for w in watches:
        marques[w.marque] = marques.get(w.marque, 0) + 1
    marques = sorted(marques.items(), key=lambda x: x[1], reverse=True)[:8]
    return render_template("dashboard.html",
        watches=watches, valeur_totale=valeur_totale,
        cout_total=cout_total, plus_value_totale=plus_value_totale,
        meilleure=meilleure, pire=pire, marques=marques)

@app.route("/collection")
@login_required
def collection():
    watches = current_user.watches.all()
    return render_template("collection.html", watches=watches)

@app.route("/ajouter", methods=["GET", "POST"])
@login_required
def ajouter():
    if not current_user.can_add_watch:
        flash(f"Limite de {MAX_FREE_WATCHES} montres atteinte. Passez en Pro pour en ajouter plus.", "danger")
        return redirect(url_for("collection"))
    if request.method == "POST":
        try:
            prix = float(request.form.get("prix_achat", 0))
            annee_raw = request.form.get("annee", "").strip()
            annee = int(annee_raw) if annee_raw else None
            watch = Watch(
                user_id   = current_user.id,
                marque    = (request.form.get("marque") or "").strip(),
                reference = (request.form.get("reference") or "").strip(),
                annee     = annee,
                prix_achat= prix,
                etat      = request.form.get("etat", "Bon"),
                full_set  = request.form.get("full_set") == "on",
                notes     = (request.form.get("notes") or "").strip(),
            )
            db.session.add(watch)
            db.session.flush()
            # Prix initial dans l'historique
            ph = PriceHistory(watch_id=watch.id, prix=prix, source="Prix d'achat")
            db.session.add(ph)
            db.session.commit()
            flash(f"{watch.marque} {watch.reference} ajoutée à votre collection.", "success")
            return redirect(url_for("montre", watch_id=watch.id))
        except ValueError:
            flash("Prix invalide.", "danger")
    MARQUES = ["Rolex","Patek Philippe","Audemars Piguet","Richard Mille","Omega",
               "Breitling","IWC","Jaeger-LeCoultre","Cartier","Panerai",
               "Tudor","Seiko","Grand Seiko","TAG Heuer","Hublot",
               "Vacheron Constantin","A. Lange & Söhne","Zenith","Chopard"]
    return render_template("ajouter.html", marques=MARQUES)

@app.route("/montre/<int:watch_id>")
@login_required
def montre(watch_id):
    w = Watch.query.get_or_404(watch_id)
    if w.user_id != current_user.id:
        abort(403)
    conseil     = conseil_ia(w)
    history     = w.prix_history.all()
    prix_marche = get_prix_marche(w.marque, w.reference)
    return render_template("montre.html", watch=w, conseil=conseil,
                           history=history, prix_marche=prix_marche)

@app.route("/montre/<int:watch_id>/prix", methods=["POST"])
@login_required
def ajouter_prix(watch_id):
    w = Watch.query.get_or_404(watch_id)
    if w.user_id != current_user.id:
        abort(403)
    try:
        prix   = float(request.form.get("prix", 0))
        source = (request.form.get("source") or "Manuel").strip()
        ph = PriceHistory(watch_id=w.id, prix=prix, source=source)
        db.session.add(ph)
        db.session.commit()
        flash("Prix mis à jour.", "success")
    except ValueError:
        flash("Prix invalide.", "danger")
    return redirect(url_for("montre", watch_id=watch_id))

@app.route("/montre/<int:watch_id>/supprimer", methods=["POST"])
@login_required
def supprimer_montre(watch_id):
    w = Watch.query.get_or_404(watch_id)
    if w.user_id != current_user.id:
        abort(403)
    label = f"{w.marque} {w.reference}"
    db.session.delete(w)
    db.session.commit()
    flash(f"{label} supprimée de votre collection.", "success")
    return redirect(url_for("collection"))

@app.route("/marche")
def marche():
    # Page publique : tendances générales
    tendances = [
        {"marque": "Rolex", "modele": "Submariner Date", "ref": "126610LN", "tendance": "hausse", "variation": "+12%"},
        {"marque": "Patek Philippe", "modele": "Nautilus", "ref": "5711/1A", "tendance": "stable", "variation": "+2%"},
        {"marque": "Audemars Piguet", "modele": "Royal Oak", "ref": "15500ST", "tendance": "hausse", "variation": "+8%"},
        {"marque": "Rolex", "modele": "Daytona", "ref": "116500LN", "tendance": "hausse", "variation": "+15%"},
        {"marque": "Omega", "modele": "Speedmaster Pro", "ref": "310.30.42.50", "tendance": "stable", "variation": "0%"},
        {"marque": "Tudor", "modele": "Black Bay 58", "ref": "M79030N", "tendance": "baisse", "variation": "-3%"},
    ]
    return render_template("marche.html", tendances=tendances)

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5001)
