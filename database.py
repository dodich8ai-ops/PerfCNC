"""
PerfCNC — modèles SQLAlchemy.
3 tables : User, Analysis, Consent.
Initialiser avec db.init_app(app) puis db.create_all().
"""
import datetime
import json

from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

MAX_FREE_ANALYSES = 5
CGU_VERSION       = "1.0"


class User(db.Model, UserMixin):
    __tablename__ = "user"

    id            = db.Column(db.Integer, primary_key=True)
    email         = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    plan          = db.Column(db.String(10),  nullable=False, default="free")
    created_at    = db.Column(db.DateTime,    nullable=False,
                              default=datetime.datetime.utcnow)

    analyses = db.relationship(
        "Analysis", backref="user", lazy="dynamic",
        cascade="all, delete-orphan",
        order_by="Analysis.uploaded_at.desc()",
    )
    consent = db.relationship(
        "Consent", backref="user", uselist=False,
        cascade="all, delete-orphan",
    )

    # ── Mot de passe ──

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    # ── Quotas ──

    @property
    def analyses_count(self) -> int:
        return self.analyses.count()

    @property
    def can_save_analysis(self) -> bool:
        return self.plan == "pro" or self.analyses_count < MAX_FREE_ANALYSES

    @property
    def analyses_remaining(self):
        """None si Pro (illimité), int sinon."""
        if self.plan == "pro":
            return None
        return max(0, MAX_FREE_ANALYSES - self.analyses_count)

    @property
    def display_plan(self) -> str:
        return "Pro" if self.plan == "pro" else "Gratuit"

    @property
    def has_accepted_cgu(self) -> bool:
        return self.consent is not None


class Analysis(db.Model):
    __tablename__ = "analysis"

    id          = db.Column(db.Integer,    primary_key=True)
    user_id     = db.Column(db.Integer,    db.ForeignKey("user.id"), nullable=False, index=True)
    filename    = db.Column(db.String(255), nullable=False)
    uploaded_at = db.Column(db.DateTime,   nullable=False,
                            default=datetime.datetime.utcnow)
    # Sérialisation complète du résultat process_excel() (sans preview)
    stats_json  = db.Column(db.Text, nullable=False)

    @property
    def stats(self) -> dict:
        return json.loads(self.stats_json)

    # Raccourcis pour les colonnes du tableau historique
    @property
    def trs(self):
        try:
            return self.stats.get("indicateurs", {}).get("trs")
        except Exception:
            return None

    @property
    def total_saisies(self) -> int:
        try:
            return self.stats.get("total_saisies", 0)
        except Exception:
            return 0

    @property
    def total_rebuts(self) -> int:
        try:
            return self.stats.get("total_rebuts", 0)
        except Exception:
            return 0

    @property
    def uploaded_at_fr(self) -> str:
        return self.uploaded_at.strftime("%d/%m/%Y à %H:%M")


class Consent(db.Model):
    __tablename__ = "consent"

    id          = db.Column(db.Integer,   primary_key=True)
    user_id     = db.Column(db.Integer,   db.ForeignKey("user.id"), nullable=False)
    accepted_at = db.Column(db.DateTime,  nullable=False,
                            default=datetime.datetime.utcnow)
    version_cgu = db.Column(db.String(10), nullable=False, default=CGU_VERSION)
