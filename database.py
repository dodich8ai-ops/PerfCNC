from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
import datetime

db = SQLAlchemy()
MAX_FREE_WATCHES = 10
CGU_VERSION = "1.0"

class User(db.Model):
    __tablename__ = "user"
    id           = db.Column(db.Integer, primary_key=True)
    email        = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash= db.Column(db.String(256), nullable=False)
    plan         = db.Column(db.String(20), nullable=False, default="free")
    created_at   = db.Column(db.DateTime, default=datetime.datetime.utcnow)
    watches      = db.relationship("Watch", backref="owner", lazy="dynamic",
                                   cascade="all, delete-orphan",
                                   order_by="Watch.date_ajout.desc()")

    def set_password(self, pw): self.password_hash = generate_password_hash(pw)
    def check_password(self, pw): return check_password_hash(self.password_hash, pw)
    def get_id(self): return str(self.id)
    @property
    def is_authenticated(self): return True
    @property
    def is_active(self): return True
    @property
    def is_anonymous(self): return False

    @property
    def watches_count(self): return self.watches.count()
    @property
    def can_add_watch(self):
        return self.plan == "pro" or self.watches_count < MAX_FREE_WATCHES
    @property
    def display_plan(self): return "Pro" if self.plan == "pro" else "Gratuit"


class Watch(db.Model):
    __tablename__ = "watch"
    id         = db.Column(db.Integer, primary_key=True)
    user_id    = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False, index=True)
    marque     = db.Column(db.String(100), nullable=False)
    reference  = db.Column(db.String(100), nullable=False)
    annee      = db.Column(db.Integer)
    prix_achat = db.Column(db.Float, nullable=False)
    etat       = db.Column(db.String(20), nullable=False, default="Bon")
    full_set   = db.Column(db.Boolean, default=False)
    notes      = db.Column(db.Text, default="")
    date_ajout = db.Column(db.DateTime, default=datetime.datetime.utcnow)
    prix_history = db.relationship("PriceHistory", backref="watch", lazy="dynamic",
                                   cascade="all, delete-orphan",
                                   order_by="PriceHistory.date.desc()")

    @property
    def prix_actuel(self):
        last = self.prix_history.first()
        return last.prix if last else self.prix_achat

    @property
    def plus_value(self):
        return self.prix_actuel - self.prix_achat

    @property
    def plus_value_pct(self):
        if self.prix_achat == 0: return 0
        return (self.plus_value / self.prix_achat) * 100

    @property
    def date_ajout_fr(self):
        return self.date_ajout.strftime("%d/%m/%Y") if self.date_ajout else "—"


class PriceHistory(db.Model):
    __tablename__ = "price_history"
    id       = db.Column(db.Integer, primary_key=True)
    watch_id = db.Column(db.Integer, db.ForeignKey("watch.id"), nullable=False, index=True)
    prix     = db.Column(db.Float, nullable=False)
    source   = db.Column(db.String(100), default="Manuel")
    date     = db.Column(db.DateTime, default=datetime.datetime.utcnow)

    @property
    def date_fr(self):
        return self.date.strftime("%d/%m/%Y") if self.date else "—"
