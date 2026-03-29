# WatchFolio

Gestionnaire de collection de montres de luxe avec conseil de marché par intelligence artificielle (Mistral via Ollama).

## Fonctionnalités

- Suivi de la valeur de chaque montre (historique des prix)
- Tableau de bord : valeur totale, plus-values, distribution par marque
- Conseil IA par montre (tendance, fourchette de prix, recommandation de vente)
- Plan gratuit (10 montres) et plan Pro (illimité)
- Interface sombre luxe, PWA-ready

## Installation

```bash
pip install -r requirements.txt
python app.py
```

L'application démarre sur `http://localhost:5001`.

## Variables d'environnement

| Variable       | Description                              | Défaut                            |
|----------------|------------------------------------------|-----------------------------------|
| `SECRET_KEY`   | Clé secrète Flask (à changer en prod)    | `watchfolio_dev_secret_change_in_prod` |
| `DATABASE_URL` | URL SQLAlchemy de la base de données     | `sqlite:///instance/watchfolio.db` |

## IA — Ollama / Mistral

Pour activer les conseils IA, installez Ollama puis téléchargez le modèle :

```bash
ollama pull mistral
```

Redémarrez ensuite l'application. Si Ollama est hors ligne, les pages restent fonctionnelles (le conseil affiche un message d'indisponibilité).

## Production (Gunicorn)

```bash
SECRET_KEY=votre_cle_secrete gunicorn -w 4 -b 0.0.0.0:8000 app:app
```
