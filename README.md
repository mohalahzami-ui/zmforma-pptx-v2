markdown# ZMForma PowerPoint Generator API v2

API Flask pour générer des présentations PowerPoint professionnelles.

## Déploiement sur Render

1. Pousse ce code sur GitHub
2. Crée un nouveau Web Service sur Render
3. Configure :
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
   - Instance Type: Free

## Endpoints

- `GET /health` - Vérifie que l'API fonctionne
- `GET /test` - Génère un PowerPoint de test
- `POST /generate` - Génère un PowerPoint depuis JSON

## Structure du projet
```
zmforma-pptx-v2/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
└── utils/
    ├── __init__.py
    ├── styles.py
    └── slide_builder.py
```

## Développement local
```bash
pip install -r requirements.txt
python app.py
```

L'API sera accessible sur `http://localhost:5000`
