[README.md](https://github.com/user-attachments/files/27444495/README.md)
# ADV — Extraction de bons de commande (MVP)

Outil d'extraction automatique de bons de commande par IA, avec export Excel prêt pour SAP.

## Installation

```bash
# 1. Créer un environnement virtuel (recommandé)
python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows

# 2. Installer les dépendances
pip install -r requirements.txt

# 3. Lancer l'application
streamlit run app.py
```

L'application s'ouvre automatiquement sur http://localhost:8501

## Utilisation

1. Entrez votre clé API Anthropic dans la barre latérale (console.anthropic.com)
2. Uploadez un bon de commande (PDF ou image)
3. Cliquez sur "Analyser le BC"
4. Vérifiez les données extraites
5. Téléchargez le fichier Excel pour SAP

## Prochaines étapes (roadmap MVP)

- [ ] Connexion Salesforce (simple-salesforce) pour enrichissement client
- [ ] Sélection de contrat par ligne de commande
- [ ] Macro VBA intégrée dans l'Excel pour pilotage SAP GUI
- [ ] Historique des commandes traitées (SQLite)
- [ ] Gestion multi-utilisateurs

## Coût API estimé

- Claude claude-opus-4-5 : ~0,015 € par BC analysé
- Budget démo (100 BCs) : ~1,50 €
