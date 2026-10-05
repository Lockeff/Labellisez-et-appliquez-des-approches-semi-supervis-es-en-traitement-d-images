# Traitement semi-supervisé d’images IRM

Exploration du dataset et développement des approches semi-supervisées dans des notebooks Jupyter.

## Environnement de développement

Le projet utilise **Python 3.12** et **uv**. `pyproject.toml` déclare les dépendances,
`uv.lock` verrouille leurs versions et `.python-version` indique la version de Python.
L’environnement virtuel local est créé dans `.venv/` ; il n’est pas versionné.

Si uv n’est pas encore installé (macOS / Linux) :

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Ouvrir ensuite un nouveau terminal pour disposer de la commande `uv`.

Depuis la racine du projet, créer ou restaurer l’environnement :

```bash
uv sync --locked
```

Lancer JupyterLab :

```bash
uv run jupyter lab
```

Ouvrir `01_exploration_dataset.ipynb` et utiliser le noyau Python de cet environnement.
Dans VS Code, choisir **Sélectionner le noyau → Environnements Python → .venv/bin/python**.

L’activation manuelle de `.venv` n’est pas nécessaire avec `uv run`. Pour l’activer
dans un terminal macOS / Linux :

```bash
source .venv/bin/activate
```

Ajouter une dépendance au projet :

```bash
uv add nom-du-package
```

Ajouter un outil de développement :

```bash
uv add --dev nom-du-package
```

Les dépendances scientifiques sont NumPy, pandas, Matplotlib et Pillow.
JupyterLab et ipykernel sont déclarés comme outils de développement.
