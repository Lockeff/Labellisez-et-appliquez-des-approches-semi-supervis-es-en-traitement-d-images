import base64
import contextlib
import hashlib
import io
import json
from pathlib import Path

import pandas as pd
from sklearn.cluster import DBSCAN, KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import adjusted_rand_score
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

root = Path(__file__).resolve().parents[2]
build = Path(__file__).resolve().parent
notebook_path = root / "01_exploration_dataset.ipynb"
notebook_bytes = notebook_path.read_bytes()
notebook = json.loads(notebook_bytes.decode("utf-8"))
cells = notebook["cells"]
byid = {cell["id"]: cell for cell in cells}

env = {
    "pd": pd, "StandardScaler": StandardScaler, "PCA": PCA,
    "KMeans": KMeans, "DBSCAN": DBSCAN,
    "adjusted_rand_score": adjusted_rand_score,
    "train_test_split": train_test_split,
    "chemin_features": root / "resultats/etape2/features_resnet18.csv",
    "colonnes_features": [f"feature_{i:03d}" for i in range(512)],
    "display": lambda value: None,
}

# Execute the current clustering cells in memory. Never save the notebook.
with contextlib.redirect_stdout(io.StringIO()):
    for marker in [
        "features = pd.read_csv", "X_features =", "pca = PCA",
        "kmeans = KMeans", "essais_dbscan =", "labels_reference =",
        "faibles_train = sans_label_train",
    ]:
        source = next(
            "".join(cell["source"]) for cell in cells
            if cell["cell_type"] == "code" and marker in "".join(cell["source"])
        )
        exec(source, env)

def output_text(cid):
    return "\n".join(
        "".join(output.get("text", output.get("data", {}).get("text/plain", [])))
        for output in byid[cid].get("outputs", [])
    )

# Numeric CNN scores and confusion counts come from the saved output, not a new training.
comparison_text = output_text("ee07cebf")
assert "0.889" in comparison_text and "0.947" in comparison_text
assert "0.90" in comparison_text and "0.95" in comparison_text
semi_search_text = output_text("hyperparametres-semi-recherche")
super_search_text = output_text("hyperparametres-supervise-recherche")
assert "0.9756" in semi_search_text and "0.7692" in super_search_text

data = {
    "notebook_sha256": hashlib.sha256(notebook_bytes).hexdigest(),
    "source_notebook": str(notebook_path),
    "clustering_source": "Current notebook cells executed in memory on the saved features CSV",
    "cnn_source": "Saved notebook outputs; no new CNN training",
    "kmeans": {
        "groups": [
            env["X_pca"][env["groupes_kmeans_sans_label"] == group].tolist()
            for group in [0, 1]
        ],
        "counts": pd.Series(env["groupes_kmeans_sans_label"]).value_counts().sort_index().tolist(),
        "ari": float(env["ari_kmeans"]),
    },
    "dbscan": [
        {key: value for key, value in trial.items() if key != "groupes"}
        for trial in env["essais_dbscan"]
    ],
    "weak_counts": env["donnees_faibles"]["label_faible"].value_counts().sort_index().tolist(),
    "sizes": {key: len(env[key]) for key in [
        "donnees_fortes", "donnees_sans_label", "fortes_train", "fortes_validation",
        "fortes_test", "sans_label_train", "sans_label_test",
    ]},
    "cnn_metrics": {
        "semi": [0.90, 0.889, 1.0, 0.8, 0.833],
        "super": [0.95, 0.947, 1.0, 0.9, 0.918],
    },
    "cnn_confusions": {"semi": [[10, 0], [2, 8]], "super": [[10, 0], [1, 9]]},
    "f2_validation": {"semi": [0.5682, 0.5682, 0.9091, 0.9756], "super": [0.7317, 0.75, 0.5405, 0.7692]},
    "weak_test": {"accuracy": 0.973, "f1_group1": 0.968, "n": 263},
    "images": {
        label: str(sorted((root / "mri_dataset_brain_cancer_oc" / folder).glob("*.jpg"))[0])
        for label, folder in [("normal", "avec_labels/normal"), ("cancer", "avec_labels/cancer"), ("unlabelled", "sans_label")]
    },
    "markdowns": ["".join(cell["source"]) for cell in cells if cell["cell_type"] == "markdown"],
}

# Preserve the recorded confusion image as evidence for the source numbers.
for output in byid["ee07cebf"].get("outputs", []):
    png = output.get("data", {}).get("image/png")
    if png:
        if isinstance(png, list):
            png = "".join(png)
        (build / "source_confusions.png").write_bytes(base64.b64decode(png))

(build / "sources.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
assert notebook_path.read_bytes() == notebook_bytes
print("Prepared current clustering data and saved CNN results. Notebook unchanged.")
print("K-Means counts:", data["kmeans"]["counts"], "ARI:", data["kmeans"]["ari"])
