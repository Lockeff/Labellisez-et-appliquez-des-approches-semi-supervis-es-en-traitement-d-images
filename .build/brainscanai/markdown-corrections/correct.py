import copy
import json
from datetime import datetime
from pathlib import Path

path = Path('01_exploration_dataset.ipynb')
original_bytes = path.read_bytes()
original = json.loads(original_bytes)
notebook = copy.deepcopy(original)

corrections = {
    '1400e547': """## Étape 1 : import des données et exploration des images

Le dataset est organisé en deux répertoires : un avec labels et un sans label.

normal : image labellisée comme ne présentant pas de tumeur.
cancer : image labellisée comme présentant une tumeur cancéreuse.

Le fichier .txt décrit le contenu du dataset : **1 500 images**, au format **JPEG**, dont 100 étiquetées et 1 400 sans label.

Le code compte les images et affiche l'écart avec ce descriptif.""",
    'f166cf32': """Affichage de quelques exemples de chaque dossier.

`random.seed(42)` permet de retrouver la même sélection lors d'une réexécution.""",
    'aaccdcc1': """Sélection d'un échantillon de 100 images.

Affichage du nombre d'images sélectionnées.""",
    '78089379': """Pour chaque fichier de l'échantillon, appel de `image.load()` pour lire ses pixels. Ses dimensions, son mode de couleur, son nombre de canaux et son format sont ajoutés à une liste. Les erreurs de lecture sont conservées séparément.

Affichage des premières lignes du tableau `controle`.""",
    '0eeec8e2': """Regroupons les images pour vérifier qu'elles ont les mêmes caractéristiques.

Affichage des effectifs par dimensions, mode de couleur, nombre de canaux et format, puis du nombre d'images lisibles et d'erreurs de lecture.""",
    '102ebaa3': """On effectue la recherche de doublons par dossier cette fois-ci.""",
    '115501c3': """Conclusion

- Le comptage donne **1 506 images JPEG** : 50 dans cancer, 50 dans normal et 1 406 dans sans_label.
- Les deux classes étiquetées sont équilibrées.
- Le dossier sans label contient **six images de plus** que le descriptif.
- Les **100 images contrôlées** mesurent **512 × 512 pixels** et sont stockées en **RGB, avec trois canaux**.
- Aucune erreur de lecture n'a été rencontrée dans cet échantillon.
- Les images montrent différentes orientations de coupe et différents cadrages.
- La place occupée par le cerveau, le contraste et la netteté apparente varient.
- Les exemples apparaissent principalement en niveaux de gris malgré le stockage RGB.
- Une même taille de 512 × 512 pixels ne garantit pas une qualité visuelle uniforme.
- Le dataset contient 96 copies supplémentaires.

Ces doublons seront à traiter.""",
    '6f08787d': """Traitement des doublons.

On écarte les copies supplémentaires de la liste `chemins_images`, en conservant la première occurrence de chaque image. Les fichiers sources restent inchangés.

Affichage du nombre d'images avant filtrage, des copies écartées et des images conservées pour l'étape 2.""",
    'ba54e8fb': """## Étape 2 : prétraitement et extraction des features""",
    'mentor-ressources-explication': """### Mesurer la durée et la RAM

Définition de `mesurer_ressources`, qui mesurera la durée et suivra la RAM du processus Python toutes les 0,05 seconde avec `psutil`.
Chaque mesure conservera la RAM au départ, le maximum observé et l'augmentation maximale, en Mio. Les pauses entre cellules ne sont pas comptées dans ces mesures.

La cellule affiche que la mesure est prête.""",
    'fa1213ea': """Choix des poids pré-entraînés de ResNet18 : `IMAGENET1K_V1`.

Configuration du prétraitement qui réalisera le redimensionnement, le recadrage, la conversion et la normalisation des images.

Affichage des paramètres de ce prétraitement.""",
    'mentor-egalisation-explication': """Égalisation des histogrammes""",
    'd53f9bbb': """Application de `pretraitement` à chaque image conservée dans `chemins_images`.
Stockage du résultat dans la liste `images_preparees`.

Affichage du nombre d'images préparées, de leur forme et du type de leurs valeurs.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    '4ac06b74': """Chargement de ResNet18 avec les poids pré-entraînés précédemment choisis.
Gel des paramètres avec `requires_grad = False`, y compris ceux des couches convolutionnelles.

On remplace la couche de classification par `torch.nn.Identity()` pour que le modèle renvoie les 512 nombres décrivant chaque image. `modele.eval()` le place en mode évaluation.

Puis on vérifie que les paramètres sont bien gelés en affichant le nombre de paramètres entraînables.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    '5219907b': """Extraction des embeddings des images préparées :

- On traite les images préparées par lots de 16 : pour chaque image, le modèle crée un vecteur de 512 nombres décrivant l'image.
- On ajoute avec `append()` chaque lot de vecteurs dans `blocs_features`.
- Après la boucle, on concatène avec `np.concatenate()` les blocs en un tableau unique.

Affichage de la progression et des dimensions de `matrice_features`.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    'a7319dfa': """Transformation de `matrice_features` en tableau pandas nommé `features`, avec les colonnes numériques nommées `feature_000` à `feature_511`.

La première colonne du tableau contient le chemin relatif de l'image.

Affichage des dimensions et des premières lignes du tableau.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    'bc5842e2': """## Étape 3 : modélisation non supervisée""",
    'de6350e6': """Réalisation du clustering avec K-Means.

- `n_clusters=2` pour la séparation en deux groupes, comme pour le dataset labellisé.
- `random_state=42` pour pouvoir reproduire le clustering.

K-Means est entraîné uniquement sur le train sans label. Affichage des effectifs de chaque groupe.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    'de5bb28e': """Graphique illustrant les coordonnées des images du train sans label.

Chaque point représente une image, placée selon ses deux coordonnées PCA. Sa couleur indique le groupe attribué par K-Means.""",
    '3dbfb75e': """K-Means répartit les images en deux groupes, principalement séparés par l'abscisse PCA (gauche et droite). Les groupes présentent une dispersion interne. Cette représentation montre des différences entre les images, sans pour autant identifier les groupes comme « normal » et « cancer ».""",
    'a4f6f460': """Réalisation du clustering avec DBSCAN.

Choix de neuf combinaisons à tester pour visualiser ensuite les résultats. On teste toutes les combinaisons entre :

- `eps` (rayon du voisinage) : 0.5, 1.0, 2.0 ;
- `min_samples` (nombre minimum de points dans ce voisinage, en comptant le point lui-même) : 3, 5, 10.

Chaque résultat est conservé dans `essais_dbscan` avec ses paramètres, ses numéros de groupes, le nombre de groupes et le nombre d'images hors groupe (`-1`).

DBSCAN est entraîné uniquement sur le train sans label. Affichage des paramètres et des effectifs pour chaque essai.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    'fe141681': """Affichage d'un graphique pour chaque essai de paramètres avec DBSCAN : une couleur représente chaque groupe ; les points hors groupe (`-1`) sont affichés en gris.""",
    '6324dcb3': """Labels faibles du train sans label issus de K-Means ; labels du test sans label prédits sans réentraînement. Les labels forts restent séparés.

Affichage des effectifs et des premières lignes du tableau des labels faibles.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    '78faa038': """## Étape 4 : entraînement du CNN

Métriques utilisées :

- L'accuracy mesure la proportion globale de bonnes prédictions.
- La précision indique la fiabilité des prédictions « cancer ».
- Le rappel mesure la proportion d'images cancéreuses reconnues.
- Le F1 résume l'équilibre entre précision et rappel.
- Dans cette étude exploratoire, je privilégie le F2 pour sélectionner les hyperparamètres, car il donne davantage d'importance au rappel et donc à la réduction des faux négatifs.
- La matrice de confusion permet de voir directement les erreurs.""",
    '280a88b9': """Mise en forme des données pour le CNN :

- On associe chaque image à son étiquette. Les labels faibles conservent les numéros des groupes K-Means ; les labels forts sont encodés avec normal = 0 et cancer = 1.
- On prépare les chargeurs qui fourniront les couples image/étiquette au CNN.

Affichage du nombre de lots de chaque ensemble et de l'appareil utilisé.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    '668b5564': """Création de deux CNN :

- un pour l'entraînement semi-supervisé ;
- un pour l'entraînement supervisé.

On utilise les mêmes poids pré-entraînés qu'à l'étape 2. Les paramètres pré-entraînés sont gelés et seule la nouvelle couche finale à deux sorties est entraînable.

Affichage du nombre de paramètres entraînables de chaque CNN.
La durée et le pic de RAM de cette cellule sont mesurés.""",
    'eacb5d90': """Création de deux fonctions :

- `entrainer_cnn` entraîne le CNN en ajustant les poids de la couche finale et affiche la perte moyenne à chaque époque.
- `evaluer_cnn` réalise les prédictions sans modifier les poids, puis renvoie les métriques et la matrice de confusion. Les métriques sont l'accuracy, le F1, la précision, le rappel et le F2 pour le label 1 : groupe 1 sur les labels faibles, cancer sur les labels forts.

La cellule affiche que les fonctions sont définies.""",
    'hyperparametres-semi-explication': """Recherche d'hyperparamètres pour le modèle semi-supervisé :

- D'abord, comparaison de deux configurations sur les labels faibles : taux de 0,001 et 5 ou 10 époques. On utilise les 16 images de validation réservées à l'étape 3, avec leurs groupes K-Means comme cibles. On retient la meilleure accuracy d'accord avec ces groupes ; à score égal, on conserve 5 époques.
- Ensuite, chaque essai reprend le meilleur modèle faible pour comparer quatre configurations sur les labels forts : taux de 0,0001 ou 0,001 et 5 ou 10 époques. On utilise les vrais labels des mêmes 16 images de validation pour retenir le meilleur F2 cancer.

Les tableaux affichent l'accuracy de validation faible, puis les métriques de validation forte. Les paramètres retenus sont affichés et le CNN semi-supervisé est recréé pour les entraînements suivants.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    '4709dc64': """Entraînement sur les labels faibles avec le CNN semi-supervisé recréé, le taux de son optimiseur et le nombre d'époques retenu.

Puis évaluation sur les 263 images du test faible : le tableau mesure l'accord avec leurs groupes K-Means.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    'e1ec12cb': """Puis entraînement du même CNN sur les données fortes de `fortes_train`, avec le taux et le nombre d'époques fortes retenus.

Évaluation sur les 20 images du test fort et affichage des métriques pour la classe cancer.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    'f5d694fb': """Entraînement de `modele_supervise` sur les données fortes uniquement, avec les hyperparamètres retenus.

Puis évaluation sur les mêmes 20 images du test fort et avec les mêmes métriques que pour l'entraînement semi-supervisé.

La durée et le pic de RAM de cette cellule sont mesurés.""",
    '3e8e10ca': """Comparaison de la méthode supervisée et de la méthode semi-supervisée.

Comparaison des deux modèles sur le même test fort, avec leurs hyperparamètres sélectionnés sur la validation forte.

Affichage du tableau des métriques, des matrices de confusion et de l'écart de F1 cancer.""",
    'a9dcddf9': """Durée écoulée depuis le début du notebook jusqu'à cette cellule.

Affichage de la durée en secondes et en minutes.""",
    '1d403f38': """Le fait d'avoir beaucoup d'images non labellisées et peu d'images labellisées justifie de tester un entraînement semi-supervisé.
Cependant, les résultats montrent ici que le CNN supervisé est plus performant. Dans cette configuration et sur ce test, le semi-supervisé n'apporte pas d'amélioration.""",
    'cbac9625': """Mesure du temps nécessaire pour lire, prétraiter et attribuer un label à 1 000 images avec le CNN supervisé déjà entraîné. Les images sont traitées par lots de 16. La cellule affiche la durée mesurée et une estimation pour 4 millions d'images, en supposant une vitesse comparable sur la même machine. L'entraînement n'est pas inclus.

Le même prétraitement avec égalisation est utilisé. La durée et le pic de RAM du processus Python sont mesurés séparément de la création du modèle.""",
    'mentor-bilan-ressources-explication': """### Durée et RAM : création du modèle et traitement final

Afficher le détail de `mesures_ressources`, puis un bilan en deux lignes.

- La création regroupe les opérations mesurées : prétraitement, deux extractions, essais de clustering, recherches d'hyperparamètres, entraînements et leurs évaluations.
- Le traitement final porte sur l'échantillon de 1 000 images de la cellule précédente.

Dans chaque groupe, les durées mesurées sont additionnées et la RAM retenue est le maximum observé pour le processus Python.""",
}

by_id = {c['id']: c for c in notebook['cells']}
observation_id = 'cnn-surapprentissage-explication'
observation = ''.join(by_id[observation_id]['source'])
corrections[observation_id] = observation.rstrip() + "\n\nL'écart observé entre train et test suggère un possible surapprentissage pour le modèle semi-supervisé."

changed = []
for cid, text in corrections.items():
    cell = by_id[cid]
    assert cell['cell_type'] == 'markdown', cid
    if ''.join(cell['source']) != text:
        cell['source'] = text.splitlines(keepends=True)
        changed.append(cid)

# Every difference must be the source of an explicitly selected Markdown cell.
assert len(original['cells']) == len(notebook['cells'])
assert {k:v for k,v in original.items() if k != 'cells'} == {k:v for k,v in notebook.items() if k != 'cells'}
for before, after in zip(original['cells'], notebook['cells']):
    if before != after:
        assert before['cell_type'] == after['cell_type'] == 'markdown'
        assert before['id'] in changed
        assert {k:v for k,v in before.items() if k != 'source'} == {k:v for k,v in after.items() if k != 'source'}
    else:
        assert before == after

build = Path('.build/brainscanai/markdown-corrections')
backup = build / ('before-' + datetime.now().strftime('%Y%m%d-%H%M%S') + '.ipynb')
backup.write_bytes(original_bytes)
assert path.read_bytes() == original_bytes, 'Notebook edited concurrently; stopping before write.'
path.write_text(json.dumps(notebook, ensure_ascii=False, indent=1) + '\n', encoding='utf-8', newline='\n')
saved = json.loads(path.read_bytes())
assert saved == notebook
assert [c for c in saved['cells'] if c['cell_type'] != 'markdown'] == [c for c in original['cells'] if c['cell_type'] != 'markdown']
print(f'Corrected {len(changed)} Markdown cells. Code, outputs, execution counts, cell order and metadata preserved.')
print('Edited cell numbers:', [i+1 for i,c in enumerate(saved['cells']) if c['id'] in changed])
