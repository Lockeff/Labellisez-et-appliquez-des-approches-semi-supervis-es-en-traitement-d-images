"""Add only the requested preprocessing, layer trial and resource measurements."""
from pathlib import Path
import copy
import json
import textwrap
import ast
import inspect

ROOT = Path(__file__).resolve().parents[2]
PATH = ROOT / '01_exploration_dataset.ipynb'
BUILD = ROOT / '.build/brainscanai/mentor-notebook'
BUILD.mkdir(exist_ok=True)
raw = PATH.read_bytes()
(BUILD / 'original.ipynb').write_bytes(raw)
notebook = json.loads(raw)
original = copy.deepcopy(notebook)
cells = {cell['id']: cell for cell in notebook['cells']}

def source(code):
    return textwrap.dedent(code).strip('\n').splitlines(keepends=True)

def pair(markdown_id, code_id, markdown, code):
    return [
        {'cell_type': 'markdown', 'id': markdown_id, 'metadata': {}, 'source': source(inspect.cleandoc(markdown))},
        {'cell_type': 'code', 'id': code_id, 'metadata': {}, 'source': source(inspect.cleandoc(code)),
         'execution_count': None, 'outputs': []},
    ]

def insert_before(cell_id, new_cells):
    position = next(i for i, c in enumerate(notebook['cells']) if c['id'] == cell_id)
    notebook['cells'][position:position] = new_cells

def insert_after(cell_id, new_cells):
    position = next(i for i, c in enumerate(notebook['cells']) if c['id'] == cell_id) + 1
    notebook['cells'][position:position] = new_cells

insert_before('fa1213ea', pair(
    'mentor-ressources-explication', 'mentor-ressources-definition',
    '''### Mesurer la durée et la RAM

    Définir `mesurer_ressources` pour chronométrer les cellules de calcul et suivre la RAM du processus Python toutes les 0,05 seconde avec `psutil`.
    Chaque mesure conserve la RAM au départ, le pic observé et la hausse maximale, en Mio. Les pauses entre cellules ne sont pas comptées ; la mémoire du GPU n'est pas mesurée.
    Résultat : la fonction de mesure et le dictionnaire de résultats sont prêts.''',
    '''from contextlib import contextmanager
    from threading import Event, Thread
    import psutil

    mesures_ressources = {}
    processus_python = psutil.Process()

    @contextmanager
    def mesurer_ressources(nom, groupe="Création du modèle"):
        if torch.cuda.is_available():
            torch.cuda.synchronize()
        ram_depart = processus_python.memory_info().rss
        pic_ram = [ram_depart]
        arret = Event()

        def suivre_ram():
            while not arret.wait(0.05):
                pic_ram[0] = max(pic_ram[0], processus_python.memory_info().rss)

        suivi = Thread(target=suivre_ram, daemon=True)
        suivi.start()
        debut = time.perf_counter()
        try:
            yield
        finally:
            if torch.cuda.is_available():
                torch.cuda.synchronize()
            duree = time.perf_counter() - debut
            arret.set()
            suivi.join()
            pic_ram[0] = max(pic_ram[0], processus_python.memory_info().rss)
            mesures_ressources[nom] = {
                "Groupe": groupe,
                "Opération": nom,
                "Durée (s)": duree,
                "RAM au départ (Mio)": ram_depart / 1024**2,
                "Pic RAM (Mio)": pic_ram[0] / 1024**2,
                "Hausse maximale (Mio)": (pic_ram[0] - ram_depart) / 1024**2,
            }

    print("Mesure de la durée et du pic de RAM prête.")''',
))

insert_after('1ef4c71e', pair(
    'mentor-egalisation-explication', 'mentor-egalisation',
    'Egalisation des histogrammes',
    '''from torchvision.transforms.functional import equalize

    # Egaliser les pixels uint8 avant la normalisation ImageNet.
    pretraitement_imagenet = poids.transforms()

    def pretraitement(image):
        image_egalisee = equalize(image)
        return pretraitement_imagenet(image_egalisee)

    print("Egalisation des histogrammes avant le prétraitement ImageNet.")
    print(pretraitement_imagenet)''',
))

insert_after('0f6b5cdd', pair(
    'mentor-intermediaire-explication', 'mentor-intermediaire-extraction',
    '''### Essai d'une couche intermédiaire

    Réutiliser les couches gelées de `modele` jusqu'à `layer3`, en arrêtant avant `layer4`.
    Une moyenne spatiale transforme ses 256 cartes en un vecteur de 256 nombres par image. Traiter `images_preparees` par lots de 16 et afficher les dimensions des deux matrices.
    Cet essai reste en mémoire et ne remplace pas `matrice_features`, qui contient les embeddings habituels de 512 dimensions. La durée et le pic de RAM de cet essai sont mesurés.''',
    '''with mesurer_ressources("Essai des embeddings layer3"):
        modele_intermediaire = torch.nn.Sequential(
            modele.conv1, modele.bn1, modele.relu, modele.maxpool,
            modele.layer1, modele.layer2, modele.layer3,
            torch.nn.AdaptiveAvgPool2d((1, 1)),
            torch.nn.Flatten(),
        )
        modele_intermediaire.eval()
        blocs_intermediaires = []

        with torch.no_grad():
            for debut in range(0, len(images_preparees), taille_lot):
                lot = torch.stack(images_preparees[debut:debut + taille_lot])
                blocs_intermediaires.append(modele_intermediaire(lot).numpy())

        matrice_features_intermediaires = np.concatenate(blocs_intermediaires, axis=0)

    display(pd.DataFrame({
        "Embeddings": ["Habituels : layer4", "Intermédiaires : layer3"],
        "Nombre d'images": [len(matrice_features), len(matrice_features_intermediaires)],
        "Dimensions du vecteur": [matrice_features.shape[1], matrice_features_intermediaires.shape[1]],
    }))''',
))

insert_after('1454a0b8', pair(
    'mentor-intermediaire-ari-explication', 'mentor-intermediaire-ari',
    '''### Comparer les embeddings habituels et intermédiaires

    Tester `matrice_features_intermediaires` avec la même méthode : standardisation, PCA à deux dimensions et K-Means à deux groupes appris uniquement sur `sans_label_train`.
    Calculer l'ARI sur les mêmes 79 images fortes d'entraînement et de validation (`donnees_ari`), puis afficher les deux scores. Les 20 images fortes de test restent réservées aux CNN.
    Cet essai mesure sa durée et son pic de RAM ; il ne remplace pas les groupes K-Means déjà calculés pour la suite.''',
    '''with mesurer_ressources("Comparaison ARI des embeddings layer3"):
        standardisation_intermediaire = StandardScaler()
        X_intermediaire_standardise = standardisation_intermediaire.fit_transform(
            matrice_features_intermediaires[sans_label_train.index]
        )
        pca_intermediaire = PCA(n_components=2, random_state=42)
        X_intermediaire_pca = pca_intermediaire.fit_transform(X_intermediaire_standardise)
        kmeans_intermediaire = KMeans(n_clusters=2, random_state=42)
        kmeans_intermediaire.fit(X_intermediaire_pca)

        X_ari_intermediaire = pca_intermediaire.transform(
            standardisation_intermediaire.transform(
                matrice_features_intermediaires[donnees_ari.index]
            )
        )
        ari_intermediaire = adjusted_rand_score(
            labels_reference, kmeans_intermediaire.predict(X_ari_intermediaire)
        )

    comparaison_embeddings = pd.DataFrame({
        "Embeddings": ["Habituels : layer4", "Intermédiaires : layer3"],
        "Dimensions": [512, 256],
        "ARI train + validation": [ari_kmeans, ari_intermediaire],
    })
    display(comparaison_embeddings.round(4))''',
))

# Time each computation independently so a pause between cells is not counted.
profile = {
    '375b1310': ('Préparation et égalisation des images', 'd53f9bbb'),
    '2690bd6b': ('Chargement du modèle d’extraction', '4ac06b74'),
    '0f6b5cdd': ('Extraction des embeddings layer4', '5219907b'),
    'd169025f': ('Création du tableau de features', 'a7319dfa'),
    'sauvegarde-features-csv': ('Sauvegarde du tableau CSV', None),
    '5597b309': ('Lecture du CSV et séparation des ensembles', '3dd62a34'),
    'd5c99912': ('Standardisation des embeddings', '0d26f789'),
    '1815ac75': ('PCA des embeddings', 'a7463040'),
    'd4bb2089': ('Clustering K-Means', 'de6350e6'),
    'b4c72202': ('Essais DBSCAN', 'a4f6f460'),
    '1454a0b8': ('ARI des embeddings layer4', 'd6cd3bd7'),
    'bca25a5c': ('Attribution des labels faibles', '6324dcb3'),
    'ba2b3b8a': ('Préparation des chargeurs CNN', '280a88b9'),
    'cc73ef88': ('Création des deux CNN', '668b5564'),
    'hyperparametres-semi-recherche': ('Hyperparamètres faibles puis forts du CNN semi-supervisé', 'hyperparametres-semi-explication'),
    'b184ba1a': ('Entraînement faible et évaluation', '4709dc64'),
    '82faeff6': ('Entraînement fort semi-supervisé et évaluation', 'e1ec12cb'),
    'hyperparametres-supervise-recherche': ('Hyperparamètres du CNN supervisé', None),
    '9dfb6484': ('Entraînement supervisé et évaluation', 'f5d694fb'),
}
for cell_id, (name, markdown_id) in profile.items():
    code = ''.join(cells[cell_id]['source']).strip('\n')
    cells[cell_id]['source'] = source(
        f'with mesurer_ressources({name!r}):\n' + textwrap.indent(code, '    ')
    )
    cells[cell_id]['outputs'] = []
    cells[cell_id]['execution_count'] = None
    if markdown_id:
        markdown = ''.join(cells[markdown_id]['source']).rstrip()
        cells[markdown_id]['source'] = source(markdown + '\n\nLa durée et le pic de RAM de cette cellule sont mesurés.')

# Keep the two Markdown texts explicitly requested by the user unchanged.
assert cells['sauvegarde-features-csv-explication'] == next(c for c in original['cells'] if c['id'] == 'sauvegarde-features-csv-explication')
assert cells['hyperparametres-supervise-explication'] == next(c for c in original['cells'] if c['id'] == 'hyperparametres-supervise-explication')

benchmark = ''.join(cells['a3f5cc5d']['source'])
begin = benchmark.index('modele_supervise.eval()')
end = benchmark.index('duree_labellisation = time.perf_counter() - debut_labellisation')
work = benchmark[begin:end].replace('debut_labellisation = time.perf_counter()\n\n', '')
replacement = (
    benchmark[:begin]
    + 'with mesurer_ressources("Prédiction des images finales", "Traitement final"):\n'
    + textwrap.indent(work.rstrip(), '    ') + '\n\n'
    + 'duree_labellisation = mesures_ressources["Prédiction des images finales"]["Durée (s)"]\n'
    + benchmark[end:].split('\n', 1)[1]
)
cells['a3f5cc5d']['source'] = source(replacement)
cells['a3f5cc5d']['outputs'] = []
cells['a3f5cc5d']['execution_count'] = None
cells['cbac9625']['source'] = source(
    ''.join(cells['cbac9625']['source']).rstrip()
    + '\n\nLe même prétraitement avec égalisation est utilisé. La durée et le pic de RAM du processus Python sont mesurés séparément de la création du modèle.'
)

insert_after('a3f5cc5d', pair(
    'mentor-bilan-ressources-explication', 'mentor-bilan-ressources',
    '''### Durée et RAM : création du modèle et traitement final

    Afficher le détail de `mesures_ressources`, puis un bilan en deux lignes.
    La création comprend le prétraitement, les deux extractions, les essais de clustering, les recherches d'hyperparamètres, les entraînements et leurs évaluations. Le traitement final porte sur l'échantillon de 1 000 images de la cellule précédente.
    Additionner les durées des cellules de calcul et retenir le plus grand pic de RAM pour chaque groupe : les pics ne s'additionnent pas. La RAM au départ comprend les tableaux et modèles déjà en mémoire ; la hausse maximale distingue les allocations supplémentaires. Le pic est une estimation par échantillonnage, pas la mémoire exacte des poids du modèle.''',
    '''details_ressources = pd.DataFrame(mesures_ressources.values())
    display(details_ressources.round(2))

    bilan_ressources = details_ressources.groupby("Groupe", sort=False).agg({
        "Durée (s)": "sum",
        "Pic RAM (Mio)": "max",
    })
    bilan_ressources["Durée (min)"] = bilan_ressources["Durée (s)"] / 60
    display(bilan_ressources.round(2))''',
))

# Verify unaffected earlier exploration, existing metadata and untouched cells.
for before in original['cells'][:28]:
    assert cells[before['id']] == before
for key, value in original.items():
    if key != 'cells':
        assert notebook[key] == value
for cell in notebook['cells']:
    if cell['cell_type'] == 'code':
        ast.parse(''.join(cell['source']))
for i, cell in enumerate(notebook['cells']):
    if cell['id'].startswith('mentor-') and cell['cell_type'] == 'markdown':
        assert notebook['cells'][i+1]['cell_type'] == 'code'

(BUILD / 'edited-cell-ids.json').write_text(json.dumps([
    c['id'] for c in notebook['cells'] if c['id'] not in cells or c != next(
        old for old in original['cells'] if old['id'] == c['id']
    )
], indent=2), encoding='utf-8')
PATH.write_text(json.dumps(notebook, ensure_ascii=False, indent=1) + '\n', encoding='utf-8')
print('Added requested notebook cells; earlier exploration preserved; all Python cells compile.')
