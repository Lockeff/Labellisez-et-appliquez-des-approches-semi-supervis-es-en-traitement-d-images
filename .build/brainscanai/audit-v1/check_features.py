import json
import sys
from contextlib import nullcontext
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.model_selection import train_test_split
from sklearn.metrics import adjusted_rand_score

n=json.loads(Path('01_exploration_dataset.ipynb').read_text(encoding='utf-8'))
env={'np':np,'pd':pd,'StandardScaler':StandardScaler,'PCA':PCA,'KMeans':KMeans,
     'train_test_split':train_test_split,'mesurer_ressources':lambda *a,**k:nullcontext(),
     'display':lambda *a:None,'chemin_features':Path('resultats/etape2/features_resnet18.csv'),
     'colonnes_features':[f'feature_{i:03d}' for i in range(512)]}
for cid in ['5597b309','d5c99912','1815ac75','d4bb2089']:
    c=next(c for c in n['cells'] if c.get('id')==cid)
    exec(''.join(c['source']),env)
f=env['features']; x=f[env['colonnes_features']].to_numpy()
print('CSV shape',f.shape,'missing numeric values',int(pd.isna(x).sum()),'non-finite numeric values',int((~np.isfinite(x)).sum()),'duplicate paths',int(f['chemin'].duplicated().sum()))
groups=env['kmeans'].predict(env['X_fort_pca'])
print('\nStrong TRAIN ONLY, group correspondence:')
print(pd.crosstab(env['fortes_train']['label_fort'].to_numpy(),groups,rownames=['Known label'],colnames=['K-Means group']))
both=np.concatenate([env['X_fort_pca'],env['X_validation_pca']])
labels=pd.concat([env['fortes_train'],env['fortes_validation']])['label_fort']
print('ARI train + validation:',adjusted_rand_score(labels,env['kmeans'].predict(both)))
for a,b in [('fortes_train','fortes_test'),('fortes_train','fortes_validation'),('sans_label_train','sans_label_test')]:
    print(a,b,'path overlap',len(set(env[a]['chemin']) & set(env[b]['chemin'])))
print('PCA explained variance ratio sum',float(env['pca'].explained_variance_ratio_.sum()))
