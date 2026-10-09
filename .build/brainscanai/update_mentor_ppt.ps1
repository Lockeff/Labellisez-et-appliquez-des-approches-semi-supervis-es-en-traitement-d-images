param([string]$Workspace=(Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/mentor-ppt'
[void][IO.Directory]::CreateDirectory($build)
$source=Join-Path $Workspace 'output/presentations/BrainScanAI_presentation_13_slides_couts.pptx'
$candidate=Join-Path $build 'candidate-com.pptx'
if(Test-Path -LiteralPath $candidate){throw 'Candidate already exists.'}
$notebookHash=(Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash
$helperSource=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$helperStart=$helperSource.IndexOf('function RGB(')
$helperEnd=$helperSource.IndexOf('function Picture(')
Invoke-Expression $helperSource.Substring($helperStart,$helperEnd-$helperStart)
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,0,0,0)
function Heading($Slide,$Title){
 for($i=$Slide.Shapes.Count;$i -ge 1;$i--){$Slide.Shapes.Item($i).Delete()}
 [void](Text $Slide $Title 48 30 866 80 32 $true)
 [void](Line $Slide 48 114 912 114)
 [void](Text $Slide ([string]$Slide.SlideIndex) 878 508 34 18 12 $false $gray 3)
}
function CellText($Table,$Row,$Column,$Value){$Table.Cell($Row,$Column).Shape.TextFrame.TextRange.Text=[string]$Value}
try {
 $s=$deck.Slides.Item(3)
 Heading $s 'Prétraitement et extraction des features'
 [void](Text $s 'Égalisation des histogrammes : redistribuer les intensités des pixels pour ajuster le contraste, avant la normalisation.' 48 135 864 58 22)
 $rows=@(
  @('Ordre du traitement','Opération réalisée'),
  @('1. Égalisation','Appliquer equalize aux images RGB en uint8.'),
  @('2. Prétraitement ImageNet','Redimensionner à 256, recadrer à 224 × 224, normaliser.'),
  @('3. ResNet18 gelé','Extraire les caractéristiques avec les poids IMAGENET1K_V1.'),
  @('4. Tableau de features','Conserver un vecteur de 512 valeurs et le chemin par image.')
 )
 $table=Table $s $rows 48 207 864 212 @(288,576) 19
 for($r=1;$r -le 5;$r++){$table.Table.Cell($r,2).Shape.TextFrame.TextRange.ParagraphFormat.Alignment=1}
 [void](Text $s '1 410 images : un CSV de 1 410 lignes, avec 512 features et le chemin.' 48 438 864 32 22 $true)
 Foot $s 'Le traitement conserve les trois canaux RGB. Il ne supprime pas les images colorées ni les fichiers sources.'
 Note $s @'
Source : 01_exploration_dataset.ipynb, étapes 2. Cellules mentor-egalisation, 375b1310, 2690bd6b, 0f6b5cdd et sauvegarde-features-csv. La fonction torchvision.transforms.functional.equalize traite les pixels uint8 RGB puis poids.transforms() applique le prétraitement ImageNet. Ce traitement est commun à l'extraction des embeddings et aux images utilisées par les CNN, ainsi qu'à l'inférence finale. Il ne convertit pas en niveaux de gris et n'élimine pas une catégorie d'images. Aucun bénéfice clinique ni gain isolé de performance de l'égalisation n'a été mesuré. La comparaison des couches utilise le même prétraitement égalisé. Le CSV principal conserve les embeddings habituels layer4, 512 dimensions.
'@

 $s=$deck.Slides.Item(6)
 Heading $s 'Test de la couche intermédiaire et ARI'
 [void](Text $s 'Comparer deux représentations des mêmes 1 410 images. Après chaque couche, une moyenne spatiale produit un vecteur par image.' 48 134 864 58 21)
 $rows=@(
  @('Extraction ResNet18','Dimensions du vecteur','ARI sur 79 labels forts'),
  @('Habituelle : layer4','512','0,3773'),
  @('Intermédiaire : layer3','256','0,0339')
 )
 [void](Table $s $rows 48 208 864 126 @(336,244,284) 21)
 [void](Text $s 'Même protocole : standardisation, PCA à 2 dimensions, K-Means à 2 groupes. Apprentissage sur le train sans label uniquement.' 48 355 864 55 21)
 [void](Text $s 'Choix : conserver layer4. Son ARI montre une meilleure correspondance avec les labels connus, qui reste partielle.' 48 426 864 53 22 $true)
 Foot $s "ARI : train + validation forts. Test fort exclu. L'essai layer3 sert au clustering. Les CNN apprennent sur les images."
 Note $s @'
Source : 01_exploration_dataset.ipynb, mentor-intermediaire-extraction, mentor-intermediaire-ari et comparaison_embeddings. Extraction habituelle : conv1 à layer4, moyenne spatiale et aplatissement, 512 nombres par image. Extraction intermédiaire : conv1 à layer3, AdaptiveAvgPool2d((1,1)) et Flatten(), 256 nombres par image. Les deux matrices contiennent les mêmes 1410 images, dans le même ordre. Chaque représentation a son propre StandardScaler, sa PCA et son KMeans, ajustés uniquement sur les mêmes 1048 images sans_label_train. Pour l'ARI, les transformations utilisent transform et les groupes predict sur les 79 images fortes de train et validation (63+16). Les 20 images fortes de test ne servent pas à ce choix. Résultats affichés : layer4 0.3772887757615568, layer3 0.033925049309664695. 0.3773 est un indice ajusté pour le hasard, pas 37.73% de bonnes prédictions. Le CSV principal et la labellisation faible continuent d'utiliser layer4. La matrice intermédiaire est un essai distinct de clustering, pas une nouvelle entrée du CNN. Les groupes 0 et 1 ne signifient pas automatiquement normal et cancer. Cette expérience ne démontre pas qu'une couche est meilleure pour tous les jeux de données.
'@

 # Actualiser les résultats dépendant du nouveau prétraitement.
 $s=$deck.Slides.Item(5)
 $s.Shapes.Item('TextBox 7').TextFrame.TextRange.Text='Groupe 0 : 565 images. Groupe 1 : 483.'
 $counts=@(93,36,2,15,8,12,1,1,1)
 $noise=@(368,706,1028,51,112,373,8,10,18)
 $t=$s.Shapes.Item('Table 10').Table
 for($r=2;$r -le 10;$r++){CellText $t $r 3 $counts[$r-2];CellText $t $r 4 $noise[$r-2]}
 $sh=$s.Shapes.Item('TextBox 11')
 $sh.TextFrame.TextRange.Text='2 groupes à eps = 0,5 / minimum = 10, mais 1 028 images hors groupe.'
 $sh.TextFrame.TextRange.Font.Size=16
 $sh.Height=[single]44
 $sh=$s.Shapes.Item('TextBox 13')
 $sh.TextFrame.TextRange.Text='K-Means attribue un label faible à chaque image du train sans label.'
 $sh.Left=[single]48;$sh.Top=[single]510;$sh.Width=[single]805;$sh.Height=[single]18
 $sh.TextFrame.TextRange.Font.Size=12
 Note $s 'Source : notebook enregistré, étapes 3.5 à 3.9, après égalisation. Graphique natif : coordonnées PCA recalculées depuis le CSV enregistré, avec les mêmes splits, StandardScaler, PCA(n_components=2, random_state=42), KMeans(n_clusters=2, random_state=42). Groupes : 565 et 483. Tableau : neuf essais DBSCAN du notebook. eps=0.5/min_samples=10 produit deux groupes mais 1028/1048 points hors groupe. K-Means conserve donc la labellisation faible utilisée par la suite.'

 $s=$deck.Slides.Item(7)
 $s.Shapes.Item('TextBox 3').TextFrame.TextRange.Text='7'
 $sh=$s.Shapes.Item('TextBox 4')
 $sh.TextFrame.TextRange.Text="D'abord les époques faibles, puis la grille forte pour les deux méthodes."
 $sh.Height=[single]55;$sh.TextFrame.TextRange.Font.Size=19
 $t=$s.Shapes.Item('Table 9').Table
 CellText $t 1 2 'Accuracy faible'
 CellText $t 2 2 '1,0000';CellText $t 3 2 '1,0000'
 $t=$s.Shapes.Item('Table 5').Table
 $semi=@('0,5682','0,5682','0,7143','0,9524')
 $super=@('0,5000','0,6250','0,6579','0,9756')
 for($r=2;$r -le 5;$r++){CellText $t $r 3 $semi[$r-2];CellText $t $r 4 $super[$r-2]}
 Note $s 'Source : notebook enregistré, hyperparametres-semi-recherche et hyperparametres-supervise-recherche, après égalisation. Phase faible : lr=0.001, 5 ou 10 époques, accuracy=1.0 sur les groupes K-Means des 16 images de validation existantes. Ce score mesure l''accord aux pseudo-labels et non la justesse médicale. Égalité : conserver 5 époques. La recherche forte semi reprend ce meilleur modèle faible pour chaque essai. Grille forte commune : lr=0.0001 ou 0.001, 5 ou 10 époques, comparaison du F2 cancer sur les mêmes 16 labels forts de validation. Choix : lr=0.001 et 10 époques dans les deux méthodes. F2 de validation : 0.9524 semi, 0.9756 supervisé. Test final distinct : 20 images. Les numéros de page reflètent la position réelle des slides.'

 $s=$deck.Slides.Item(8)
 $s.Shapes.Item('TextBox 3').TextFrame.TextRange.Text='8'
 $s.Shapes.Item('TextBox 8').TextFrame.TextRange.Text="Le même CNN poursuit son entraînement.`nTest faible : accuracy de 0,977 sur 263 images (labels faibles)."
 Note $s 'Source : notebook enregistré, entraînement faible, fort semi-supervisé et supervisé après égalisation. 1048 labels faibles, 5 époques, puis 63 labels forts, 10 époques. Taux=0.001 pour les trois phases finales. Test faible sur 263 images avec cibles K-Means : accuracy arrondie 0.977, mesure de reproduction des groupes. Les deux modèles finaux utilisent le même test fort de 20 images. Les images reçoivent le prétraitement égalisé. L''essai des embeddings layer3 reste limité au clustering et ne modifie pas le CNN final.'

 $s=$deck.Slides.Item(9)
 $t=$s.Shapes.Item('Table 4').Table
 $semi=@('0,850','0,842','0,889','0,800','0,816')
 $super=@('0,900','0,900','0,900','0,900','0,900')
 for($r=2;$r -le 6;$r++){CellText $t $r 2 $semi[$r-2];CellText $t $r 3 $super[$r-2]}
 $t=$s.Shapes.Item('Table 7').Table
 CellText $t 2 2 '9';CellText $t 2 3 '1';CellText $t 3 2 '2';CellText $t 3 3 '8'
 $t=$s.Shapes.Item('Table 10').Table
 CellText $t 2 2 '9';CellText $t 2 3 '1';CellText $t 3 2 '1';CellText $t 3 3 '9'
 Note $s 'Source : notebook enregistré, comparaison finale après égalisation. Test fort=20 images, 10 normal et 10 cancer, jamais utilisées pour entraîner ou sélectionner les CNN. Semi : accuracy .85, F1 .8421052631578947, précision .8888888888888888, rappel .8, F2 .8163265306122449, confusion [[9,1],[2,8]]. Supervisé : tous ces scores .9, confusion [[9,1],[1,9]]. Le supervisé reconnaît une image cancéreuse de plus et produit le même nombre de faux positifs. Ce petit test est exploratoire et ne constitue pas une validation clinique.'

 $s=$deck.Slides.Item(10)
 $sh=$s.Shapes.Item('TextBox 11')
 $sh.TextFrame.TextRange.Text="Machine utilisée : i9-14900K, CPU local`n`nAprès égalisation : 1 000 images en 7,4 s`n`nEstimation pour 4 millions : 8,2 heures"
 $sh.Width=[single]820;$sh.Height=[single]200
 Note $s 'Source : notebook enregistré, a3f5cc5d. Prédiction locale de 1000 images en 7.3758490000036545 secondes, lecture RGB, égalisation, prétraitement ImageNet, lots de 16 et CNN supervisé déjà entraîné. Extrapolation linéaire : 8.19538777778184 h pour 4 M. Il s''agit d''une estimation locale, pas d''un benchmark AWS. Les coûts des slides suivantes conservent cette durée comme hypothèse à vérifier.'

 $deck.SaveAs($candidate,24)
 if((Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash -ne $notebookHash){throw 'Notebook changed.'}
 [pscustomobject]@{Source=$source;Candidate=$candidate;NotebookHash=$notebookHash;Slides=$deck.Slides.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
 Write-Output $candidate
} finally {$deck.Close()}
