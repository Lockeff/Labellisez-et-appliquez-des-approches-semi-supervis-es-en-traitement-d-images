param([string]$Workspace = (Get-Location).Path)
$ErrorActionPreference = 'Stop'
$build = Join-Path $Workspace '.build/brainscanai'
$data = Get-Content -LiteralPath (Join-Path $build 'sources.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$outDir = Join-Path $Workspace 'output/presentations'
[void][System.IO.Directory]::CreateDirectory($outDir)
$out = Join-Path $build 'candidate-v3.pptx'
if (Test-Path -LiteralPath $out) { throw 'Output already exists. Use a new filename for a revision.' }
$renders = Join-Path $build 'renders'
[void][System.IO.Directory]::CreateDirectory($renders)

function RGB([string]$Hex) {
    $h = $Hex.TrimStart('#')
    return [Convert]::ToInt32($h.Substring(0,2),16) + 256*[Convert]::ToInt32($h.Substring(2,2),16) + 65536*[Convert]::ToInt32($h.Substring(4,2),16)
}
$navy = RGB '172F43'
$teal = RGB '007F82'
$blue = RGB '3978A9'
$orange = RGB 'D58A3D'
$gray = RGB '536473'
$light = RGB 'EDF3F5'
$white = RGB 'FFFFFF'
$muted = RGB 'CAD7DF'
$font = 'Arial'

function Text($Slide, [string]$Value, [double]$X, [double]$Y, [double]$W, [double]$H, [double]$Size=20, [bool]$Bold=$false, [int]$Color=$navy, [int]$Align=1) {
    $shape = $Slide.Shapes.AddTextbox(1,$X,$Y,$W,$H)
    $shape.TextFrame.MarginLeft=0; $shape.TextFrame.MarginRight=0
    $shape.TextFrame.MarginTop=0; $shape.TextFrame.MarginBottom=0
    $shape.TextFrame.WordWrap=-1; $shape.TextFrame.AutoSize=0
    $range=$shape.TextFrame.TextRange
    $range.Text=$Value; $range.Font.Name=$font; $range.Font.Size=$Size
    $range.Font.Bold=$(if($Bold){-1}else{0}); $range.Font.Color.RGB=$Color
    $range.LanguageID=1036
    $range.ParagraphFormat.Alignment=$Align
    $range.ParagraphFormat.SpaceBefore=0; $range.ParagraphFormat.SpaceAfter=0
    $shape.TextFrame2.AutoSize=0
    $shape.Height=[single]$H; $shape.Width=[single]$W
    return $shape
}
function Line($Slide,$X1,$Y1,$X2,$Y2,$Color=$muted,$Width=1.2) {
    $s=$Slide.Shapes.AddLine($X1,$Y1,$X2,$Y2)
    $s.Line.ForeColor.RGB=$Color; $s.Line.Weight=[single]$Width
    return $s
}
function Arrow($Slide,$X1,$Y1,$X2,$Y2,$Color=$teal) {
    $s=$Slide.Shapes.AddConnector(1,$X1,$Y1,$X2,$Y2)
    $s.Line.ForeColor.RGB=$Color; $s.Line.Weight=2
    $s.Line.EndArrowheadStyle=3
    return $s
}
function Node($Slide,[string]$Value,$X,$Y,$W,$H,$Color=$teal,$Size=18) {
    $s=$Slide.Shapes.AddShape(1,$X,$Y,$W,$H)
    $s.Fill.ForeColor.RGB=$white; $s.Line.ForeColor.RGB=$Color; $s.Line.Weight=1.5
    $s.TextFrame.MarginLeft=10; $s.TextFrame.MarginRight=10
    $s.TextFrame.MarginTop=6; $s.TextFrame.MarginBottom=6
    $s.TextFrame.VerticalAnchor=3; $s.TextFrame.WordWrap=-1; $s.TextFrame.AutoSize=0
    $t=$s.TextFrame.TextRange; $t.Text=$Value; $t.Font.Name=$font
    $t.Font.Size=$Size; $t.Font.Color.RGB=$navy; $t.ParagraphFormat.Alignment=2
    $t.LanguageID=1036
    return $s
}
function NewSlide([string]$Title) {
    $s=$deck.Slides.Add($deck.Slides.Count+1,12)
    $s.FollowMasterBackground=0; $s.Background.Fill.ForeColor.RGB=$white
    if($Title) {
        [void](Text $s $Title 48 30 866 80 32 $true)
        [void](Line $s 48 114 912 114)
    }
    [void](Text $s ([string]$s.SlideIndex) 878 508 34 18 12 $false $gray 3)
    return $s
}
function Note($Slide,[string]$Value) {
    $Slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text=$Value
}
function Foot($Slide,[string]$Value) { [void](Text $Slide $Value 48 503 805 25 11 $false $gray) }
function Table($Slide,$Rows,$X,$Y,$W,$H,$Widths,$Size=17) {
    $nr=$Rows.Count; $nc=$Rows[0].Count
    $shape=$Slide.Shapes.AddTable($nr,$nc,$X,$Y,$W,$H)
    $t=$shape.Table
    if($Widths) { for($c=1;$c -le $nc;$c++){ $t.Columns.Item($c).Width=$Widths[$c-1] } }
    for($r=1;$r -le $nr;$r++) {
        $t.Rows.Item($r).Height=[single]($H/$nr)
        for($c=1;$c -le $nc;$c++) {
            $cell=$t.Cell($r,$c).Shape
            $cell.Fill.ForeColor.RGB=$(if($r -eq 1){$navy}elseif($r%2 -eq 0){$light}else{$white})
            $cell.TextFrame.MarginLeft=8; $cell.TextFrame.MarginRight=7
            $cell.TextFrame.MarginTop=3; $cell.TextFrame.MarginBottom=3
            $cell.TextFrame.VerticalAnchor=3
            $tr=$cell.TextFrame.TextRange; $tr.Text=[string]$Rows[$r-1][$c-1]
            $tr.Font.Name=$font; $tr.Font.Size=$Size; $tr.LanguageID=1036
            $tr.Font.Bold=$(if($r -eq 1){-1}else{0})
            $tr.Font.Color.RGB=$(if($r -eq 1){$white}else{$navy})
            $tr.ParagraphFormat.Alignment=$(if($c -eq 1){1}else{2})
            for($edge=1;$edge -le 4;$edge++){ $t.Cell($r,$c).Borders.Item($edge).ForeColor.RGB=$white; $t.Cell($r,$c).Borders.Item($edge).Weight=1 }
        }
    }
    return $shape
}
function Picture($Slide,[string]$Path,$X,$Y,$W,$H) {
    return $Slide.Shapes.AddPicture($Path,0,-1,$X,$Y,$W,$H)
}
function Scatter($Slide,$X,$Y,$W,$H) {
    $shape=$Slide.Shapes.AddShape(1,$X,$Y,$W,$H)
    $shape.Name='KMeansChartPlaceholder'
    $shape.Fill.ForeColor.RGB=$white
    $shape.Line.Visible=0
    return $shape
}
function Confusion($Slide,$Matrix,[string]$Label,$X,$Y) {
    [void](Text $Slide $Label $X $Y 176 28 19 $true)
    [void](Text $Slide 'Prédiction' ($X+34) ($Y+36) 142 22 13 $false $gray 2)
    $rows=@(@('Vrai','normal','cancer'),@('normal',[string]$Matrix[0][0],[string]$Matrix[0][1]),@('cancer',[string]$Matrix[1][0],[string]$Matrix[1][1]))
    $matrixShape=Table $Slide $rows $X ($Y+63) 184 120 @(62,61,61) 14
    for($r=1;$r -le 3;$r++) {
        for($c=1;$c -le 3;$c++) {
            $matrixShape.Table.Cell($r,$c).Shape.TextFrame.MarginLeft=4
            $matrixShape.Table.Cell($r,$c).Shape.TextFrame.MarginRight=4
        }
        $matrixShape.Table.Rows.Item($r).Height=[single]40
    }
}

$ppt = New-Object -ComObject PowerPoint.Application
$openBefore=$ppt.Presentations.Count
$ppt.DisplayAlerts=1
$deck=$ppt.Presentations.Add(0)
$deck.PageSetup.SlideWidth=960; $deck.PageSetup.SlideHeight=540
try {
    # 1. Mission
    $s=NewSlide ''
    [void](Text $s 'BrainScanAI' 48 58 550 74 52 $true)
    [void](Text $s "Features, clustering et`napprentissage semi-supervisé" 48 150 560 100 27)
    [void](Text $s 'CurelyticsIA, première phase de recherche' 48 272 550 30 18 $false $gray)
    [void](Text $s 'Objectif : explorer la labellisation automatique des images cérébrales à partir de peu de labels connus.' 48 329 554 95 22)
    [void](Picture $s $data.images.unlabelled 650 105 262 262)
    [void](Text $s 'Exemple du dossier sans label' 650 380 262 32 15 $false $gray 2)
    Foot $s 'Synthèse du notebook et de la fiche d''auto-évaluation'
    Note $s 'Sources : énoncé BrainScanAI fourni dans la conversation, contexte de AGENTS.md, 01_exploration_dataset.ipynb. La mission vise une exploration pour assister les professionnels de santé. Le dataset est nommé mri_dataset_brain_cancer_oc. Image originale du dossier sans_label, sans annotation diagnostique ajoutée. La partie financière sera développée ultérieurement à la demande de l''utilisateur.'

    # 2. Exploration and cleaning
    $s=NewSlide 'Données disponibles et nettoyage'
    [void](Text $s '1 506 fichiers JPEG recensés' 48 138 570 34 25 $true)
    $rows=@(@('Ensemble','Avant','Après filtrage'),@('Cancer','50','50'),@('Normal','50','49'),@('Sans label','1 406','1 311'),@('Total','1 506','1 410'))
    [void](Table $s $rows 48 186 555 185 @(245,125,185) 18)
    [void](Text $s '96 copies écartées de la liste de travail. Les images sources restent intactes.' 48 390 555 66 19)
    [void](Picture $s $data.images.normal 636 143 128 128)
    [void](Picture $s $data.images.cancer 786 143 128 128)
    [void](Text $s 'Label normal' 636 281 128 24 15 $false $gray 2)
    [void](Text $s 'Label cancer' 786 281 128 24 15 $false $gray 2)
    [void](Text $s "Échantillon de 100 images`n512 × 512 pixels, RGB`nAucune erreur de lecture" 636 329 278 110 19)
    Foot $s 'Le descriptif annonce 1 500 JPEG. L''énoncé annonce des PNG.'
    Note $s 'Sources : notebook, étape 1, comptage, caractéristiques de l''échantillon et filtrage des doublons. Après filtrage : 50 cancer, 49 normal, 1311 sans label. Le total 96 comprend 64 copies internes à sans_label, 1 copie interne à normal, et 31 copies supplémentaires détectées par comparaison entre dossiers. Recherche fondée sur les pixels RGB et les dimensions. Contrôle des dimensions et de la lisibilité sur un échantillon de 100 images. Les différences d''orientation et de cadrage observées ne justifient pas une suppression automatique. Fiche d''auto-évaluation, page 2 : nettoyage, cohérence, biais potentiels.'

    # 3. Features
    $s=NewSlide 'Prétraitement et extraction des features'
    [void](Node $s "Prétraitement`nImageNet" 48 158 250 106 $teal 23)
    [void](Node $s "ResNet18`nParamètres gelés" 355 158 250 106 $teal 23)
    [void](Node $s "512 features`npar image" 662 158 250 106 $teal 23)
    [void](Arrow $s 304 211 348 211)
    [void](Arrow $s 611 211 655 211)
    [void](Text $s "Redimensionnement à 256`nRecadrage central à 224 × 224`nNormalisation prévue par le modèle" 48 286 270 110 18)
    [void](Text $s "Poids IMAGENET1K_V1`n0 paramètre entraînable`nLa couche finale devient Identity" 355 286 273 110 18)
    [void](Text $s "Vecteur numérique descriptif`nExtraction par lots de 16`n1 410 lignes dans le tableau" 662 286 250 110 18)
    [void](Text $s 'Le CSV conserve le chemin de l''image et les 512 features. L''étape 3 recharge ce tableau.' 48 430 864 53 21 $true)
    Note $s 'Source : notebook, étape 2 et Markdown associés. Resize_size=256 puis crop_size=224, conversion RGB et float32. Normalisation : mean=[0.485,0.456,0.406], std=[0.229,0.224,0.225]. ResNet18_Weights.IMAGENET1K_V1. requires_grad=False sur tous les paramètres. fc=Identity et eval(). Matrice 1410×512 et DataFrame 1410×513 avec chemin. CSV : resultats/etape2/features_resnet18.csv, sauvegardé avec index=False et rechargé avec pd.read_csv. Choix : réutiliser la représentation du modèle pré-entraîné, conformément à l''énoncé. Fiche d''auto-évaluation, page 2 : transformations adaptées et embeddings exploitables.'

    # 4. Separation
    $s=NewSlide 'Ensembles séparés avant le clustering'
    [void](Text $s 'Données fortement labellisées : 99 images' 48 139 864 30 22 $true $blue)
    [void](Node $s 'Entraînement : 63' 48 182 247 64 $blue 21)
    [void](Node $s 'Validation : 16' 356 182 247 64 $blue 21)
    [void](Node $s 'Test : 20' 665 182 247 64 $blue 21)
    [void](Text $s 'Split stratifié selon normal/cancer. Le test contient 10 images de chaque classe.' 48 260 864 43 18)
    [void](Text $s 'Données sans label : 1 311 images' 48 320 864 30 22 $true $teal)
    [void](Node $s 'Entraînement : 1 048' 48 364 399 64 $teal 21)
    [void](Node $s 'Test : 263' 513 364 399 64 $teal 21)
    [void](Text $s 'StandardScaler et PCA apprennent uniquement sur le train sans label. Les autres ensembles utilisent transform.' 48 451 864 48 19)
    Note $s 'Source : notebook, cellules de séparation, standardisation et PCA actuelles. train_test_split réserve 20 % de test fort, puis 20 % de validation parmi les données fortes restantes. random_state=42. Le split sans label réserve 20 % pour le test, sans stratification car les labels sont encore inconnus. StandardScaler.fit_transform et PCA.fit_transform utilisent sans_label_train uniquement. Les données fortes et le test sans label utilisent transform séparément. À l''étape clustering, les labels forts servent à l''ARI. À l''étape CNN, fortes_train sert à l''entraînement, fortes_validation à la sélection des paramètres et fortes_test au test final. Aucun test ne sert à l''entraînement. Fiche page 2 : préparation, encodage et évaluation.'

    # 5. Clustering
    $s=NewSlide 'Clustering sur les coordonnées PCA'
    [void](Text $s 'K-Means : deux groupes' 48 136 445 32 23 $true)
    [void](Text $s 'PCA : 512 features vers 2 coordonnées. Chaque point représente une image.' 48 176 445 48 18)
    [void](Scatter $s 39 225 473 248)
    [void](Text $s ('Groupe 0 : '+$data.kmeans.counts[0]+' images. Groupe 1 : '+$data.kmeans.counts[1]+'.') 48 471 464 29 16 $false $gray)
    [void](Text $s 'DBSCAN : neuf essais' 552 136 360 32 23 $true)
    [void](Text $s 'eps : rayon ; minimum : points voisins.' 552 176 360 24 16 $false $gray)
    $rows=New-Object System.Collections.Generic.List[object]
    $rows.Add(@('eps','Minimum','Groupes','Hors groupe'))
    foreach($trial in $data.dbscan){ $rows.Add(@(([string]$trial.eps).Replace('.',','),[string]$trial.min_samples,[string]$trial.nombre_groupes,[string]$trial.nombre_hors_groupe)) }
    [void](Table $s $rows.ToArray() 552 205 360 250 @(48,100,90,122) 16)
    [void](Text $s 'Aucun essai ne produit deux groupes. K-Means fournit les labels faibles retenus.' 552 465 360 42 17)
    Note $s ('Source : cellules actuelles de l''étape 3, recalculées en mémoire depuis le CSV pour ce support. Le notebook reste intact. Les anciennes sorties conservées dans le notebook précèdent la modification de standardisation/PCA. Le graphique est un nuage de points natif éditable, avec toutes les coordonnées du train sans label. K-Means n_clusters=2, random_state=42. DBSCAN teste eps=[0.5,1.0,2.0] et min_samples=[3,5,10]. -1 signifie hors groupe. Les neuf essais restent exploratoires. Aucun essai actuel ne produit deux groupes. La conclusion suit le Markdown : retenir K-Means pour la labellisation faible. Le nombre de groupes demandé ne garantit pas une correspondance aux classes médicales. Fiche page 2 : plusieurs méthodes de clustering. SHA-256 du notebook source : '+$data.notebook_sha256)

    # 6. Weak labels and ARI
    $s=NewSlide 'Labels faibles et contrôle par l''ARI'
    [void](Node $s 'Images sans label' 48 149 235 77 $teal 21)
    [void](Node $s 'Groupes K-Means 0 / 1' 355 149 250 77 $teal 20)
    [void](Node $s 'Colonne label_faible' 677 149 235 77 $teal 20)
    [void](Arrow $s 290 188 348 188)
    [void](Arrow $s 612 188 670 188)
    [void](Text $s 'Le train reçoit ses groupes. K-Means prédit ceux du test sans réentraînement.' 48 248 864 49 21)
    [void](Text $s 'Les numéros 0 et 1 n''identifient pas automatiquement normal et cancer.' 48 307 528 72 24 $true)
    [void](Text $s 'ARI' 676 304 235 32 23 $true $teal)
    $ariText=([double]$data.kmeans.ari).ToString('0.000',[Globalization.CultureInfo]::GetCultureInfo('fr-FR'))
    [void](Text $s $ariText 676 346 235 64 43 $true $navy)
    [void](Text $s 'Sur 16 images fortes de validation' 676 416 235 47 17 $false $gray)
    [void](Text $s 'L''ARI proche de zéro ne montre pas d''accord convaincant avec les labels connus.' 48 411 528 72 20)
    Note $s 'Source : étape 3, Markdown sur les groupes et labels faibles, cellules de labellisation et ARI. donnees_faibles contient exclusivement les anciennes images sans label. donnees_fortes reste séparé avec label_fort. faibles_train et faibles_test sont concaténés uniquement entre eux. L''ARI compare kmeans.predict(X_validation_pca) aux 16 labels forts réservés. Valeur recalculée : -0.026058631921824105. L''ARI mesure l''accord entre partitions et ne fournit pas une correspondance sémantique 0/1 avec normal/cancer. Les groupes peuvent refléter des différences visuelles. Le code CNN actuel reprend directement les groupes 0/1 avant la phase forte, sans association explicite aux classes médicales. Fiche page 2 : vérifier les sorties et interpréter les méthodes.'

    # 7. CNN protocol
    $s=NewSlide 'Comparaison des deux entraînements CNN'
    [void](Text $s 'Semi-supervisé' 48 139 240 30 24 $true $teal)
    [void](Node $s "1 048 labels faibles`n5 époques" 48 185 240 83 $teal 20)
    [void](Node $s "63 labels forts`n10 époques" 365 185 240 83 $teal 20)
    [void](Arrow $s 296 226 357 226)
    [void](Text $s "Le même CNN poursuit son entraînement.`nTest faible : accuracy de 0,973 sur 263 images (labels K-Means)." 48 282 564 58 17)
    [void](Text $s 'Supervisé' 48 351 240 30 24 $true $blue)
    [void](Node $s "63 labels forts`n10 époques" 365 342 240 83 $blue 20)
    [void](Node $s "Même test fort`n20 images" 704 253 208 101 $navy 23)
    [void](Line $s 613 226 696 226 $teal 2)
    [void](Line $s 696 226 696 277 $teal 2)
    [void](Arrow $s 696 277 704 277)
    [void](Line $s 613 383 696 383 $blue 2)
    [void](Line $s 696 383 696 329 $blue 2)
    [void](Arrow $s 696 329 704 329 $blue)
    [void](Text $s 'ResNet18 pré-entraîné, couches gelées. Seule la couche finale à deux sorties apprend : 1 026 paramètres.' 48 452 864 49 19)
    Note $s 'Source : étape 4. creer_cnn initialise les deux ResNet18 avec les mêmes poids ImageNet et torch.manual_seed(42). Les couches existantes sont gelées. fc=Linear(512,2), 1026 paramètres entraînables. Optimiseur Adam, CrossEntropyLoss, lots de 16. Le mode eval conserve les statistiques du réseau pré-entraîné et fc.train entraîne la tête. Semi-supervisé : 5 époques faibles puis 10 fortes retenues sur validation. Supervisé : 10 époques fortes retenues sur validation. Test fort de 20 images distinctes de l''entraînement. L''évaluation intermédiaire du CNN faible se fait sur 263 images faibles avec les labels K-Means : accuracy=0.973 et F1 groupe 1=0.968. Cela mesure la reproduction des groupes, pas une performance clinique. Fiche page 2 : comparaison avec et sans labellisation partielle.'

    # 8. Search and metric choices
    $s=NewSlide 'Hyperparamètres et métriques de validation'
    [void](Text $s 'Même grille pour les deux méthodes' 48 140 490 31 23 $true)
    $rows=@(@('Taux','Époques fortes','F2 semi','F2 supervisé'),@('0,0001','5','0,5682','0,7317'),@('0,0001','10','0,5682','0,7500'),@('0,001','5','0,9091','0,5405'),@('0,001','10','0,9756','0,7692'))
    [void](Table $s $rows 48 186 495 205 @(112,147,115,121) 17)
    [void](Text $s 'Les deux recherches retiennent 0,001 et 10 époques fortes.' 48 409 495 62 22 $true $teal)
    [void](Text $s 'Sélection : meilleur F2 sur 16 labels forts de validation. Le test final reste réservé.' 590 140 322 69 19 $true)
    [void](Text $s "Accuracy : bonnes prédictions.`n`nPrécision : fiabilité des prédictions cancer.`n`nRappel : cancers reconnus.`n`nF1 : équilibre précision/rappel.`n`nF2 : davantage de poids au rappel." 590 229 322 253 18)
    Note $s 'Source : recherches d''hyperparamètres semi-supervisée et supervisée et fonction evaluer_cnn de l''étape 4. Les valeurs viennent des tableaux enregistrés dans le notebook. Quatre combinaisons pour chaque méthode, sur la même validation forte de 16 images. 5 époques faibles fixes pour la méthode semi-supervisée. Chaque essai repart des mêmes poids. Le meilleur F2 détermine le taux et le nombre d''époques fortes, puis le modèle final repart des poids initiaux. Les deux méthodes retiennent taux=0.001 et 10 époques fortes. Métriques positives sur cancer=1 après la phase forte. F2 pondère davantage le rappel et pénalise plus fortement les faux négatifs dans la formule du notebook. Aucune priorité clinique ni seuil de réussite n''a été validé. Fiche page 2 : métriques pertinentes et ajustement des hyperparamètres.'

    # 9. Final results
    $s=NewSlide 'Résultats sur les mêmes 20 images de test'
    $rows=@(@('Métrique','Semi-supervisé','Supervisé'),@('Accuracy','0,900','0,950'),@('F1 cancer','0,889','0,947'),@('Précision cancer','1,000','1,000'),@('Rappel cancer','0,800','0,900'),@('F2 cancer','0,833','0,918'))
    [void](Table $s $rows 48 146 450 230 @(182,144,124) 18)
    Confusion $s $data.cnn_confusions.semi 'Faible puis fort' 533 142
    Confusion $s $data.cnn_confusions.super 'Fort uniquement' 734 142
    [void](Text $s '2 faux négatifs' 533 343 184 31 18 $true $teal)
    [void](Text $s '1 faux négatif' 734 343 184 31 18 $true $blue)
    [void](Text $s 'Le supervisé obtient le meilleur F1 cancer dans cette expérience.' 48 405 864 52 25 $true)
    [void](Text $s 'L''écart correspond à une image cancéreuse supplémentaire correctement reconnue.' 48 465 864 36 18 $false $gray)
    Note $s 'Source : tableau comparaison_etape4 et matrices de confusion enregistrés dans le notebook. Même test : 10 normal et 10 cancer. Semi [[10,0],[2,8]], supervisé [[10,0],[1,9]]. Accuracy 0.90 et 0.95, F1 cancer 0.889 et 0.947, précision 1 pour les deux, rappel 0.8 et 0.9, F2 0.833 et 0.918. Écart de F1 semi moins supervisé = -0.058 selon l''affichage arrondi du notebook. Ces résultats sont les sorties CNN enregistrées, sans nouvel entraînement lors de la création du support. L''utilisateur indique que la modification de standardisation/PCA ne change pas le résultat final. La conclusion suit son Markdown : le semi-supervisé n''apporte pas d''amélioration dans cette configuration et sur ce test. Fiche pages 2 et 3 : comparaison et cohérence des conclusions.'

    # 10. Conclusions and limitations
    $s=NewSlide 'Bilan et limites de l''expérience'
    [void](Text $s 'La chaîne technique est en place' 48 146 404 59 25 $true $teal)
    [void](Text $s 'Le notebook extrait les features, compare les clusterings et entraîne les deux approches CNN.' 48 223 404 88 22)
    [void](Text $s 'Dans cette configuration, le supervisé obtient les meilleures performances sur le test fort.' 48 350 404 98 24 $true)
    [void](Text $s 'Les résultats restent exploratoires' 516 146 396 59 25 $true)
    [void](Text $s "20 images de test et 16 de validation limitent la portée de la comparaison.`n`nL'ARI est proche de zéro. La correspondance des groupes 0/1 aux classes médicales n'est pas établie.`n`nCette correspondance reste à vérifier." 516 223 396 238 20)
    Foot $s 'Aucun seuil de performance ni validation clinique n''est établi.'
    Note $s 'Sources : conclusion de l''étape 4, Markdown sur les groupes K-Means, résultats ARI, énoncé et AGENTS.md. Critères observables couverts : features enregistrées et réutilisées, plusieurs méthodes de clustering, jeux forts/faibles séparés, comparaison CNN sur un test commun, métriques et sélection sur validation. Cela ne constitue pas une définition de réussite clinique. Aucun seuil métier n''est fixé dans les sources. Limites : test fort 20, validation forte 16, labels faibles issus de groupes arbitraires, ARI proche de zéro. Recommandation technique limitée à l''interprétation des résultats présents : vérifier la correspondance des groupes avant de présenter les labels faibles comme normal/cancer. La fiche d''auto-évaluation guide l''organisation, elle ne permet pas de déclarer tous les critères validés. Les quatre diapositives financières demandées pour plus tard ne figurent pas dans cette version.'

    if($deck.Slides.Count -ne 10){throw 'Incorrect slide count'}
    $deck.SaveAs($out,24)
    $checks=New-Object System.Collections.Generic.List[object]
    foreach($slide in $deck.Slides) {
        $slide.Export((Join-Path $renders ('slide-{0:00}.png' -f $slide.SlideIndex)),'PNG',1600,900)
        foreach($shape in $slide.Shapes) {
            if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $tr=$shape.TextFrame.TextRange
                $availableW=$shape.Width-$shape.TextFrame.MarginLeft-$shape.TextFrame.MarginRight
                $availableH=$shape.Height-$shape.TextFrame.MarginTop-$shape.TextFrame.MarginBottom
                $checks.Add([pscustomobject]@{slide=$slide.SlideIndex;name=$shape.Name;text=$tr.Text;boundWidth=$tr.BoundWidth;boundHeight=$tr.BoundHeight;availableWidth=$availableW;availableHeight=$availableH})
            }
        }
    }
    $checks | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $build 'text_fit.json') -Encoding UTF8
    [pscustomobject]@{path=$out;slides=$deck.Slides.Count;renderDirectory=$renders} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $build 'delivery.json') -Encoding UTF8
    Write-Output ('Created '+$out)
    Write-Output ('Rendered '+$deck.Slides.Count+' slides')
} finally {
    if($deck){$deck.Close()}
    if($openBefore -eq 0){$ppt.Quit()}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
