$ErrorActionPreference='Stop'
$workspace=(Get-Location).Path
$build=Join-Path $workspace '.build/brainscanai/cnn-aligned-ppt'
$source=Join-Path $workspace 'output/presentations/BrainScanAI_presentation_V1_Final.pptx'
$candidate=Join-Path $build 'candidate.pptx'
if(Test-Path -LiteralPath $candidate){throw 'Candidate already exists.'}
$hashes=Get-Content -LiteralPath (Join-Path $build 'source-hashes.json') -Raw -Encoding UTF8|ConvertFrom-Json
function CheckSources {
 foreach($entry in $hashes.PSObject.Properties){
  if((Get-FileHash -LiteralPath (Join-Path $workspace $entry.Name)).Hash.ToLowerInvariant() -ne $entry.Value){throw ('Source changed: '+$entry.Name)}
 }
}
function ShapeById($Slide,[int]$Id){
 foreach($shape in $Slide.Shapes){if($shape.Id -eq $Id){return $shape}}
 throw ('Missing shape '+$Id+' on slide '+$Slide.SlideIndex)
}
function Fmt($Value,[int]$Decimals){return ([double]$Value).ToString(('F'+$Decimals),[Globalization.CultureInfo]::GetCultureInfo('fr-FR'))}
CheckSources
$facts=Get-Content -LiteralPath (Join-Path $build 'facts.json') -Raw -Encoding UTF8|ConvertFrom-Json
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,-1,0,0)
try {
 $slide=$deck.Slides.Item(7)
 $table=(ShapeById $slide 6).Table
 for($r=0;$r -lt 4;$r++){
  $table.Cell($r+2,3).Shape.TextFrame.TextRange.Text=Fmt $facts.semi_grid[$r].'F2 cancer validation' 4
  $table.Cell($r+2,4).Shape.TextFrame.TextRange.Text=Fmt $facts.super_grid[$r].'F2 cancer validation' 4
 }
 $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text="Source : 01_exploration_dataset.ipynb, cellules hyperparametres-semi-recherche et hyperparametres-supervise-recherche, sorties enregistrées après alignement des cibles faibles et fortes. Phase faible : 5 et 10 époques donnent une accuracy de validation de 1,0000 par rapport aux groupes K-Means alignés. À égalité, le code retient 5 époques. F2 semi-supervisé : 0,4054 / 0,5263 / 0,5263 / 0,8537. F2 supervisé : 0,5000 / 0,6250 / 0,6579 / 0,9756. Les deux modèles retiennent un taux de 0,001 et 10 époques fortes. Le meilleur F2 de validation diffère entre modèles. Les 20 images du test final restent exclues de la sélection."

 $slide=$deck.Slides.Item(8)
 $accuracy=Fmt $facts.weak_test[0].Accuracy 3
 (ShapeById $slide 9).TextFrame.TextRange.Text="Le même CNN poursuit son entraînement : 0 = normal et 1 = cancer.`rTest faible : accuracy de $accuracy sur 263 images (accord avec K-Means)."
 $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text="Source : 01_exploration_dataset.ipynb, cellules ba2b3b8a, b184ba1a, 82faeff6 et 9dfb6484. La correspondance entre les groupes K-Means et les cibles du CNN est déterminée uniquement sur les 63 images fortes d'entraînement. Dans cette exécution, groupe 0 vers cible 1 (cancer), groupe 1 vers cible 0 (normal). Les numéros originaux restent dans label_faible. Les cibles du CNN sont alignées dans les phases faible et forte. Les labels faibles restent approximatifs. La validation faible applique la même correspondance. Après 5 époques faibles, accuracy de 0,973 sur les 263 images du test faible : accord avec les cibles dérivées de K-Means, et non évaluation médicale. Puis le même CNN poursuit son entraînement pendant 10 époques sur les 63 labels forts. Le supervisé apprend uniquement sur ces 63 images, pendant 10 époques. Taux de 0,001 pour les trois phases finales."

 $slide=$deck.Slides.Item(9)
 $table=(ShapeById $slide 5).Table
 $metrics=@('Accuracy','F1 cancer','Précision cancer','Rappel cancer','F2 cancer')
 for($r=0;$r -lt $metrics.Length;$r++){
  $metric=$metrics[$r]
  $table.Cell($r+2,2).Shape.TextFrame.TextRange.Text=Fmt $facts.semi_test[0].$metric 3
  $table.Cell($r+2,3).Shape.TextFrame.TextRange.Text=Fmt $facts.super_test[0].$metric 3
 }
 # The two confusion matrices in the saved notebook output are [[9,1],[1,9]].
 $matrix=(ShapeById $slide 8).Table
 $matrix.Cell(3,2).Shape.TextFrame.TextRange.Text='1'
 $matrix.Cell(3,3).Shape.TextFrame.TextRange.Text='9'
 (ShapeById $slide 12).TextFrame.TextRange.Text='1 faux négatif'
 (ShapeById $slide 14).TextFrame.TextRange.Text="Les deux CNN obtiennent les mêmes performances sur ce test.`rLa phase faible n'apporte pas d'amélioration mesurée."
 (ShapeById $slide 15).TextFrame.TextRange.Text='Écart de F2 train/test : +6,9 points (semi-supervisé), +4,3 points (supervisé).'
 $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text="Source : 01_exploration_dataset.ipynb, cellules ee07cebf et cnn-comparaison-train-test, sorties enregistrées après alignement des labels. Test fort : 20 images, 10 normales et 10 cancéreuses. Les deux modèles affichent accuracy, F1, précision, rappel et F2 cancer égaux à 0,900. Matrice de confusion pour chacun : [[9,1],[1,9]], lignes = classes réelles, colonnes = classes prédites. Un faux positif et un faux négatif par modèle. L'égalité concerne les métriques et les effectifs des matrices sur ce petit test, sans prouver que les modèles ou leurs prédictions individuelles sont identiques. Aucun gain mesuré du semi-supervisé sur ce test. Train fort, 63 images : F2 semi de 0,969 et supervisé de 0,943. Écarts de F2 train moins test : +6,9 et +4,3 points. Ces écarts suggèrent un possible surapprentissage, à interpréter avec prudence sur 20 images de test."

 if($deck.Slides.Count -ne 13){throw 'Slide count changed.'}
 $deck.SaveAs($candidate,24)
 $renderDir=Join-Path $build 'candidate-renders'
 [void][IO.Directory]::CreateDirectory($renderDir)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 CheckSources
 Write-Output 'Slides 7-9 updated in a private candidate; original presentation and notebook unchanged.'
} finally {$deck.Close()}
