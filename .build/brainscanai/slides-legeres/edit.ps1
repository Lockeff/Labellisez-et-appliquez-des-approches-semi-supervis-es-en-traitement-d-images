param([string]$Workspace=(Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/slides-legeres'
$source=Join-Path $Workspace 'output/presentations/BrainScanAI_presentation_13_slides_CPU_AWS.pptx'
$candidate=Join-Path $build 'candidate.pptx'
if(Test-Path -LiteralPath $candidate){throw 'Candidate already exists.'}
$notebookHash=(Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$helpers=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$start=$helpers.IndexOf('function RGB('); $end=$helpers.IndexOf('function Picture(')
Invoke-Expression $helpers.Substring($start,$end-$start)
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,0,0,0)
function Heading($Slide,$Title){
 for($i=$Slide.Shapes.Count;$i -ge 1;$i--){$Slide.Shapes.Item($i).Delete()}
 [void](Text $Slide $Title 48 30 864 80 32 $true)
 [void](Line $Slide 48 114 912 114)
 [void](Text $Slide ([string]$Slide.SlideIndex) 878 508 34 18 12 $false $gray 3)
}
try {
 $originals=Join-Path $build 'source-renders';[void][IO.Directory]::CreateDirectory($originals)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $originals ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 # Keep all existing detailed calculations and official source notes.
 $s=$deck.Slides.Item(10)
 $notes=$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s 'Machine AWS et besoins en RAM'
 [void](Text $s 'Mon PC de référence' 48 145 405 34 24 $true)
 [void](Text $s "i9-14900K`n24 cœurs physiques" 48 195 405 82 28 $true)
 [void](Text $s "RAM mesurée : environ 3,2 Gio`npour le processus Python" 48 304 405 62 22)
 [void](Text $s 'Machine AWS du devis' 506 145 406 34 24 $true)
 [void](Text $s "c7i.12xlarge, Paris`n24 cœurs physiques, 48 vCPU" 506 195 406 82 27 $true)
 [void](Text $s "96 Gio de RAM inclus`n2,28 € HT / heure, IPv4 comprise" 506 304 406 62 22)
 [void](Text $s 'La taille du CPU impose les 96 Gio. Le notebook utilise bien moins de RAM.' 48 405 864 57 23 $true $teal)
 [void](Text $s 'Même nombre de cœurs. Débit à comparer avec le même code sur AWS.' 48 480 805 31 18 $false $gray)
 Note $s ($notes+"`n`nMémoire : les 3,2 Gio représentent le maximum observé pour le processus Python, pas toute la mémoire de la machine et de son système d'exploitation. Les 96 Gio sont inclus dans le prix EC2 et liés à la configuration CPU. Le traitement progressif par lots doit conserver une consommation bornée. Le pic réel sur AWS doit être mesuré. Affichage du tarif arrondi à 2,28 EUR/h. Tous les calculs conservent le tarif exact 2,279814053280887 EUR/h.")

 $s=$deck.Slides.Item(11)
 $notes=$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s 'Coût du prototype'
 [void](Text $s '30 heures de VM, une seule fois' 48 145 864 38 26)
 [void](Text $s '72,55 € HT' 48 219 410 70 44 $true $teal)
 [void](Text $s 'Coût cloud estimé' 48 301 410 34 23)
 [void](Text $s '68,39 €  Machine et IPv4' 506 230 406 40 24 $true)
 [void](Text $s '4,16 €  Disque, images et accès' 506 294 406 70 24)
 [void](Text $s 'Essais et entraînements inclus dans les 30 heures.' 48 385 864 40 24)
 [void](Text $s 'Budget IA : 300 €. Temps de travail humain exclu.' 48 442 864 40 23 $true)
 Foot $s 'Estimation AWS Paris du 09/10/2026. Disque et images conservés un mois. Calculs et sources en notes.'
 Note $s ($notes+"`n`nPrésentation synthétique : 4,16 EUR = disque 4,15 EUR + images et accès 0,01 EUR. Total des lignes affichées : 68,39 + 4,16 = 72,55 EUR. Les essais et entraînements sont inclus dans les 30 heures de VM active. Le salaire correspondant aux 30 heures de travail humain reste exclu. La durée de conservation du disque et des images est d'un mois.")

 $s=$deck.Slides.Item(12)
 $notes=$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s "Coût pour 4 millions d'images"
 [void](Text $s 'Estimation sur la première année, avec le modèle déjà entraîné' 48 141 864 38 23)
 [void](Text $s 'En une fois' 48 203 405 35 26 $true)
 [void](Text $s '67,03 € HT' 48 253 405 62 40 $true $teal)
 [void](Text $s "4 millions d'images au mois 1`n8,2 heures de calcul au total" 48 334 405 62 22)
 [void](Text $s 'Chaque mois, sur un an' 506 203 406 35 26 $true)
 [void](Text $s '101,78 € HT' 506 253 406 62 40 $true $teal)
 [void](Text $s "Environ 333 333 images / mois`n41 minutes de calcul / mois" 506 334 406 62 22)
 [void](Text $s 'Calcul, stockage et accès inclus. Images : environ 99 Go, conservées 12 mois chacune.' 48 417 864 48 20)
 [void](Text $s 'Hypothèse : AWS atteint le débit du i9. VM arrêtée entre les lots.' 48 476 864 30 18 $true $gray)
 Foot $s 'Hors prototype et contrôle humain. Coûts HT. Calculs détaillés et sources en notes.'
 Note $s ($notes+"`n`nLes deux montants sont des totaux sur le premier exercice de 12 mois, pas des tarifs par heure ou par lot. La RAM est incluse dans le tarif de la même c7i.12xlarge. En une fois : 18,69 calcul + 23,74 stockage images + 4,15 disque + 20,45 accès = 67,03 EUR. Glissant année 1 : 18,69 + 12,86 + 49,78 + 20,45 = 101,78 EUR. Le coût mensuel n'est pas constant : le stock d'images augmente pendant la première année. Le disque est conservé 1 mois en une fois, 12 mois en glissant. Démarrer avec stockage vide et supprimer chaque image après 12 mois. La durée de calcul reste conditionnelle au test du débit réel AWS, et les temps de préparation doivent être ajoutés s'ils sont facturés.")

 $s=$deck.Slides.Item(13)
 $notes=$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s "Conditions du passage à l'échelle"
 [void](Text $s 'Budget de 5 000 € : compatible avec le scénario estimé' 48 142 864 59 26 $true $teal)
 [void](Text $s 'Débit : chronométrer le même traitement sur AWS.' 48 225 864 43 25)
 [void](Text $s 'RAM : lire les images par lots de 16.' 48 285 864 43 25)
 [void](Text $s 'Coûts : arrêter la VM entre les lots.' 48 345 864 43 25)
 [void](Text $s 'Qualité : contrôler les labels sur de nouvelles images.' 48 405 864 43 25)
 [void](Text $s 'Les 20 images de test ne suffisent pas à valider 4 millions de labels.' 48 478 805 31 18 $false $gray)
 Note $s $notes

 if($deck.Slides.Count -ne 13){throw 'Slide count changed.'}
 $deck.SaveAs($candidate,24)
 $renders=Join-Path $build 'candidate-renders';[void][IO.Directory]::CreateDirectory($renders)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renders ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Source changed.'}
 if((Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash -ne $notebookHash){throw 'Notebook changed.'}
 [pscustomobject]@{Source=$source;SourceHash=$sourceHash;NotebookHash=$notebookHash;Slides=13}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
} finally {$deck.Close()}
