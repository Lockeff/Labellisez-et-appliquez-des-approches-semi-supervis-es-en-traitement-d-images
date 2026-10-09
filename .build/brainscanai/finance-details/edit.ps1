$ErrorActionPreference='Stop'
$workspace=(Get-Location).Path
$build=Join-Path $workspace '.build/brainscanai/finance-details'
$source=Join-Path $workspace 'output/presentations/BrainScanAI_presentation_V1.pptx'
$candidate=Join-Path $build 'candidate.pptx'
if(Test-Path -LiteralPath $candidate){throw 'Candidate exists.'}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$notebookHash=(Get-FileHash -LiteralPath (Join-Path $workspace '01_exploration_dataset.ipynb')).Hash
$helpers=Get-Content -LiteralPath (Join-Path $workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$start=$helpers.IndexOf('function RGB(');$end=$helpers.IndexOf('function Picture(')
Invoke-Expression $helpers.Substring($start,$end-$start)
$pricingUrl='https://docs.aws.amazon.com/awsaccountbilling/latest/aboutv2/using-the-aws-price-list-bulk-api-fetching-price-list-files-manually.html'
function Heading($Slide,$Title){
 for($i=$Slide.Shapes.Count;$i -ge 1;$i--){$Slide.Shapes.Item($i).Delete()}
 [void](Text $Slide $Title 48 30 864 80 32 $true)
 [void](Line $Slide 48 114 912 114)
 [void](Text $Slide ([string]$Slide.SlideIndex) 878 508 34 18 12 $false $gray 3)
}
function PricingLink($Slide){
 $link=Text $Slide 'Source : AWS Price List, tarifs officiels' 48 504 640 25 14 $false $teal
 $link.TextFrame.TextRange.Font.Underline=-1
 $link.ActionSettings.Item(1).Action=7
 $link.ActionSettings.Item(1).Hyperlink.Address=$pricingUrl
}
function ConsolidateNotes([string]$Original){
 # One AWS reference for the financial slides. Keep the existing calculation
 # facts and the dated non-AWS currency conversion attribution.
 $text=[regex]::Replace($Original,'https?://\S+',{
  param($Match)
  if($Match.Value -match '(aws\.amazon\.com|amazonaws\.com)'){return ''}
  return $Match.Value
 })
 return $text.Trim()
}
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,0,0,0)
try {
 $renders=Join-Path $build 'source-renders';[void][IO.Directory]::CreateDirectory($renders)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renders ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 $s=$deck.Slides.Item(11)
 $notes=ConsolidateNotes $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s 'Coût du prototype'
 [void](Text $s '30 heures de VM, essais et entraînements inclus' 48 145 864 40 25)
 [void](Text $s '72,55 € HT' 48 215 375 65 43 $true $teal)
 [void](Text $s 'Une seule fois' 48 300 375 35 24)
 [void](Text $s 'Machine + IPv4' 480 208 432 34 24 $true)
 [void](Text $s '30 h × 2,2798 €/h = 68,39 €' 480 252 432 34 23)
 [void](Text $s 'Disque : 1 mois × 4,148 € = 4,15 €' 480 301 432 34 21)
 [void](Text $s 'Images et accès : 0,01 €' 480 348 432 34 21)
 [void](Text $s '68,39 € + 4,15 € + 0,01 € = 72,55 €' 48 407 864 38 25 $true)
 [void](Text $s 'Paris, prix HT du 09/10/2026. Conversion : 1 € = 1,1186 $. Temps humain exclu.' 48 463 864 29 16 $false $gray)
 PricingLink $s
 Note $s ($notes+"`n`nDÉTAIL DU CALCUL DEMANDÉ`nMachine c7i.12xlarge Linux On-Demand à Paris, 96 Gio de RAM incluse, adresse IPv4 publique comprise. Tarif EC2 : 2,5452 USD/h. IPv4 : 0,005 USD/h. Conversion datée du 08/10/2026 : 1 EUR = 1,1186 USD. Tarif exact (2,5452 + 0,005) / 1,1186 = 2,279814053280887 EUR/h, affiché à 2,2798 EUR/h pour expliquer les multiplications. 30 h × tarif exact = 68,3944215984266 EUR, soit 68,39 EUR. Disque EBS gp3 de 50 Gio, 0,0928 USD/Gio-mois : 1 mois × 50 × 0,0928 / 1,1186 = 4,148042195601645 EUR, soit 4,15 EUR. Stockage des 1506 fichiers du dataset (37272490 octets, un mois), une PUT et une GET par fichier : 0,0084457581551753 EUR, soit 0,01 EUR. Somme des montants affichés : 68,39 + 4,15 + 0,01 = 72,55 EUR. Le tarif arrondi à 2,28 EUR/h de la slide 10 est destiné à la lecture, les calculs utilisent le tarif exact. Aucun coût humain inclus.`n`nSOURCE AWS UNIQUE POUR LES TARIFS`n"+$pricingUrl+"`nAWS Price List donne accès aux catalogues par service, version et région. Choisir eu-west-3 (Paris), les tarifs en vigueur le 09/10/2026, Linux, On-Demand, Shared. EC2 : c7i.12xlarge, SKU RKFRM5JW99YQMR8K, prix effectif au 01/10/2026. EBS gp3 figure dans AmazonEC2, IPv4 dans AmazonVPC, stockage et requêtes dans AmazonS3. Les totaux en euros sont nos calculs à partir des tarifs unitaires AWS, et ne sont pas des montants publiés tels quels par AWS.")

 $s=$deck.Slides.Item(12)
 $notes=ConsolidateNotes $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
 Heading $s "Coût pour 4 millions d'images"
 [void](Text $s "Base de ces montants : 8,2 h de calcul au total, sur la première année" 48 141 864 32 22)
 [void](Text $s 'En une fois' 48 187 415 33 25 $true)
 [void](Text $s '67,03 € HT' 48 230 415 53 35 $true $teal)
 [void](Text $s '4 millions au mois 1' 48 291 415 29 20)
 [void](Text $s 'Chaque mois, sur un an' 505 187 407 33 25 $true)
 [void](Text $s '101,78 € HT' 505 230 407 53 35 $true $teal)
 [void](Text $s 'Environ 333 333 images / mois' 505 291 407 29 20)
 [void](Text $s 'Calcul : 8,2 h × 2,2798 €/h = 18,69 €' 48 334 415 28 18)
 [void](Text $s 'Images : 12 mois × 1,97815 € = 23,74 €' 48 366 415 28 18)
 [void](Text $s 'Disque : 1 mois × 4,148 € = 4,15 €' 48 398 415 28 18)
 [void](Text $s 'Accès : 20,45 €' 48 430 415 28 18)
 [void](Text $s 'Calcul : 8,2 h × 2,2798 €/h = 18,69 €' 505 334 407 28 18)
 [void](Text $s 'Stockage progressif des images : 12,86 €' 505 366 407 28 18)
 [void](Text $s 'Disque : 12 mois × 4,148 € = 49,78 €' 505 398 407 28 18)
 [void](Text $s 'Accès : 20,45 €' 505 430 407 28 18)
 [void](Text $s 'Dernière mesure : 9 h. Ces coûts utilisent la mesure précédente de 8,2 h.' 48 471 864 26 17 $true $gray)
 PricingLink $s
 Note $s ($notes+"`n`nDÉTAIL DES DEUX TOTAUX DEMANDÉS`nCes montants sont conservés pour expliquer les estimations de 67,03 et 101,78 EUR. Ils reposent sur l'ancienne mesure locale extrapolée à 8,2 h, soit 41 minutes par lot mensuel. La dernière mesure du notebook, également reportée sur la slide 10, est d'environ 9 h : 45 minutes par lot, sous la même hypothèse de débit. Les deux totaux doivent être recalculés si l'on adopte cette nouvelle durée. Le débit réel sur AWS reste à mesurer. 96 Gio de RAM inclus dans le prix EC2. VM arrêtée entre lots et adresse IPv4 libérée.`n`nTarif exact machine + IP : (2,5452 + 0,005) / 1,1186 = 2,279814053280887 EUR/h. Prix EBS : 50 Gio × 0,0928 USD/Gio-mois / 1,1186 = 4,148042195601645 EUR/mois. Taille moyenne du JPEG observée : 37272490 / 1506 = 24749,329349 octets. 4 M fichiers représentent 92,1984365 Gio (environ 99 Go). S3 Standard : 0,024 USD/Gio-mois, soit 1,978153474377387 EUR/mois au stock complet. Prix PUT : 0,0053 USD/1000 requêtes. GET : 0,00042 USD/1000 requêtes. Un envoi et une lecture par image = 4000000 / 1000 × (0,0053 + 0,00042) / 1,1186 = 20,45413910244949 EUR, soit 20,45 EUR.`n`nEN UNE FOIS`nMachine : 8,2 h × 2,279814053280887 = 18,6944752369 EUR, soit 18,69 EUR. Images conservées 12 mois : 12 × 1,978153474377387 = 23,7378416925 EUR, soit 23,74 EUR. Disque conservé un mois : 4,15 EUR. Accès : 20,45 EUR. Total affiché : 18,69 + 23,74 + 4,15 + 20,45 = 67,03 EUR HT.`n`nGLISSANT, PREMIÈRE ANNÉE`nMême temps total de calcul de 8,2 h réparti en 12 lots mensuels. Machine : 18,69 EUR. Le stock démarre vide. Les lots arrivent au début de chaque mois et chaque image reste 12 mois. Le stock représente successivement 1/12, 2/12, ... 12/12 du volume complet. Somme : (1 + 2 + ... + 12) / 12 = 6,5 mois de stock complet. Stockage : 6,5 × 1,978153474377387 = 12,85799758345 EUR, soit 12,86 EUR. Disque conservé 12 mois : 12 × 4,148042195601645 = 49,7765063472 EUR, soit 49,78 EUR. Accès : 20,45 EUR. Total affiché : 18,69 + 12,86 + 49,78 + 20,45 = 101,78 EUR HT. Les images stockées et les coûts mensuels augmentent pendant cette première année. Les coûts suivants, après stabilisation du stock, sont décrits dans les notes originales.`n`nImages conservées 12 mois chacune. Une seule copie, EC2/S3 dans la même région, aucun transfert intégral vers Internet, aucun NAT Gateway. Hors prototype, annotation/contrôle humain, réentraînements et taxes. Les arrondis sont effectués à la fin des calculs de chaque poste.`n`nSOURCE AWS UNIQUE POUR LES TARIFS`n"+$pricingUrl+"`nAWS Price List : catalogues AmazonEC2 (machine et EBS), AmazonVPC (IPv4) et AmazonS3 (stockage et requêtes), région eu-west-3, tarifs consultés le 09/10/2026. La page explique comment retrouver un catalogue par service, région et version historique. Les totaux en euros proviennent des multiplications de nos quantités par ces tarifs AWS et du taux de conversion conservé dans le support.")

 if($deck.Slides.Count -ne 13){throw 'Slide count changed.'}
 $deck.SaveAs($candidate,24)
 $renders=Join-Path $build 'candidate-renders';[void][IO.Directory]::CreateDirectory($renders)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renders ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Source modified.'}
 if((Get-FileHash -LiteralPath (Join-Path $workspace '01_exploration_dataset.ipynb')).Hash -ne $notebookHash){throw 'Notebook modified.'}
 [pscustomobject]@{SourceHash=$sourceHash;NotebookHash=$notebookHash;Source=$source;Candidate=$candidate;Slides=13}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
} finally {$deck.Close()}
