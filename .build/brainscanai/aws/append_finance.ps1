param([string]$Workspace = (Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/aws'
$final=Join-Path $Workspace 'output/presentations/BrainScanAI_presentation_13_slides_AWS.pptx'
$snapshot=Join-Path $build 'source-live.pptx'
$candidate=Join-Path $build 'finance-candidate.pptx'
foreach($path in @($final,$snapshot,$candidate)){ if(Test-Path -LiteralPath $path){throw "Existing output: $path"} }
$notebook=Join-Path $Workspace '01_exploration_dataset.ipynb'
$beforeHash=(Get-FileHash -LiteralPath $notebook -Algorithm SHA256).Hash
$helpers=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$start=$helpers.IndexOf('function RGB(')
$end=$helpers.IndexOf('function Picture(')
Invoke-Expression $helpers.Substring($start,$end-$start)
$ppt=New-Object -ComObject PowerPoint.Application
$original=$null
foreach($presentation in $ppt.Presentations){
    if($presentation.Name -eq 'BrainScanAI_presentation_10_slides.pptx'){$original=$presentation;break}
}
if($null -eq $original){throw 'The user presentation is not open. Do not replace it with a disk version.'}
if($original.Slides.Count -ne 10){throw 'Source slide count changed; review before continuing.'}
$original.SaveCopyAs($snapshot,24)
$deck=$ppt.Presentations.Open($snapshot,0,0,0)
function Panel($Slide,$Value,$X,$Y,$W,$H,$Size=20){
    $shape=Node $Slide $Value $X $Y $W $H $teal $Size
    $shape.TextFrame2.AutoSize=0
    $shape.Height=[single]$H;$shape.Width=[single]$W
    return $shape
}
try {
    $s=NewSlide 'AWS : machine et budget du prototype'
    [void](Text $s 'Paris · Linux · paiement à l''usage' 48 132 545 30 22 $true)
    $rows=@(@('Machine','CPU','RAM','Prix / h'),@('c6i.large','2','4 Gio','0,101 $'),@('c6i.xlarge','4','8 Gio','0,202 $'))
    [void](Table $s $rows 48 179 545 117 @(220,68,105,152) 19)
    [void](Text $s 'Point de départ proposé : c6i.xlarge. Le minimum utilisable et la durée doivent être vérifiés sur AWS.' 48 314 545 78 21)
    [void](Text $s 'Cette VM a moins de ressources CPU que le PC local i9-14900K : les 7,5 h locales ne sont pas une mesure AWS.' 48 407 545 69 18)
    [void](Panel $s 'Budget actuel : 300 €' 638 133 274 50 23)
    [void](Text $s "Simulation pour le dataset actuel`n10 h de création et d'essais : 2,02 $`n50 Gio de disque, 1 mois : 4,64 $`nS3, requêtes et IPv4 : 0,06 $" 638 205 274 132 17)
    [void](Text $s 'Total : 6,72 $ HT' 638 354 274 60 24 $true $teal)
    [void](Text $s 'Le mail ne précise pas si les essais sont compris : hypothèse retenue ici, à confirmer avec Clara.' 638 423 274 69 17)
    Foot $s 'Tarifs AWS Paris au 09/10/2026, USD HT. Les 10 h sont une provision, pas un temps mesuré.'
    Note $s @'
Sources officielles consultées le 9 octobre 2026 :
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv
Catalogue EC2 publié le 8 octobre 2026, tarifs effectifs au 1er octobre 2026 : Linux, tenancy Shared, OnDemand, eu-west-3, RunInstances, CapacityStatus Used. c6i.large : SKU B8Y99WMXMHSS82BK, 2 vCPU, 4 GiB, 0.101 USD/h. c6i.xlarge : SKU CKHYJHTXURYY5WKR, 4 vCPU, 8 GiB, 0.202 USD/h. EBS gp3 : SKU XCKADFE99D2BK5V6, 0.0928 USD/GiB-mois.
https://aws.amazon.com/ec2/pricing/on-demand/
https://docs.aws.amazon.com/ec2/latest/instancetypes/co.html
https://aws.amazon.com/ebs/pricing/
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonS3/current/eu-west-3/index.json
Catalogue S3 publié le 28 septembre 2026. S3 Standard : 0.024 USD/GiB-mois. PUT : 0.0053 USD/1000. GET : 0.00042 USD/1000. Ici : 1506 fichiers sources, 37272490 octets, un envoi et une lecture par fichier, un mois de conservation. Coût S3 et requêtes = 0.009447425 USD.
https://aws.amazon.com/vpc/pricing/ : IPv4 publique 0.005 USD/h, allouée seulement pendant 10 h = 0.05 USD, aucune IP conservée à l'arrêt.
Calcul illustratif : 10*0.202 + 50*0.0928 + 0.05 + 0.009447425 = 6.719447425 USD HT. La provision de 10 h couvre création, extraction des embeddings, recherche d'hyperparamètres, entraînement et essais d'inférence ; elle n'est pas une mesure cloud. Une fois le temps mesuré, remplacer 10 h par la durée réelle. Installation et essais consomment aussi du temps de VM.
Sources projet : notebook actuel (CPU), présentation ouverte de l'utilisateur (PC Intel i9-14900K), énoncé de Clara. Le coût ne comprend pas le travail humain ni une annotation humaine. Les crédits gratuits ne sont pas déduits. TVA et conversion EUR/USD non incluses ; les budgets de mission sont en euros. Le mail ne permet pas d'établir si essais et stockage sont inclus : les inclure est une hypothèse de chiffrage à confirmer.
'@

    $s=NewSlide '4 millions d''images : coût sur un an'
    [void](Text $s 'Hypothèses : JPEG de même taille moyenne (24,75 ko), soit ≈ 99 Go ; une copie conservée jusqu''au mois 12.' 48 133 864 47 18)
    $rows=@(
        @('Poste · USD HT','4 M d''un coup','400 000 / mois'),
        @('Création et essais · provision 10 h','2,02','2,02'),
        @('Prédictions · provision 30 h au total','6,06','6,06'),
        @('4 M envois + 4 M lectures S3','22,88','22,88'),
        @('Images S3 · conservation sur un an','26,55','16,60'),
        @('Disque VM 50 Gio + IPv4','55,88','55,88'),
        @('TOTAL ILLUSTRATIF','113,39 $','103,44 $')
    )
    [void](Table $s $rows 48 193 864 245 @(402,231,231) 18)
    [void](Text $s '400 000 / mois pendant 10 mois = 4 M. Éteindre la VM entre les lots ; sur 12 mois, ce rythme donnerait 4,8 M.' 48 451 864 40 17)
    Foot $s 'À 1 $ = 1 € (hypothèse, pas un taux vérifié) : ≈ 113 € / 103 € HT. Budget 5 000 € plausible sous ces conditions.'
    Note $s @'
Sources : tarifs AWS officiels Paris du 09/10/2026 cités dans les notes de la diapositive 11, mêmes instances et références de catalogues.
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonS3/current/eu-west-3/index.json
https://aws.amazon.com/s3/pricing/
https://aws.amazon.com/ebs/pricing/
https://aws.amazon.com/vpc/pricing/
Taille réellement observée : 1506 JPEG du dataset source, 37272490 octets, moyenne 24749.32934926959 octets. Hypothèse : mêmes tailles pour les nouvelles images. 4 M = 98997317397.07837 octets = 92.19843651827271 GiB. AWS S3 facture les GB selon la base binaire ; 0.024 USD par unité de 2^30 octets et par mois. Cette estimation serait à refaire pour des PNG ou des images plus grandes.
Scénario 1 : les 4 M images sont disponibles au début du mois 1 et conservées pendant 12 mois. S3 = 92.19843651827271 * 0.024 * 12 = 26.553149717262542 USD.
Scénario 2 : 400000 images arrivent au début de chacun des mois 1 à 10 ; elles sont conservées jusqu'à la fin du mois 12. Stock occupé = 0.1+0.2+...+1.0+1.0+1.0 = 7.5 fois le volume complet sur un mois. S3 = 92.19843651827271 * 0.024 * 7.5 = 16.59571857328909 USD. Il s'agit de 4 M sur une fenêtre de 12 mois, avec dix mois d'arrivées ; douze mois à 400000 donnent 4.8 M.
Calcul commun : un entraînement/ensemble d'essais initial, provision 10 h * 0.202 = 2.02 USD ; une seule prédiction par image, provision illustrative 30 h * 0.202 = 6.06 USD (équivalent à quatre fois l'estimation locale de 7.5 h, hypothèse de sensibilité et NON une vitesse mesurée ou garantie AWS). Les dix lots mensuels consomment au total ces mêmes 30 h, soit provision 3 h par lot. Pas de réentraînement mensuel supposé.
Un objet S3 par JPEG, un PUT et un GET par image : 4 M * (0.0000053 + 0.00000042) = 22.88 USD. Un CSV consolidé pour les labels. S3 et EC2 sont dans la même région ; entrée des données et transfert S3 vers EC2 de même région gratuits, requêtes payantes. Pas de rapatriement des 4 M images vers Internet ; télécharger les labels seulement. Pas de NAT Gateway, de versions supplémentaires ou de réplication. Les temps de transfert, d'installation et d'attente doivent être mesurés et sont à ajouter s'ils dépassent les provisions.
Disque gp3 de travail 50 GiB conservé 12 mois, traitement progressif des images : 50 * 0.0928 * 12 = 55.68 USD. Le disque de 50 GiB ne contient pas toutes les 4 M images à la fois. IPv4 éphémère pendant 40 h : 0.20 USD, aucune IP allouée au repos. La VM est arrêtée entre les traitements. Une c6i.xlarge laissée active 365 jours coûte 8760 * 0.202 = 1769.52 USD de calcul seul ; EBS reste facturé après arrêt tant que le volume existe.
Totaux non arrondis : 113.39314971726253 USD et 103.43571857328908 USD. L'hypothèse de conversion 1 USD = 1 EUR sert uniquement à situer les ordres de grandeur par rapport aux budgets en euros ; ce n'est pas le taux de change du jour. Taxes, travail humain, annotation humaine et support payant ne sont pas inclus. La provision de création est comprise dans ce scénario de 5000 EUR ; ne pas additionner automatiquement le budget du prototype de 300 EUR. Le périmètre précis de ces budgets reste à confirmer avec Clara.
Conclusion conditionnelle : la facture d'infrastructure peut être compatible avec 5000 EUR pour ces volumes et ces durées. Aucune garantie de coût réel, de vitesse AWS ou de qualité médicale n'en découle. Mesurer la chaîne après les changements demandés par le mentor avant de valider le chiffrage.
'@

    $s=NewSlide '7,5 heures : conditions et points bloquants'
    [void](Text $s 'Mesure locale : 1 000 images en 6,8 s → estimation de 7,5 h pour 4 M. Lecture, prétraitement et prédiction inclus.' 48 134 864 48 20 $true)
    [void](Panel $s 'Lire les images' 48 202 193 54 19)
    [void](Panel $s 'Prétraiter' 272 202 193 54 19)
    [void](Panel $s 'Prédire · CNN' 496 202 193 54 19)
    [void](Panel $s 'Labels · CSV' 720 202 193 54 19)
    [void](Arrow $s 247 229 266 229)
    [void](Arrow $s 471 229 490 229)
    [void](Arrow $s 695 229 714 229)
    [void](Text $s 'Lecture et CPU' 48 277 409 27 22 $true $teal)
    [void](Text $s 'Petits fichiers S3, décodage JPEG et prétraitement peuvent freiner la chaîne. Charger les lots en parallèle, puis mesurer le gain.' 48 316 409 85 19)
    [void](Text $s 'RAM : traiter par lots. Charger 4 M images normalisées à la fois demanderait environ 2,2 Tio, hors autres données.' 48 417 409 69 18)
    [void](Text $s 'GPU : une option à tester' 503 277 409 27 22 $true $teal)
    [void](Text $s 'g4dn.xlarge · NVIDIA T4 : 0,615 $/h. Il faut environ 3,1 fois plus de débit que la VM CPU à 0,202 $/h pour réduire le coût de calcul.' 503 316 409 86 19)
    [void](Text $s 'Pilote AWS : mesurer durée et pic RAM séparément pour la création du modèle et le traitement final, après les ajouts du mentor.' 503 417 409 72 18 $true)
    Foot $s 'Le temps cloud et les pics RAM restent à mesurer. Les 7,5 h excluent l''entraînement et le transfert initial vers AWS.'
    Note $s @'
Source expérimentale : notebook 01_exploration_dataset.ipynb, cellule finale de benchmark a3f5cc5d, résultat sauvegardé : 1000 images en 6.8 secondes sur cpu, estimation 4 M en 7.5 heures. time.perf_counter couvre read_image en RGB, prétraitement ImageNet, création des lots de 16, copie vers le périphérique et inférence du modèle supervisé avec no_grad/eval. Aucun benchmark cloud ni GPU n'est présent. La diapositive 10 de la présentation ouverte de l'utilisateur indique 6.7 s et 7.5 h ; elle est conservée sans modification. La valeur affichée ici est celle de la sortie sauvegardée du notebook. Une mesure sur 1000 fichiers en local peut bénéficier du cache ; l'extrapolation ne prouve pas le temps sur 4 M objets AWS.
RAM calculée, non mesurée : chaque image normalisée RGB en float32 de 224*224 représente 3*224*224*4 = 602112 octets. 4 M images = 2408448000000 octets = 2.1904706955 TiB, sans compter les listes, embeddings, modèles et buffers. L'étape d'extraction actuelle prépare une liste de toutes les images ; elle doit être adaptée à un chargement progressif pour ce volume. L'inférence finale du notebook traite déjà par lots de 16.
Sources GPU et prix :
https://aws.amazon.com/ec2/instance-types/g4/ : G4dn utilise NVIDIA T4.
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv : g4dn.xlarge Linux Shared OnDemand Paris, SKU H4KD74Z7H4KRNZMZ, 4 vCPU, 16 GiB, 1 GPU, 0.615 USD/h. c6i.xlarge : 0.202 USD/h. Seuil de rentabilité du calcul seul : 0.615/0.202 = 3.04455 fois le débit ; la recommandation arrondit à 3.1. Ne pas appliquer aux coûts fixes de stockage. Le GPU n'accélère pas nécessairement une lecture S3 ou un prétraitement séquentiel. Aucun gain mesuré n'est annoncé.
Les blocages sont des hypothèses à tester, pas des résultats de profilage. Égalisation des histogrammes, essai d'une représentation intermédiaire et contrôle du surapprentissage sont les demandes du mentor ; la présentation n'affirme pas qu'ils sont déjà implémentés. Chronométrer création (embeddings, recherche, entraînement) et inférence séparément, en incluant les temps de lecture et de préparation ; mesurer aussi les pics RAM. Le prétraitement des nouvelles images doit correspondre à celui du modèle final. Après tout changement de couche ou de prétraitement, refaire la mesure et vérifier la qualité sur le test réservé. La faisabilité financière ne constitue pas une validation clinique.
'@
    if($deck.Slides.Count -ne 13){throw 'Expected thirteen slides.'}
    $deck.SaveAs($candidate,24)
    $renderDir=Join-Path $build 'final-renders'
    [void][IO.Directory]::CreateDirectory($renderDir)
    for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
    $sourceRenderDir=Join-Path $build 'snapshot-renders'
    [void][IO.Directory]::CreateDirectory($sourceRenderDir)
    $sourceDeck=$ppt.Presentations.Open($snapshot,-1,0,0)
    try {for($i=1;$i -le 10;$i++){$sourceDeck.Slides.Item($i).Export((Join-Path $sourceRenderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}} finally {$sourceDeck.Close()}
    $afterHash=(Get-FileHash -LiteralPath $notebook -Algorithm SHA256).Hash
    if($beforeHash -ne $afterHash){throw 'Notebook changed during presentation work; inspect before delivery.'}
    [pscustomobject]@{Candidate=$candidate;Slides=$deck.Slides.Count;NotebookSHA256=$afterHash;OriginalSaved=$original.Saved;SourcesPreserved=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
} finally {if($null -ne $deck){$deck.Close()}}
