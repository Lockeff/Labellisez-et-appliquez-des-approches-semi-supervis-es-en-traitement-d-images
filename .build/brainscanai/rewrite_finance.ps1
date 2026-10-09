param([string]$Workspace=(Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/aws-clear'
[void][IO.Directory]::CreateDirectory($build)
$sourcePath=Join-Path $build 'source-live.pptx'
$candidate=Join-Path $build 'candidate.pptx'
foreach($p in @($sourcePath,$candidate)){if(Test-Path -LiteralPath $p){throw "File exists: $p"}}
$notebookHash=(Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash
$helperSource=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$helperStart=$helperSource.IndexOf('function RGB(')
$helperEnd=$helperSource.IndexOf('function Picture(')
Invoke-Expression $helperSource.Substring($helperStart,$helperEnd-$helperStart)
$ppt=New-Object -ComObject PowerPoint.Application
$original=$null
foreach($p in $ppt.Presentations){if($p.Name -eq 'BrainScanAI_presentation_13_slides_AWS.pptx'){$original=$p;break}}
if($null -eq $original){throw 'Expected source presentation must be open.'}
if($original.Slides.Count -ne 13){throw 'Source slide count changed.'}
$original.SaveCopyAs($sourcePath,24)
$deck=$ppt.Presentations.Open($sourcePath,0,0,0)
function ClearContent($Slide){
    for($i=$Slide.Shapes.Count;$i -ge 1;$i--){$Slide.Shapes.Item($i).Delete()}
    [void](Text $Slide ([string]$Slide.SlideIndex) 878 508 34 18 12 $false $gray 3)
}
function Heading($Slide,$Title){
    ClearContent $Slide
    [void](Text $Slide $Title 48 30 866 80 32 $true)
    [void](Line $Slide 48 114 912 114)
}
try {
    $s=$deck.Slides.Item(11)
    Heading $s 'Prototypage : 30 heures de travail'
    [void](Text $s 'Hypothèse : VM AWS Paris allumée pendant les 30 h. Machine c6i.xlarge : 4 CPU virtuels, 8 Gio de RAM.' 48 136 864 53 21)
    $rows=@(
        @('Poste AWS','Tarif et unité','Quantité / fréquence','Coût HT'),
        @('Machine + adresse IPv4','0,1851 € / heure','30 h, une seule fois','5,55 €'),
        @('Disque de travail, 50 Gio','4,148 € / mois','1 mois de conservation','4,15 €'),
        @('Images et accès au stockage','≈ 0,01 € / dataset','37 Mo ; 1 envoi + 1 lecture/image','0,01 €'),
        @('TOTAL CLOUD','','Pour ce prototype','9,71 €')
    )
    $table=Table $s $rows 48 213 864 215 @(253,191,286,134) 18
    for($c=1;$c -le 4;$c++){$table.Table.Cell(5,$c).Shape.TextFrame.TextRange.Font.Bold=-1}
    [void](Text $s 'Les essais et les entraînements sont inclus dans les 30 h de VM. Le temps humain reste à chiffrer : 30 h × tarif horaire.' 48 447 864 48 20 $true)
    Foot $s 'EUR HT : prix AWS Paris au 09/10/2026, convertis au taux BCE du 08/10/2026 (1 € = 1,1186 $).'
    Note $s @'
Scénario de l'utilisateur : environ 30 heures de travail pour le prototypage. Hypothèse explicite : la VM est allumée pendant ces 30 heures, même lors des phases de rédaction ou d'analyse. Les heures AWS sont des heures de VM allumée, pas des heures de CPU à pleine charge. Les essais, installation, extraction, recherches d'hyperparamètres et entraînements consomment cette enveloppe. Si la VM est arrêtée pendant une partie du travail, le coût de calcul diminue. L'exécution locale complète mesurée après égalisation dure 141.4115 secondes de cellules de calcul, mais ce temps ne représente pas le temps de travail humain ni le temps cloud.
La mission annonce 300 EUR pour la labellisation par IA du dataset actuel. La simulation porte sur l'infrastructure, pas sur le coût complet du projet : rémunération = 30 * tarif horaire humain, tarif non fourni. Le coût du travail humain et une annotation humaine ne sont pas inclus. Le mail ne précise pas le périmètre des tests et du stockage ; les essais et le stockage décrits sont inclus dans ce chiffrage et ce périmètre est à confirmer avec Clara.
AWS Paris eu-west-3, Linux, Shared, OnDemand. c6i.xlarge : 4 vCPU, 8 GiB, 0.202 USD/h. IPv4 publique allouée uniquement pendant la marche : 0.005 USD/h. Total 0.207 USD/h = 0.185052744502056 EUR/h. Affichage arrondi à 0.1851 EUR/h. 30 h = 5.551582335062 EUR HT. Prix CPU et IPv4 combinés pour ne pas masquer le coût réseau.
Disque EBS gp3 de travail : 50 GiB * 0.0928 USD/GiB-mois = 4.64 USD/mois = 4.148042195602 EUR/mois. Provision d'un mois entier, coût affiché 4.15 EUR ; AWS prorate l'usage réel. Le disque et les images ne sont pas conservés indéfiniment.
Dataset source observé : 1506 JPEG, 37272490 octets (37.27 Mo décimaux). S3 Standard Paris : 0.024 USD/GB-mois avec unité binaire 2^30 octets. PUT : 0.0053 USD/1000 ; GET : 0.00042 USD/1000. Hypothèse : un mois, un envoi et une lecture par fichier. Total source images/accès = 0.00944742507238 USD = 0.008445758155175 EUR, affiché 0.01 EUR.
Somme non arrondie : 9.708070288818504 EUR. Somme des lignes affichées : 5.55+4.15+0.01 = 9.71 EUR. Les tableaux et poids de modèle restent sur le disque de travail provisionné ; pas de stockage séparé ajouté.
Conversion : taux de référence BCE du 08/10/2026, 1 EUR = 1.1186 USD, vérifié le 09/10/2026. Ce n'est pas un taux de facturation bancaire garanti. Prix HT, crédits gratuits et remises non appliqués.
Sources officielles :
https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml
https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/eurofxref-graph-usd.en.html
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv
Catalogue EC2 publié le 08/10/2026, effectif au 01/10/2026. c6i.xlarge SKU CKHYJHTXURYY5WKR ; EBS gp3 SKU XCKADFE99D2BK5V6.
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonS3/current/eu-west-3/index.json
Catalogue S3 publié le 28/09/2026 : Standard SKU 4PMVMK58J94ER67S ; PUT JSTCNSNFYXFKBECA ; GET 9PAT6GAN57PYFYDZ.
https://aws.amazon.com/ec2/pricing/on-demand/
https://aws.amazon.com/ebs/pricing/
https://aws.amazon.com/s3/pricing/
https://aws.amazon.com/vpc/pricing/
'@

    $s=$deck.Slides.Item(12)
    Heading $s '4 millions d''images : en une fois ou sur un an'
    [void](Text $s 'Modèle déjà entraîné. Hypothèse de coût : 8,2 h de VM pour 4 M images, soit 41 min par lot mensuel. Durée AWS à vérifier.' 48 135 864 48 19)
    $rows=@(
        @('Dépense sur 12 mois','En une fois`n4 M au mois 1','Glissant, année 1`n≈ 333 333 images / mois'),
        @('Machine : 0,1851 € / heure`n8,2 h de VM dans l''année','1,52 €','1,52 €'),
        @('Images : ≈ 99 Go, 12 mois / image`n1,98 € / mois au stock complet','23,74 €','12,86 €'),
        @('Disque : 50 Gio à 4,148 € / mois','1 mois : 4,15 €','12 mois : 49,78 €'),
        @('Accès : 1 envoi + 1 lecture / image','20,45 €','20,45 € (≈ 1,70 € / lot)'),
        @('TOTAL CLOUD HT','49,86 €','84,61 €')
    )
    # Expand the intentional line breaks in native table cells.
    for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $rows[$r].Count;$c++){$rows[$r][$c]=$rows[$r][$c].Replace('`n',"`n")}}
    $table=Table $s $rows 48 192 864 252 @(398,222,244) 17
    for($r=1;$r -le 6;$r++){$table.Table.Rows.Item($r).Height=[single]$(if($r -le 3){47}else{35})}
    for($c=1;$c -le 3;$c++){$table.Table.Cell(6,$c).Shape.TextFrame.TextRange.Font.Bold=-1}
    [void](Text $s "Stockage des images : 0,16 à 1,98 € / mois en année 1. Ensuite : total cloud de 95,49 € / an.`nVM arrêtée entre les lots ; disque supprimé après le lot unique." 48 452 864 43 17)
    Foot $s 'Infrastructure seule, hors prototype, taxes et contrôle humain. Une copie des images ; aucune sauvegarde supplémentaire.'
    Note $s @'
Objectif : comparer deux calendriers pour 4 millions d'images à labelliser, avec le modèle CNN supervisé déjà entraîné sur le prototype. Il s'agit d'un coût de prédiction et de stockage, pas d'un nouvel entraînement sur 4 M images. Le prototype à 9.71 EUR de la slide 11 n'est pas additionné dans ce tableau. Si le budget de 5000 EUR inclut aussi ce prototype, ajouter 9.71 EUR une seule fois, et seulement si les ressources des deux phases sont facturées séparément. Le texte de Clara ne précise pas ce périmètre.
Hypothèse de calcul illustratif, visible dans la diapositive : 8.2 heures de VM AWS pour 4 M. Cette durée est l'extrapolation LOCALE après égalisation, arrondie depuis 8.19538777778184 h, et n'est PAS une mesure AWS. Il faut la remplacer par une durée cloud mesurée. Prix VM+IPv4 = 0.207 USD/h / 1.1186 = 0.185052744502056 EUR/h. 8.2 h = 1.517432504917 EUR. Chaque heure AWS supplémentaire augmente le calcul de 0.185052744502056 EUR ; la partie stockage dépend de sa propre durée de conservation. Un modèle/optimiseur persistant n'est pas supposé pour la production : chargement du CNN puis inférence avec le même prétraitement et traitement par lots. Les démarrages mensuels, installation et transferts doivent être inclus dans la mesure réelle ; les 41 minutes sont 8.2*60/12 et non un benchmark d'un lot cloud.
En une fois : 4 M images au début du mois 1, une prédiction par image. VM active pendant le traitement, puis arrêtée ; IP libérée. Le disque gp3 de travail de 50 GiB a une provision d'un mois entier puis est supprimé. Les images S3 sont conservées 12 mois, puis supprimées. Conserver le modèle et le fichier de labels dans S3 représente un volume faible par rapport aux 99 Go ; ce volume devra être ajouté au devis réel.
Glissant : environ 333333 images arrivent au début de chacun des 12 mois. Pour totaliser exactement 4 M, 333333 pendant onze mois et 333337 au dernier. Une prédiction par image, 12 lots ; VM arrêtée et IP libérée entre les lots. Le disque de travail de 50 GiB est conservé 12 mois et donc facturé pendant les arrêts. Chaque image est conservée pendant 12 mois : au-delà du mois 12, suppression des images atteignant cet âge pour éviter une croissance indéfinie. L'année 1 part d'un stockage vide et remplit progressivement 99 Go ; à partir de l'année 2, le stock stable représente 4 M images si le même rythme et la même taille moyenne continuent.
Taille réellement observée : 37272490 octets /1506 JPEG = 24749.32934926959 octets/image. Hypothèse : mêmes JPEG et mêmes tailles pour 4 M. Volume = 98997317397.07837 octets = 98.997317397 Go décimaux = 92.19843651827271 GiB facturés S3. Tarif Standard Paris 0.024 USD/GiB-mois = 0.021455390667 EUR/GiB-mois. Stock complet = 1.978153474377387 EUR/mois.
En une fois : 12 * 1.978153474377387 = 23.73784169252864 EUR de stockage des images.
Glissant : approximation de douze apports égaux au début des mois. Somme des volumes mensuels (1+2+...+12)/12 = 6.5 volumes complets-mois. Stockage année 1 = 12.85799758345302 EUR. Répartition exacte 11*333333+333337 : multiplicateur 6.4999945, différence négligeable à l'arrondi au centime. Mois 1 = 0.164846122864782 EUR ; mois 12 = 1.978153474377387 EUR. À partir de l'année 2, 23.73784169252864 EUR/an si suppression après 12 mois et même rythme d'arrivée.
Disque : 50 GiB * 0.0928 USD/GiB-mois /1.1186 = 4.148042195601645 EUR/mois. Une provision mensuelle pour le lot unique : 4.148042195601645 EUR. Douze mois pour le glissant : 49.77650634721974 EUR. Il ne contient pas toutes les images à la fois. L'usage réel EBS est proratisé ; le lot unique peut donc coûter moins que ce mois provisionné si le volume est supprimé plus tôt.
Requêtes : une PUT et une GET par JPEG. 4 M * (0.0000053+0.00000042) USD /1.1186 = 20.45413910244949 EUR, soit environ 1.704511591871 EUR par lot mensuel. Une seule copie des images, un CSV consolidé pour les labels. EC2 et S3 dans la même région Paris, sans NAT Gateway ; entrée des images et transfert S3 vers EC2 dans la même région gratuits, hors requêtes. Pas de téléchargement des 4 M images vers Internet, uniquement le CSV de labels. Pas de sauvegarde, de réplication ou de réentraînement mensuel inclus. Rémunération, annotation humaine, support payant et taxes exclus.
Totaux non arrondis : en une fois, 1.517432504917+23.737841692529+4.148042195602+20.454139102449 = 49.857455495496 EUR. Glissant première année = 1.517432504917+12.857997583453+49.77650634722+20.454139102449 = 84.606075538039 EUR. Glissant à régime stable = 95.485919647115 EUR/an. Sommes des lignes affichées : 1.52+23.74+4.15+20.45=49.86 ; 1.52+12.86+49.78+20.45=84.61. Le budget de 5000 EUR paraît compatible avec l'infrastructure dans CE scénario ; cela ne garantit ni la facture réelle, ni le débit AWS, ni la qualité des labels.
Prix AWS Paris au 09/10/2026 ; conversion BCE du 08/10/2026 : 1 EUR = 1.1186 USD. Taux indicatif, pas taux de facturation garanti. Mêmes références de catalogues et SKU que la slide 11.
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonS3/current/eu-west-3/index.json
https://aws.amazon.com/ec2/pricing/on-demand/
https://aws.amazon.com/ebs/pricing/
https://aws.amazon.com/s3/pricing/
https://aws.amazon.com/vpc/pricing/
https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml
Source locale : notebook 01_exploration_dataset.ipynb, benchmark final après égalisation et mesures de ressources.
'@

    $s=$deck.Slides.Item(13)
    Heading $s 'Limites techniques du passage à l''échelle'
    [void](Text $s 'Après égalisation, en local : 1 000 images en 7,4 s → 8,2 h estimées pour 4 M. Le débit AWS reste inconnu.' 48 134 864 44 20 $true)
    $rows=@(
        @('Limite','Constat','Condition pour traiter 4 M images'),
        @('RAM','Pic local du processus Python : 3,24 Gio, modèles déjà chargés.','Chargement par lots obligatoire : 4 M tenseurs occupent ≈ 2,2 Tio. VM prévue : 8 Gio.'),
        @('Lecture et calcul','4 M petits fichiers peuvent ralentir la lecture et le décodage.','Mesurer le débit AWS, précharger les lots. GPU à tester si le CNN limite la vitesse.'),
        @('Taille des fichiers','Chiffrage pour des JPEG de 24,75 ko / image, soit ≈ 99 Go.','Mesurer les fichiers reçus : des PNG plus lourds augmenteraient le volume et la durée.'),
        @('Qualité des labels','Seulement 20 images connues sont réservées au test.','Vérifier la qualité sur un échantillon représentatif du nouveau dataset.')
    )
    [void](Table $s $rows 48 195 864 287 @(154,302,408) 17)
    Foot $s 'Conditions : débit et RAM à vérifier sur AWS ; données plus lourdes, réentraînements et contrôle humain à rechiffrer.'
    Note $s @'
Les limites sont séparées entre observations réelles et conditions de passage à l'échelle.
Observations locales, après les ajouts du mentor : 1000 images en 7.375849000003654 s sur CPU, avec read_image RGB, égalisation des histogrammes, prétraitement ImageNet, lots de 16 et CNN supervisé déjà entraîné. Extrapolation = 8.19538777778184 h pour 4 M. CPU local Intel i9-14900K déclaré par l'utilisateur dans sa présentation ; torch.get_num_threads() = 24 lors de la mesure. La c6i.xlarge AWS dispose de 4 vCPU, 8 GiB, et ne peut être supposée aussi rapide que cette machine. Aucun benchmark AWS ou GPU n'a été effectué. La slide 10 et les dix premières slides sont conservées dans leur état source, conformément au périmètre de révision limité aux trois slides financières. La mesure présentée ici est la plus récente, après égalisation.
RAM physique RSS estimée avec psutil toutes les 0.05 s, processus Python seulement, hors VRAM GPU. Création du modèle, extraction, essais et entraînements : pic 3310.07421875 MiB = 3.232494354 GiB. Traitement final : pic 3322.515625 MiB = 3.244644165 GiB. Le processus conserve les tableaux et modèles chargés plus tôt ; ce n'est pas la mémoire supplémentaire d'un seul lot. Pendant l'inférence, RAM au départ 3313.140625 MiB et hausse maximale observée 9.375 MiB.
Le notebook prépare actuellement une liste de toutes les images pour l'extraction et des TensorDataset contenant tous les tenseurs pour les entraînements ; ces parties ne sont pas adaptées à une préparation de 4 M images. Un tenseur float32 RGB 224*224 occupe 3*224*224*4 = 602112 octets. 4 M = 2408448000000 octets = environ 2.19047 TiB, sans modèles ni autres buffers. La prédiction finale utilise déjà des lots de 16 ; la généralisation du chargement progressif est la condition de mémoire, pas une augmentation infinie de la RAM. 8 GiB est une configuration à tester et non un minimum validé pour tous les nouveaux jeux.
Lecture/calcul : les 4 M fichiers et la latence S3 peuvent limiter le débit, même si le CNN est accéléré. Il s'agit d'hypothèses de blocage, pas d'un profilage cloud déjà effectué. Mesurer séparément lecture, égalisation, préparation des lots et calcul sur un échantillon de nouvelles images ; mesurer aussi un volume assez long pour dépasser le seul effet du cache. Préchargement en parallèle et disque temporaire peuvent être testés. Un GPU peut réduire le temps CNN mais pas supprimer toutes les attentes de lecture ; aucun gain GPU n'est inventé.
Volume et coût : moyenne mesurée sur les 1506 JPEG sources, 24.749329349 ko décimaux. L'énoncé annonce des PNG et emploie radiographies/IRM ; on ne suppose pas les nouveaux fichiers identiques sans vérification. Coût S3 et débit doivent être recalculés avec les tailles, formats et résolutions réellement livrés. La politique de suppression après 12 mois est nécessaire pour stabiliser le stockage glissant.
Qualité : les ensembles forts restent séparés, avec test de 20 images jamais utilisé pour l'entraînement ou la sélection. Après égalisation, accuracy supervisée 0.90, rappel cancer 0.90, F2 0.90 sur ces 20 images ; semi-supervisé accuracy 0.85, F2 0.8163. Ces résultats exploratoires ne démontrent pas une qualité garantie sur 4 M nouvelles images ou une validation clinique. Les limites de représentativité et les coûts d'un contrôle humain ne sont pas couverts par le seul devis d'infrastructure. Aucun seuil de réussite clinique n'a été fixé.
Sources projet : 01_exploration_dataset.ipynb, données locales, énoncé BrainScanAI et AGENTS.md.
Sources techniques officielles :
https://docs.aws.amazon.com/ec2/latest/instancetypes/co.html
https://aws.amazon.com/s3/pricing/
https://docs.pytorch.org/docs/stable/data.html
https://aws.amazon.com/ec2/instance-types/g4/
'@

    if($deck.Slides.Count -ne 13){throw 'Slide count changed.'}
    $deck.SaveAs($candidate,24)
    $renderDir=Join-Path $build 'candidate-renders'
    [void][IO.Directory]::CreateDirectory($renderDir)
    for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renderDir ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
    $source=$ppt.Presentations.Open($sourcePath,-1,0,0)
    $sourceRenders=Join-Path $build 'source-renders'
    [void][IO.Directory]::CreateDirectory($sourceRenders)
    try{for($i=1;$i -le 13;$i++){$source.Slides.Item($i).Export((Join-Path $sourceRenders ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}}finally{$source.Close()}
    if((Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash -ne $notebookHash){throw 'Notebook changed.'}
    [pscustomobject]@{NotebookHash=$notebookHash;Source=$sourcePath;Candidate=$candidate;Slides=13} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
} finally {$deck.Close()}


