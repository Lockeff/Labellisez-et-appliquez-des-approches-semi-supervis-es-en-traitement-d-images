param([string]$Workspace=(Get-Location).Path)
$ErrorActionPreference='Stop'
$build=Join-Path $Workspace '.build/brainscanai/aws-i9'
$source=Join-Path $Workspace 'output/presentations/BrainScanAI_presentation_13_slides_actualisee.pptx'
$candidate=Join-Path $build 'candidate.pptx'
if(Test-Path -LiteralPath $candidate){throw 'Candidate exists.'}
$notebookHash=(Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$data=Get-Content -LiteralPath (Join-Path $build 'costs-fixed-fx.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$cost=$data.display
$helper=Get-Content -LiteralPath (Join-Path $Workspace '.build/brainscanai/create_deck.ps1') -Raw -Encoding UTF8
$start=$helper.IndexOf('function RGB(');$end=$helper.IndexOf('function Picture(')
Invoke-Expression $helper.Substring($start,$end-$start)
$ppt=New-Object -ComObject PowerPoint.Application
$deck=$ppt.Presentations.Open($source,0,0,0)
function Heading($Slide,$Title){
 for($i=$Slide.Shapes.Count;$i -ge 1;$i--){$Slide.Shapes.Item($i).Delete()}
 [void](Text $Slide $Title 48 30 866 80 32 $true)
 [void](Line $Slide 48 114 912 114)
 [void](Text $Slide ([string]$Slide.SlideIndex) 878 508 34 18 12 $false $gray 3)
}
$common=@'
Sources officielles consultées le 09/10/2026 :
https://www.intel.com/content/www/us/en/products/sku/236773/intel-core-i9-processor-14900k-36m-cache-up-to-6-00-ghz/specifications.html
https://docs.aws.amazon.com/ec2/latest/instancetypes/co.html
https://aws.amazon.com/ec2/instance-types/c7i/
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/eu-west-3/index.csv
https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonS3/current/eu-west-3/index.json
https://aws.amazon.com/vpc/pricing/
https://aws.amazon.com/ebs/pricing/
https://aws.amazon.com/s3/pricing/
https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/eurofxref-graph-usd.en.html
Prix On-Demand Linux, tenancy Shared, aucun logiciel payant préinstallé, capacité Used, Paris eu-west-3. Catalogue EC2 publié le 08/10/2026, version 20261008184850, effectif 01/10/2026. c7i.12xlarge : SKU RKFRM5JW99YQMR8K, 2.5452 USD/h, 48 vCPU, 96 GiB. IPv4 publique : 0.005 USD/h, libérée pendant les arrêts. Total machine + IP : 2.5502 USD/h. Même conversion datée que le support source, BCE 08/10/2026 : 1 EUR=1.1186 USD, soit 2.279814053280887 EUR/h, affiché 2.2798 EUR/h. Ce taux historique, vérifié dans la table BCE, reste volontairement fixe pour conserver les lignes de stockage du support. Le taux du 09/10/2026 est 1.1206, mais n'est pas utilisé dans ce devis. Prix HT, taux bancaire réel non garanti, aucune remise ni crédit gratuit.
EBS gp3 Paris : 0.0928 USD/GiB-mois, 50 GiB = 4.148042195601645 EUR/mois. S3 Standard Paris : 0.024 USD/GiB-mois. PUT 0.0053 USD/1000, GET 0.00042 USD/1000. Une seule copie des fichiers, un envoi et une lecture par image. Même région EC2/S3, sans NAT Gateway, aucun téléchargement complet vers Internet. Pas de sauvegarde, réplication, annotation humaine ou support payant chiffré. Poids et CSV sur le disque de travail. Les volumes supplémentaires doivent être ajoutés au devis réel.
La c7i.12xlarge est un candidat de dimensionnement CPU prudent, pas un équivalent de puissance démontré. i9-14900K : 24 cœurs physiques hybrides (8 Performance et 16 Efficient), 32 threads matériels. c7i.12xlarge : 24 cœurs Xeon Sapphire Rapids homogènes, 48 threads logiques/vCPU, deux threads par cœur, 96 GiB. La référence i9 utilise 24 threads PyTorch pour l'exécution mesurée. Choix : conserver au moins 24 cœurs physiques côté AWS, architecture x86 Intel compatible, CPU soutenu de la gamme C7i standard. Ce n'est pas une instance Flex à performance de base partielle. Les 96 GiB sont liés à la taille CPU de l'instance, pas un minimum RAM exigé par le notebook. L'égalité du nombre de cœurs, du nombre de threads ou d'une fréquence ne démontre pas une égalité de débit. Aucun benchmark public trouvé ne compare exactement ce notebook FP32 ResNet18, lecture JPEG, égalisation, ImageNet et lots de 16 sur le i9 et cette VM. Aucun benchmark cloud n'a été effectué.
Temps source conservé : 1000 images en 7.4 s affichées, mesure locale non arrondie 7.3758490000036545 s, 4 M extrapolées en 8.19538777778184 h. Scénario chiffré : 8.2 heures de VM AWS pour 4 M, conditionné au maintien du débit local. Ce n'est pas une prédiction vérifiée de la durée AWS. Mesurer la même chaîne de bout en bout et les mêmes réglages (24 threads PyTorch, lots de 16, images identiques, prétraitement égalisé, ResNet18 supervisé déjà entraîné). Séparer la lecture des fichiers/S3 et le calcul si le débit baisse. Répéter sur un lot plus grand pour vérifier les effets de cache et inclure préparation, transferts, démarrages et chargement des poids dans la durée facturable.
Coût de prédiction paramétrable : coût fixe de stockage/requêtes/disque + durée AWS réelle * 2.279814053280887 EUR/h. Les 30 h de prototypage représentent une provision de VM active pendant le travail et les essais, pas le temps d'un seul entraînement. Les budgets de la mission sont 300 EUR pour la labellisation IA du prototype et 5000 EUR pour 4 M. Leur périmètre exact tests/stockage est à confirmer avec Clara. Les dépenses décrites incluent les tests dans les 30 h, et les postes de stockage affichés. Rémunération et contrôle humain exclus. Le budget ne vaut pas garantie de qualité.
'@
try {
 $s=$deck.Slides.Item(10)
 Heading $s 'Référence i9 et dimensionnement AWS'
 [void](Text $s 'Machine retenue pour le devis : c7i.12xlarge à Paris. Objectif : retrouver le débit mesuré sur le i9-14900K.' 48 134 864 54 22)
 $rows=@(
  @('Critère CPU','PC de référence','AWS : c7i.12xlarge'),
  @('Processeur','Intel Core i9-14900K','Intel Xeon Sapphire Rapids'),
  @('Cœurs physiques','24 : 8 Performance + 16 Efficient','24 cœurs Xeon'),
  @('Threads matériels','32','48 vCPU : 2 threads / cœur')
 )
 $table=Table $s $rows 48 208 864 173 @(205,327,332) 19
 [void](Text $s "Même nombre de cœurs physiques, architectures différentes. L'équivalence de puissance reste à vérifier avec le même code." 48 400 864 54 21 $true)
 [void](Text $s 'Référence locale : 1 000 images en 7,4 s, soit ≈ 8,2 h pour 4 millions.' 48 468 864 30 20)
 Foot $s 'AWS : 96 Gio de RAM. La taille CPU guide le choix de la VM. Sources : Intel et AWS, détails en notes.'
 Note $s ("Le PC i9 constitue la base réelle de comparaison. Les 24 cœurs AWS sont tous des cœurs Xeon, tandis que le i9 combine 8 cœurs Performance et 16 Efficient. Le nombre de vCPU AWS est un nombre de threads et ne doit pas être confondu avec les cœurs physiques. Une c7i.8xlarge aurait 32 vCPU mais seulement 16 cœurs physiques. Le choix c7i.12xlarge conserve 24 cœurs physiques pour établir le devis, sans prétendre que ce critère suffit à prouver une puissance identique. La fréquence du i9 (jusqu'à 6 GHz sur les conditions Intel de turbo) et le turbo tous cœurs du Xeon (3.2 GHz) ne sont pas directement comparables, et aucun rapport de vitesse n'en est déduit.`n`n"+$common)

 $s=$deck.Slides.Item(11)
 Heading $s 'Prototypage : 30 heures de travail'
 [void](Text $s 'Hypothèse : c7i.12xlarge allumée pendant 30 h, essais et entraînements inclus. 24 cœurs physiques, 48 vCPU, 96 Gio de RAM.' 48 135 864 57 21)
 $rows=@(
  @('Poste AWS','Tarif et unité','Quantité / fréquence','Coût HT'),
  @('Machine + adresse IPv4',($cost.hour_eur+' € / heure'),'30 h, une seule fois',($cost.machine_30h+' €')),
  @('Disque de travail, 50 Gio','4,148 € / mois','1 mois de conservation','4,15 €'),
  @('Images et accès au stockage','≈ 0,01 € / dataset','37 Mo ; 1 envoi + 1 lecture/image','0,01 €'),
  @('TOTAL CLOUD','','Pour ce prototype',($cost.prototype_total+' €'))
 )
 $table=Table $s $rows 48 213 864 215 @(253,191,286,134) 18
 for($c=1;$c -le 4;$c++){$table.Table.Cell(5,$c).Shape.TextFrame.TextRange.Font.Bold=-1}
 [void](Text $s 'Estimation sous le budget IA de 300 €. Hors temps humain. Les essais et les entraînements sont inclus dans les 30 h de VM.' 48 447 864 48 20 $true)
 Foot $s 'EUR HT : AWS Paris au 09/10/2026. Conversion BCE du 08/10/2026 : 1 € = 1,1186 $.'
 Note $s ("Prototypage de 30 heures de travail demandé par l'utilisateur. VM supposée active pendant ces 30 h, y compris installation, essais, extraction, hyperparamètres et entraînements. Cela ne signifie pas que l'entraînement seul dure 30 h. Arrêter la VM pendant la rédaction réduit son coût. Calcul exact : VM+IP = "+$data.exact.machine_30h+" EUR, disque = "+$data.exact.disk_month+" EUR, images + accès dataset = "+$data.exact.prototype_images+" EUR, total = "+$data.exact.prototype_total+" EUR. Somme des lignes affichées : 68.39+4.15+0.01 = 72.55 EUR. Temps humain non chiffré, conformément à la mention choisie par l'utilisateur dans la source.`n`n"+$common)

 $s=$deck.Slides.Item(12)
 Heading $s '4 millions d''images : en une fois ou sur un an'
 [void](Text $s 'Même c7i.12xlarge. Hypothèse : débit équivalent au i9, donc 8,2 h de VM au total ou 41 min par lot mensuel. À vérifier sur AWS.' 48 135 864 48 19)
 $rows=@(
  @('Dépense sur 12 mois',"En une fois`n4 M au mois 1", "Glissant, année 1`n≈ 333 333 images / mois"),
  @(("Machine : "+$cost.hour_eur+" € / heure`n8,2 h de VM dans l'année"),($cost.compute_4m+' €'),($cost.compute_4m+' €')),
  @("Images : ≈ 99 Go, 12 mois / image`n1,98 € / mois au stock complet",'23,74 €','12,86 €'),
  @('Disque : 50 Gio à 4,148 € / mois','1 mois : 4,15 €','12 mois : 49,78 €'),
  @('Accès : 1 envoi + 1 lecture / image','20,45 €','20,45 € (≈ 1,70 € / lot)'),
  @('TOTAL CLOUD HT',($cost.one_shot+' €'),($cost.rolling_year1+' €'))
 )
 $table=Table $s $rows 48 192 864 252 @(398,222,244) 17
 for($r=1;$r -le 6;$r++){$table.Table.Rows.Item($r).Height=[single]$(if($r -le 3){47}else{35})}
 for($c=1;$c -le 3;$c++){$table.Table.Cell(6,$c).Shape.TextFrame.TextRange.Font.Bold=-1}
 [void](Text $s ("Stockage : 0,16 à 1,98 € / mois en année 1. Ensuite : total cloud de "+$cost.rolling_steady+" € / an.`nVM arrêtée entre les lots ; disque supprimé après le lot unique.") 48 452 864 43 17)
 Foot $s 'Sous 5 000 € dans ce scénario. Infrastructure seule, hors prototype, taxes et contrôle humain. Une copie des images.'
 Note $s ("Prédiction avec CNN déjà entraîné, sans réentraînement sur 4 M. Hypothèse conditionnelle d'équivalence de débit au i9 : 8.2 h AWS, soit 41 min par lot mensuel. La machine est identique à celle du prototype. Calcul = durée réelle AWS * tarif horaire ; à vitesse différente, remplacer 8.2 h par la mesure AWS. Machine : "+$data.exact.compute_4m+" EUR. Une fois : images conservées 12 mois, disque de travail conservé un mois puis supprimé, VM arrêtée et IP libérée après le traitement. Total = "+$data.exact.one_shot+" EUR, somme affichée 18.69+23.74+4.15+20.45 = 67.03 EUR. Glissant : 333333 images au début de onze mois et 333337 au dernier, 12 mois de conservation par image, VM arrêtée entre lots, disque conservé 12 mois. Le premier exercice part d'un stockage vide. Stockage S3 année 1 = 6.5 volumes complets-mois, 12.85799758345302 EUR. Total première année = "+$data.exact.rolling_year1+" EUR, somme affichée 18.69+12.86+49.78+20.45 = 101.78 EUR. À partir de l'année 2, stock stable de 4 M images si suppression après 12 mois et même rythme annuel, total = "+$data.exact.rolling_steady+" EUR/an. Le prototype à 72.55 EUR est exclu de cette table, ajouter une fois si le budget 5000 EUR comprend également cette phase. Requêtes : une PUT et une GET par image. Budget de contrôle humain et réentraînements non inclus.`n`n"+$common)

 $s=$deck.Slides.Item(13)
 Heading $s 'Conditions et limites du passage à l''échelle'
 [void](Text $s 'Condition du devis : la c7i.12xlarge doit retrouver le débit du i9 avec le même traitement. Les 8,2 h restent une hypothèse.' 48 134 864 50 20 $true)
 $rows=@(
  @('Limite','Constat','Condition pour traiter 4 M images'),
  @('Puissance CPU','24 cœurs physiques de chaque côté, mais des architectures différentes.','Comparer le même code sur 1 000 images : référence i9 de 7,4 s.'),
  @('RAM','Processus Python : environ 3,2 Gio. VM retenue : 96 Gio.','Lire par lots : 4 M tenseurs en mémoire occuperaient ≈ 2,2 Tio.'),
  @('Lecture et calcul','4 M petits fichiers, décodage et accès S3 peuvent limiter le débit.','Mesurer la chaîne complète, avec égalisation et lots de 16. Tester le préchargement.'),
  @('Taille des fichiers','JPEG de 24,75 ko / image, soit ≈ 99 Go dans ce devis.','Recalculer volume et durée pour des PNG plus lourds. Conservation : 12 mois / image.'),
  @('Qualité des labels','Seulement 20 images connues réservées au test.','Vérifier les prédictions sur un échantillon représentatif du nouveau dataset.')
 )
 $table=Table $s $rows 48 194 864 290 @(154,302,408) 17
 for($r=1;$r -le 6;$r++){for($c=1;$c -le 3;$c++){$table.Table.Cell($r,$c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment=1}}
 Foot $s 'VM arrêtée entre lots : allumée 24 h/24 pendant un an, le calcul seul coûterait ≈ 19 971 € HT.'
 Note $s ("Les 48 vCPU ne constituent pas une preuve d'équivalence au i9 : il faut mesurer le débit de la même chaîne. Le temps source affiché est conservé, sans le présenter comme une mesure AWS. Si AWS est plus lent, le coût de calcul augmente au même tarif horaire, tandis que les postes fixes de stockage suivent les politiques indiquées. À titre de sensibilité, 8760 h de VM active en continu coûteraient "+$data.exact.vm_24_7+" EUR HT avant stockage : dépasser 5000 EUR est possible si la VM reste allumée toute l'année. Le scénario glissant suppose explicitement une VM arrêtée entre les douze lots. RAM : environ 3.2 GiB observés pour l'ensemble du processus Python sur la machine locale, pas un minimum total pour Windows ou Linux. Les 96 GiB AWS sont imposés par la taille CPU choisie. Charger les 4 M tenseurs normalisés RGB 224*224 float32 ensemble occuperait 2408448000000 octets soit environ 2.19 TiB. Le chargement progressif est nécessaire malgré les 96 GiB. Lecture : aucun goulot cloud n'a été déjà mesuré ; préchargement et GPU seulement à tester selon le profil de calcul et lecture. Qualité des 4 M labels non démontrée par un test de 20 images, aucune validation clinique. Le coût du contrôle humain reste à chiffrer.`n`n"+$common)

 if($deck.Slides.Count -ne 13){throw 'Slide count changed.'}
 $deck.SaveAs($candidate,24)
 $renders=Join-Path $build 'candidate-renders';[void][IO.Directory]::CreateDirectory($renders)
 for($i=1;$i -le 13;$i++){$deck.Slides.Item($i).Export((Join-Path $renders ('slide-{0:D2}.png' -f $i)),'PNG',1600,900)}
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Source changed.'}
 if((Get-FileHash -LiteralPath (Join-Path $Workspace '01_exploration_dataset.ipynb')).Hash -ne $notebookHash){throw 'Notebook changed.'}
 [pscustomobject]@{Source=$source;SourceHash=$sourceHash;Candidate=$candidate;NotebookHash=$notebookHash;Slides=13}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $build 'edit-receipt.json') -Encoding UTF8
} finally {$deck.Close()}
