# Contexte permanent du projet BrainScanAI

Ces consignes reprennent le contexte et les instructions fournis par l’utilisateur.
Les conserver dans toutes les propositions, analyses et réponses concernant ce projet.

## Règle de travail prioritaire

**Ne jamais modifier un notebook (.ipynb) sans demande explicite de l’utilisateur.**
Cela comprend les cellules, sorties, métadonnées et sauvegardes après exécution.
Une demande d’analyse, de conseil ou de discussion ne constitue pas une demande
de modification. Présenter les propositions dans la conversation tant qu’aucune
modification n’a été demandée. Préserver les modifications existantes de l’utilisateur.

**L’autorisation est limitée à la section explicitement désignée.** Une demande
de modification de l’étape 2 n’autorise aucune modification de l’étape 1, des
imports communs ni de leurs résultats. Une consigne générale de respect de
l’énoncé ne donne pas l’autorisation de réécrire les autres étapes. Ne pas
supprimer un travail précédemment demandé par l’utilisateur hors de la section
autorisée. La dernière clarification de l’utilisateur prime sur toute
interprétation antérieure du périmètre.

## Respect strict de l’énoncé et du cours

- L’utilisateur demande de **ne jamais faire plus que l’énoncé de l’étape en cours**.
  Cette limite s’applique au code, aux explications, aux traitements et aux livrables.
- Utiliser uniquement les outils indiqués dans l’énoncé et le cours fournis.
  Les fonctions de base de Python nécessaires pour parcourir les fichiers et les
  fonctionnalités de Jupyter pour afficher les résultats restent des moyens de réalisation.
- Ne pas ajouter de contrôle, de méthode ou de bibliothèque simplement parce que
  cela pourrait être utile. Une extension nécessite une nouvelle demande explicite.
- Étape 1 : charger le dataset, décrire la structure des fichiers, afficher des
  exemples, vérifier dimensions, canaux, lisibilité et cohérence sur un petit
  échantillon, et noter les observations. Pillow est indiqué dans ses prérequis.
- Étape 2 : redimensionner et normaliser avec le prétraitement du modèle,
  charger un modèle pré-entraîné, geler ses paramètres, extraire un vecteur par
  image et sauvegarder un seul tableau exploitable. Utiliser PyTorch/torchvision,
  NumPy et pandas ; ne pas introduire EXIF, hashlib, recherche de doublons,
  propagation de labels, tableaux dédoublonnés ou préparation du clustering.
- L’évaluation de plusieurs couches est facultative (« si besoin ») : ne pas
  l’ajouter sans besoin établi et demande de l’utilisateur.
- Ne pas supprimer ni modifier les images sources. Les recherches de doublons
  de l’étape 1 ont été demandées par l’utilisateur et doivent être conservées :
  l’autorisation actuelle de simplification concerne uniquement l’étape 2.
- Vérifier que chaque cellule répond directement à une demande de l’énoncé.
  Les objectifs des étapes suivantes ne justifient pas des ajouts à l’étape actuelle.

## Règle pédagogique obligatoire pour les notebooks

L’utilisateur est étudiant et doit pouvoir comprendre chaque opération et son résultat.

- Organiser les explications en paires : **une cellule Markdown, immédiatement
  suivie d’une seule cellule de code, avec son résultat**.
- Chaque Markdown décrit exclusivement le code qui le suit : son objectif,
  les variables nécessaires, les opérations réellement effectuées et le résultat
  affiché ou sauvegardé. Ne pas annoncer une opération exécutée dans une autre cellule.
- Distinguer précisément la configuration d’un traitement de son application :
  définir un prétraitement ne transforme pas encore les images ; charger un modèle
  ne calcule pas encore les embeddings.
- Décrire le résultat attendu de la cellule concernée (tableau, dimensions,
  graphique, compteur ou fichiers sauvegardés). Ne pas présenter une opération
  ou un résultat comme réalisé si le code de cette cellule ne le réalise pas.
- Intégrer les titres de section dans les Markdown associés au code ; éviter
  les introductions qui couvrent plusieurs cellules et les synthèses non associées
  à un code dans cette organisation pédagogique.
- Utiliser des explications courtes, progressives et des noms de variables cohérents
  avec le notebook existant, notamment `dossier_dataset` pour le chemin du dataset.
- Après une révision pédagogique, vérifier la correspondance de chaque paire.
  Préserver le code, les résultats et les sections hors du périmètre demandé,
  sauf si leur modification est explicitement nécessaire et autorisée.

## Scénario professionnel

- Entreprise : CurelyticsIA, startup de e-santé.
- Projet : BrainScanAI, première phase de recherche et développement.
- Rôle : Data Scientist junior spécialisé en Computer Vision.
- Responsable Data Science : Clara Martin.
- Finalité : explorer l’automatisation de la détection de tumeurs cérébrales
  sur des images médicales, notamment des IRM, pour assister les professionnels de santé.
- La majorité des images n’est pas étiquetée ; un petit sous-ensemble est annoté
  par des radiologues experts selon le scénario (normal / cancéreux).

## Données annoncées dans la mission

La pièce jointe annoncée est un ZIP contenant des images PNG, des métadonnées
anonymisées, une documentation technique et une liste restreinte de labels.
Le texte emploie à la fois « radios », « radiographies » et « IRM ».
Vérifier ces éléments sur les fichiers disponibles et signaler les différences
entre le scénario et le dataset, sans présumer la présence de documents,
de métadonnées patients ou d’une validation clinique effectivement vérifiée.

## Objectifs et livrables

1. Explorer les images et extraire des caractéristiques visuelles pertinentes
   avec un modèle pré-entraîné, de type ResNet ou équivalent.
2. Définir et justifier le preprocessing adapté aux modèles utilisés.
3. Réaliser une analyse non supervisée et entraîner plusieurs algorithmes
   de clustering pour rechercher des regroupements naturels.
4. Mettre en œuvre une approche semi-supervisée exploitant les labels partiels
   afin de prédire les étiquettes manquantes.
5. Fournir des livrables Notebook comprenant extraction de features,
   preprocessing, analyse non supervisée, clustering et approche semi-supervisée.
   Leur organisation précise reste à définir avec l’utilisateur.
6. Préparer un support de présentation synthétisant approche, résultats,
   limites et recommandations techniques, notamment pour le passage à l’échelle.

## Contraintes et évaluation

- Travailler en Python ; l’environnement existant utilise uv et Python 3.12.
- Tester plusieurs algorithmes et comparer leurs résultats de manière pertinente.
- Choisir les métriques selon l’erreur la plus importante pour le projet
  (F1, accuracy, précision, rappel ou autre métrique justifiée).
- Distinguer les métriques de clustering de l’évaluation des prédictions de labels.
- Expliciter le choix entre faux négatifs et faux positifs ; aucun arbitrage
  ni seuil de performance n’a encore été validé par l’utilisateur.
- Définir clairement les critères de réussite (« definition of done »)
  avant de déclarer un objectif atteint.
- Lors des propositions d’évaluation, tenir compte des doublons et, si des
  identifiants patients sont disponibles, des groupes patients afin de limiter
  les fuites entre entraînement et évaluation.
- Ne pas présenter des résultats exploratoires comme une validation clinique.

## Budgets et passage à l’échelle

- Budget de labellisation par IA du dataset actuel : **300 euros**.
- Scénario de passage à l’échelle : **4 millions d’images pour 5 000 euros**.
- Le support doit répondre à la question de faisabilité et préciser les conditions,
  hypothèses, coûts, qualité attendue et limites de l’approche proposée.
- Ces budgets concernent la labellisation par IA dans la mission ; ne pas les
  assimiler sans justification à un budget d’annotation humaine ou au coût total du projet.

## Préparation et suivi attendus dans le scénario

- Lire la mission et les documents associés disponibles.
- Prendre des notes sur la compréhension de la mission.
- Consulter les étapes d’accompagnement lorsqu’elles sont disponibles.
- Préparer les questions pour une première session de mentorat.
- Prévoir un point d’étape avant la fin de semaine selon la demande de Clara
  dans le scénario ; aucune date réelle n’a été fixée dans la conversation.
