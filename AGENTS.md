# Contexte permanent du projet BrainScanAI

Ces consignes reprennent le contexte et les instructions fournis par l’utilisateur.
Les conserver dans toutes les propositions, analyses et réponses concernant ce projet.

## Règle de travail prioritaire

**Ne jamais modifier un notebook (.ipynb) sans demande explicite de l’utilisateur.**
Cela comprend les cellules, sorties, métadonnées et sauvegardes après exécution.
Une demande d’analyse, de conseil ou de discussion ne constitue pas une demande
de modification. Présenter les propositions dans la conversation tant qu’aucune
modification n’a été demandée. Préserver les modifications existantes de l’utilisateur.

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
