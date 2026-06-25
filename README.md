# Automatisation du traitement des absences RH avec UiPath

## Description du projet

Ce projet est un robot RPA développé avec **UiPath Studio** dans le cadre d’un projet académique.
Il automatise le traitement des absences RH reçues par email sous forme de fichiers Excel.

Le robot récupère automatiquement les pièces jointes depuis Outlook, lit les fichiers Excel, nettoie les données, génère un rapport par région, puis envoie chaque rapport au responsable RH concerné.

L’objectif principal est de réduire les tâches manuelles répétitives, limiter les erreurs humaines et améliorer la traçabilité du processus grâce à une architecture modulaire, une gestion des erreurs robuste et un système de logs.

---

## Fonctionnalités principales

* Récupération automatique des pièces jointes Outlook.
* Lecture des fichiers Excel d’absences RH.
* Lecture du référentiel des destinataires.
* Nettoyage des données avec DataTable.
* Suppression des lignes vides.
* Correction des espaces inutiles avec Trim.
* Filtrage des absences par région.
* Génération automatique d’un rapport Excel par région.
* Envoi automatique des rapports par email via Outlook.
* Gestion des erreurs métier avec BusinessRuleException.
* Gestion des erreurs techniques avec Try/Catch.
* Logging à chaque étape clé du processus.

---

## Technologies utilisées

* UiPath Studio
* Outlook Classic Automation
* Excel Automation
* DataTable
* VB.NET
* Try/Catch
* BusinessRuleException
* Logging UiPath

---

## Architecture du projet

Le robot est organisé en plusieurs workflows afin de respecter une architecture modulaire et maintenable.

```text
Projet_UiPath_Absences_RH/
│
├── Main.xaml
│
├── Workflows/
│   ├── Init_Config.xaml
│   ├── Get_Email_Attachments.xaml
│   ├── Read_Excel_Files.xaml
│   ├── Clean_Data.xaml
│   ├── Generate_Reports_By_Region.xaml
│   └── Send_Emails.xaml
│
├── Data/
│   ├── Input/
│   │   └── INPUT_Absences_RH_Janvier2024.xlsx
│   │
│   ├── Output/
│   │   └── Absences_[Region]_Janvier2024.xlsx
│   │
│   └── REFERENTIEL_Destinataires.xlsx
│
└── Logs/
```

---

## Description des workflows

### Main.xaml

`Main.xaml` est le point d’entrée du robot.
Il orchestre l’exécution complète du processus en appelant les différents sous-workflows avec `Invoke Workflow File`.

Il contient également un `Try/Catch` global permettant de gérer les erreurs métier et les erreurs techniques.

---

### Init_Config.xaml

Ce workflow initialise les paramètres nécessaires au robot.

Il prépare notamment :

* Le dossier principal des données.
* Le chemin du fichier des absences.
* Le chemin du référentiel des destinataires.
* Le dossier de sortie des rapports.
* Le mois traité.

Exemple :

```text
dataFolder = Data
inputPath = Data\INPUT_Absences_RH_Janvier2024.xlsx
referentielPath = Data\REFERENTIEL_Destinataires.xlsx
outputFolder = Data\Output
mois = Janvier 2024
```

---

### Get_Email_Attachments.xaml

Ce workflow lit la boîte de réception Outlook et récupère les pièces jointes Excel.

Étapes principales :

1. Lecture des emails Outlook.
2. Parcours des emails reçus.
3. Vérification de la présence de pièces jointes.
4. Sauvegarde des fichiers Excel dans le dossier `Data/Input`.
5. Écriture d’un log pour chaque pièce jointe sauvegardée.

Une `BusinessRuleException` est levée si aucune pièce jointe Excel n’est trouvée.

---

### Read_Excel_Files.xaml

Ce workflow lit les fichiers Excel nécessaires au traitement.

Fichiers lus :

* `INPUT_Absences_RH_Janvier2024.xlsx`
* `REFERENTIEL_Destinataires.xlsx`

Les données sont chargées dans des DataTables :

```text
dtAbsences
dtDestinataires
```

Des contrôles qualité sont effectués afin de vérifier que les fichiers ne sont pas vides ou illisibles.

---

### Clean_Data.xaml

Ce workflow nettoie les données avant le traitement.

Nettoyages effectués :

* Suppression des lignes où la colonne `Matricule` est vide.
* Suppression des espaces au début et à la fin des valeurs.
* Nettoyage de toutes les colonnes de la DataTable.

Cette étape est essentielle, car les comparaisons de chaînes dans UiPath sont strictes.

Exemple :

```text
"PACA " ≠ "PACA"
```

Sans nettoyage, certaines régions pourraient ne pas être reconnues correctement.

---

### Generate_Reports_By_Region.xaml

Ce workflow génère les rapports Excel par région.

Étapes principales :

1. Parcours du référentiel des destinataires.
2. Récupération de la région avec la clé de tri.
3. Filtrage des absences correspondant à cette région.
4. Génération d’un fichier Excel si des absences existent.
5. Ajout du chemin du rapport généré dans une DataTable `dtReports`.

Exemple de fichier généré :

```text
Data/Output/Absences_PACA_Janvier2024.xlsx
```

Si aucune absence n’est trouvée pour une région, le robot écrit un log de type `Warning` et continue le traitement.

Une `BusinessRuleException` est levée si aucun rapport n’est généré.

---

### Send_Emails.xaml

Ce workflow envoie les rapports générés aux responsables RH concernés.

Étapes principales :

1. Parcours de la DataTable `dtReports`.
2. Récupération de la région et du chemin du rapport.
3. Recherche du destinataire correspondant dans le référentiel.
4. Préparation de l’email.
5. Ajout du rapport Excel en pièce jointe.
6. Envoi automatique via Outlook.

Une `BusinessRuleException` est levée si aucun destinataire n’est trouvé pour une région.

---

## Règles métier

| Référence | Règle                                                  | Action                      |
| --------- | ------------------------------------------------------ | --------------------------- |
| RM01      | Un email doit contenir au moins un fichier Excel       | BusinessRuleException       |
| RM02      | Le fichier des absences ne doit pas être vide          | BusinessRuleException       |
| RM03      | Le référentiel des destinataires ne doit pas être vide | BusinessRuleException       |
| RM04      | Au moins un rapport doit être généré                   | BusinessRuleException       |
| RM05      | Chaque région doit avoir un destinataire               | BusinessRuleException       |
| RM06      | Les données doivent être nettoyées avant comparaison   | Clean_Data.xaml obligatoire |

---

## Gestion des erreurs

Le robot distingue deux types d’erreurs.

### Erreurs métier

Les erreurs métier sont gérées avec `BusinessRuleException`.

Exemples :

* Aucune pièce jointe Excel trouvée.
* Fichier des absences vide.
* Référentiel destinataires vide.
* Aucun rapport généré.
* Aucun destinataire trouvé pour une région.

### Erreurs techniques

Les erreurs techniques sont gérées avec `System.Exception`.

Exemples :

* Outlook indisponible.
* Fichier Excel corrompu.
* Dossier Output inexistant.
* Problème réseau.
* Erreur lors de la lecture ou de l’écriture d’un fichier.

---

## Logging

Le robot utilise des logs afin d’assurer la traçabilité du traitement.

Exemples de logs :

```text
Info    === DÉBUT DU TRAITEMENT ===
Info    Pièce jointe sauvegardée depuis : [sujet]
Info    Absences lues : 18 lignes
Info    Référentiel lu : 6 régions
Info    Région : [X] → [N] absences
Info    Fichier généré : Data/Output/Absences_[Region].xlsx
Info    Mail envoyé à : [email] pour [région]
Warning Aucune absence pour la région : [X]
Error   Erreur métier : [message]
Error   Erreur système : [message]
Info    === TRAITEMENT TERMINÉ ===
```

---

## Données manipulées

### Fichiers d’entrée

```text
INPUT_Absences_RH_Janvier2024.xlsx
REFERENTIEL_Destinataires.xlsx
```

Le fichier des absences contient les informations RH à traiter, comme :

* Matricule
* Nom
* Région_Site
* Type_Absence

Le référentiel contient les informations nécessaires à l’envoi des emails :

* Clé_Tri
* Email
* Copie_CC
* Objet_Mail_Template

---

## Fichiers de sortie

Le robot génère un fichier Excel par région.

Exemple :

```text
Absences_PACA_Janvier2024.xlsx
Absences_IDF_Janvier2024.xlsx
Absences_Nord_Janvier2024.xlsx
```

Chaque fichier est ensuite envoyé automatiquement au destinataire correspondant.

---

## Scénarios de test

| Cas | Scénario testé                 | Résultat attendu                       |
| --- | ------------------------------ | -------------------------------------- |
| 1   | Cas normal                     | Rapports générés et emails envoyés     |
| 2   | Email sans pièce jointe        | BusinessRuleException                  |
| 3   | Fichier absences vide          | BusinessRuleException                  |
| 4   | Référentiel destinataires vide | BusinessRuleException                  |
| 5   | Région sans absence            | Log Warning et poursuite du traitement |
| 6   | Aucune région commune          | BusinessRuleException                  |

---

## Prérequis

Avant d’exécuter le robot, il faut vérifier que :

* UiPath Studio est installé.
* Outlook Classic est configuré.
* Les fichiers Excel sont au format `.xlsx`.
* Le dossier `Data/Input` existe.
* Le dossier `Data/Output` existe.
* Le référentiel des destinataires est complet.
* Les feuilles Excel portent les noms attendus.
* Outlook est ouvert ou accessible par UiPath.

---

## Lancement du projet

1. Ouvrir le projet dans UiPath Studio.
2. Vérifier les chemins configurés dans `Init_Config.xaml`.
3. Vérifier la présence des fichiers dans le dossier `Data`.
4. Lancer le workflow `Main.xaml`.
5. Consulter les logs après l’exécution.
6. Vérifier les rapports générés dans `Data/Output`.
7. Vérifier l’envoi des emails depuis Outlook.

---

## Limitations

* Le robot est prévu pour Outlook Classic.
* Les fichiers doivent être au format `.xlsx`.
* Les comparaisons de régions sont sensibles aux espaces et à la casse.
* Le dossier `Data/Output` doit exister avant l’exécution.
* Un volume élevé d’emails peut augmenter le temps de traitement.
* Le projet est exécuté localement dans sa version actuelle.

---

## Améliorations possibles

* Déploiement en mode Unattended via UiPath Orchestrator.
* Utilisation des Queues Orchestrator pour un traitement transactionnel.
* Support du New Outlook avec Microsoft Office 365 Activities.
* Archivage automatique des fichiers traités dans SharePoint.
* Ajout d’un tableau de bord Power BI.
* Gestion de plusieurs boîtes mail.
* Passage vers une architecture REFramework.
* Ajout d’un rapport global de synthèse.

---

## Compétences démontrées

Ce projet met en avant les compétences suivantes :

* Développement RPA avec UiPath.
* Automatisation Outlook.
* Automatisation Excel.
* Manipulation de DataTables.
* Structuration de workflows modulaires.
* Gestion des erreurs métier et techniques.
* Logging et traçabilité.
* Nettoyage et transformation de données.
* Génération automatique de rapports.
* Envoi automatique d’emails avec pièces jointes.

---


## Conclusion

Ce projet démontre la mise en place d’un robot RPA complet, modulaire et robuste pour automatiser un processus RH répétitif.

Grâce à UiPath, le traitement des absences RH devient plus rapide, plus fiable et mieux tracé.
L’architecture du robot permet également de faire évoluer facilement la solution vers un déploiement en production avec UiPath Orchestrator.
