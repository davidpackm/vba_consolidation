
# 🎯 Objectif du projet

Automatiser, à l’aide de VBA, le nettoyage, l’analyse et le reporting de données hospitalières afin de produire un tableau de bord consolidé, fiable et exploitable, dans un contexte d’organisme public.

Les données sources sont des fichiers Excel mensuels contenant les admissions hospitalières
(Hôpital, Spécialité, Durée de séjour, Coût, Date d’entrée, Sexe).

# 🧩 Architecture générale

Le projet est structuré autour de 4 phases, chacune implémentée dans des modules VBA distincts :

## 1️⃣ Vérification de la structure des données

Contrôle de la présence des colonnes obligatoires

Vérification des formats (dates, numériques, etc.)

Signalement visuel des anomalies (cellules colorées)

### ➡️ Macro principale : VerifierStructure()

## 2️⃣ Nettoyage automatique des données

Suppression des lignes incomplètes

Normalisation des spécialités médicales

Détection et suppression des doublons

Nettoyage de la casse et des espaces

### ➡️ Macro principale : NettoyerDonnees()
### ➡️ Génération d’une feuille Log_Nettoyage retraçant les corrections et suppressions

## 3️⃣ Calcul d’indicateurs & reporting

Création d’un tableau de bord avec :

Durée moyenne de séjour par spécialité

Coût moyen par spécialité

Nombre de séjours par hôpital

Top 3 spécialités les plus coûteuses

Top 3 hôpitaux les plus fréquentés

### ➡️ Macro principale : CalculerIndicateurs()
### ➡️ Mise en forme automatique et graphiques intégrés

## 4️⃣ Consolidation multi-fichiers (mensualisation)

Import automatique de plusieurs fichiers mensuels

Normalisation et fusion des données

Création d’une base consolidée unique

### ➡️ Macro principale : ConsoliderFichiersMois()
### ➡️ Génération de la feuille Base_Complete

## 🖥️ Interface utilisateur

Interface Excel simple et stylisée

Boutons pour lancer les différentes macros :

Vérification

Nettoyage

Calcul des indicateurs

Consolidation

Filtres dynamiques (hôpital, spécialité)

## 📁 Livrables

Fichier Excel .xlsm fonctionnel comprenant :

Données

Base_Complete

Analyse

Modules VBA séparés par phase

Code commenté et structuré
