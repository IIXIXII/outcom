# Outcom

Complément pour **Outlook classique LTSC 2024** sous Windows.

## Choix technique initial

Le projet démarre avec un **complément Outlook VSTO en C#**, basé sur **.NET Framework**.

Ce choix convient à Outlook LTSC 2024 lorsqu'une intégration locale au client lourd est nécessaire. VSTO produit une assembly .NET chargée comme complément COM par Outlook, tout en évitant de gérer manuellement `IDTExtensibility2` et toute l'inscription COM.

À retenir :

- VSTO fonctionne avec Outlook classique, pas avec le nouvel Outlook pour Windows.
- VSTO nécessite .NET Framework ; ne pas choisir .NET, .NET Core ou .NET 8/9/10 pour le projet chargé dans Outlook.
- Le poste détecté utilise Outlook LTSC 2024 **64 bits** (`ProPlus2024Volume`).
- Une cible `Any CPU` convient généralement au développement VSTO ; les futurs installateurs devront tenir compte de l'architecture Office.

## Outils à installer

Installer **Visual Studio 2022 Community** (ou Professional/Enterprise), puis sélectionner dans Visual Studio Installer :

1. La charge de travail **Développement Office/SharePoint**.
2. Les composants **Outils de développement Microsoft Office**.
3. Le **pack de ciblage .NET Framework 4.8**.
4. Facultatif pour plus tard : **Microsoft Visual Studio Installer Projects 2022**, utile pour produire un MSI.

Outlook doit être fermé pendant l'installation des outils Office.

Le script suivant vérifie ensuite la machine :

```powershell
.\tools\Check-DevelopmentEnvironment.ps1
```

## Créer la version 0.1 vide

Après installation de Visual Studio :

1. Ouvrir Visual Studio.
2. Choisir **Créer un projet**.
3. Rechercher `Outlook VSTO Add-in`.
4. Choisir le modèle C# **Outlook VSTO Add-in**.
5. Nommer la solution `Outcom` et le projet `Outcom.AddIn`.
6. Choisir **.NET Framework 4.8**.
7. Ne rien ajouter dans `ThisAddIn_Startup` et `ThisAddIn_Shutdown`.
8. Lancer avec `F5` : Visual Studio compile, inscrit temporairement le complément et démarre Outlook.

La première validation consiste seulement à vérifier dans Outlook :

`Fichier > Options > Compléments > Compléments COM > Atteindre`

Le complément `Outcom` doit être présent et coché. Cette version est volontairement sans bouton et sans comportement visible.

## Structure cible

```text
Outcom.sln
src/
  Outcom.AddIn/
tests/
tools/
```

Le projet VSTO doit être créé depuis Visual Studio : le modèle génère plusieurs fichiers spécifiques (`.vstol`, manifeste, designer et paramètres de débogage) qu'il est préférable de ne pas reproduire manuellement.

## Étapes suivantes conseillées

1. Ajouter une journalisation locale minimale du démarrage et de l'arrêt.
2. Ajouter un bouton Ribbon sans traitement métier.
3. Définir précisément les événements Outlook à intercepter.
4. Tester les cas où Outlook désactive un complément lent.
5. Préparer un MSI signé, distinct si un support Office 32 bits devient nécessaire.

## Documentation Microsoft

- [Configurer un ordinateur pour développer des solutions Office](https://learn.microsoft.com/visualstudio/vsto/how-to-configure-a-computer-to-develop-office-solutions)
- [Bien démarrer avec les compléments VSTO](https://learn.microsoft.com/visualstudio/vsto/getting-started-programming-vsto-add-ins)
- [Architecture des compléments VSTO](https://learn.microsoft.com/visualstudio/vsto/architecture-of-vsto-add-ins)
- [Déployer une solution VSTO avec Windows Installer](https://learn.microsoft.com/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)
