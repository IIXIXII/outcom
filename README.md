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

Le complément `Outcom` doit être présent et coché. Cette validation initiale s'effectue avant l'ajout de comportement visible.

## Fonctionnalités actuelles

- La version actuelle d’Outcom est **0.1.0**, centralisée dans la propriété `OutcomVersion` du projet. La commande **À propos d’Outcom**, dans l’onglet **Outcom**, affiche la version, la date de compilation embarquée dans le binaire, l’architecture active, le mode de connexion à ChatGPT, le journal local et la licence.
- Le démarrage et l'arrêt du complément sont journalisés dans `%LOCALAPPDATA%\Outcom\Logs\outcom.log`.
- Un onglet **Outcom** est affiché dans la fenêtre principale d'Outlook pour ouvrir le volet, configurer la connexion et définir le **Contexte Codex** ; ses actions utilisent les icônes natives d'Office.
- L'icône générale Outcom identifie le complément, les fenêtres de configuration et l'en-tête du volet ; elle n'est pas utilisée pour représenter les actions.
- Sous l'état et le Contexte Codex, le volet réserve initialement 40 % de sa hauteur utile aux activités en cours et récentes, puis 60 % aux conversations. Cette seconde partie consacre 20 % à la liste des conversations et 40 % à la conversation sélectionnée ; dans cette dernière, environ 28 % accueillent les documents de contexte et 72 % la réponse et la saisie. Il conserve plusieurs conversations multi-tours indépendantes, avec affichage progressif, annulation, copie et création de brouillon.
- Outcom propose une réponse aussi bien dans une fenêtre de rédaction que dans une réponse intégrée au volet de lecture. Deux commandes combinées peuvent également créer un brouillon **Répondre** ou **Répondre à tous**, puis lancer Codex en arrière-plan. La zone de réponse située en haut du message, arrêtée à la signature native d'Outlook ou à la section du courrier cité, est traitée comme une ébauche et des orientations de l'utilisateur, puis remplacée par un message complet.
- La commande reconnaît également un courrier transféré. Les champs actuels **À** et **Cc** définissent alors le véritable public : Codex rédige un message d'accompagnement destiné à ces personnes, pour information et, selon le contenu ou les orientations, pour préciser les actions et suites attendues. Il ne formule pas une réponse adressée à l'expéditeur d'origine.
- La réponse générée reprend la langue du message le plus récent du fil : un fil anglais produit une réponse anglaise même si l'ébauche ou les orientations sont rédigées en français, sauf demande explicite d'une autre langue. Outcom vérifie si la signature conservée contient déjà une formule de politesse : il demande d'en générer exactement une lorsqu'elle manque et aucune lorsqu'elle est déjà présente, puis retire à l'insertion une éventuelle formule ou signature finale identique.
- Le volet affiche dans une zone redimensionnable les demandes Codex actives ou récentes, y compris plusieurs générations indépendantes en parallèle. Il permet d'annuler la demande sélectionnée et d'effacer les tâches terminées.
- Les groupes Outcom du ruban affichent sur deux lignes les compteurs **En cours** et **Terminées**. Ils utilisent exactement le même suivi que la zone Activités du volet et sont actualisés au démarrage, à la fin, à l'annulation ou au nettoyage d'une opération.
- Le **Contexte Codex** centralise le modèle, la profondeur de raisonnement, le contexte de travail, le vocabulaire et les instructions transversales appliqués à toutes les conversations Outcom.
- Son onglet **Comportement** permet d’autoriser un repli d’insertion : si la zone de réponse ne peut plus être remplacée précisément, Outcom conserve le contenu existant et ajoute la proposition tout en haut du brouillon. Le repli ne s’applique jamais à un message envoyé, fermé ou remplacé par un autre message.
- Les profondeurs proposées sont recalculées à partir des capacités du modèle sélectionné ; l'option **Par défaut du modèle** laisse Codex appliquer sa valeur recommandée.
- La zone des documents de contexte accepte le glisser-déposer depuis Windows et Outlook. Les documents sont présentés horizontalement sous forme de tuiles carrées de 82 pixels, avec une icône Windows de 64 pixels. Le nom complet est disponible dans l'infobulle, une croix superposée retire le document et un double-clic l'ouvre dans son application habituelle.
- La dernière réponse peut être transformée en brouillon Outlook en texte brut, sans destinataire et sans envoi automatique.
- La connexion au compte ChatGPT, son état et ses limites d'utilisation peuvent être vérifiés sans envoyer de contenu Outlook.
- Le complément démarre `codex app-server` uniquement lorsqu'une action Outcom en a besoin ; le démarrage d'Outlook reste indépendant du réseau et de Codex.

## Connexion à ChatGPT avec `codex app-server`

Outcom utilise [`codex app-server`](https://developers.openai.com/codex/app-server), l'interface locale qui alimente notamment l'extension Codex de VS Code. L'utilisateur choisit **Se connecter avec ChatGPT** et termine l'authentification dans son navigateur. L'utilisation suit alors l'abonnement, l'espace de travail, les droits et les limites Codex du compte ChatGPT connecté. Elle ne nécessite ni clé API ni facturation API OpenAI séparée. Un abonnement ChatGPT ne garantit toutefois pas un accès illimité : les fonctionnalités et limites Codex du forfait restent applicables.

Cette intégration s'appuie sur un composant local fourni avec Codex, et non sur une API publique appelée directement par le complément. La commande `app-server` est encore présentée comme expérimentale par la version de Codex validée pour ce projet ; son protocole peut donc évoluer. Outcom reste sur la surface stable (`experimentalApi: false`) et refuse de démarrer si les options de sécurité attendues ne sont plus reconnues.

### Première connexion

1. Exécuter `.\tools\Check-DevelopmentEnvironment.ps1` et vérifier que le chemin et la version de Codex sont affichés.
2. Compiler le projet, puis le lancer avec `F5` depuis Visual Studio pour ouvrir Outlook.
3. Dans Outlook, ouvrir l'onglet **Outcom**, puis cliquer sur **Configurer Codex**.
4. Cliquer sur **Se connecter avec ChatGPT**. Le navigateur s'ouvre sur le flux d'authentification officiel.
5. Choisir le compte et, le cas échéant, l'espace de travail ChatGPT à utiliser, puis revenir dans Outlook.
6. Cliquer sur **tester Codex** pour relire le compte, le modèle disponible et les limites d'utilisation. Ce test ne lance aucune génération.

### Définir le Contexte Codex et utiliser le volet

1. Dans l'onglet **Outcom**, cliquer sur **Contexte Codex**, juste à côté de **Configurer Codex**.
2. Choisir le modèle et la profondeur de raisonnement, puis renseigner au besoin les trois onglets : contexte de travail, vocabulaire et instructions transversales. Une modification réinitialise les conversations ouvertes afin de ne jamais mélanger deux configurations dans le même fil.
3. Cliquer sur **Ouvrir Codex**. Un second clic masque le volet sans perdre les conversations en mémoire. Sous le résumé du contexte actif, la zone utile est organisée ainsi : **Activités en cours et récentes** (40 %), liste des **Conversations** (20 %), puis conversation sélectionnée (40 %), dont environ 28 % pour les documents de contexte. La séparation horizontale principale reste ajustable ; **Effacer les tâches terminées** retire uniquement l'historique achevé. La connexion et le contexte se gèrent uniquement depuis le ruban.
4. Saisir une demande puis utiliser **Envoyer** ou `Ctrl+Entrée`. Le bouton **Annuler** interrompt le tour actif.
5. Facultatif : déposer dans la zone de contexte jusqu'à dix fichiers Windows, courriers ou pièces jointes Outlook. La zone affiche les documents horizontalement ; le nom complet apparaît dans l'infobulle, la croix retire un document et un double-clic l'ouvre. Lorsqu'un courrier MSG ou EML est ajouté, ses pièces jointes non incorporées sont séparées et affichées comme des documents de contexte indépendants, dans la limite des dix documents et de 25 Mo par pièce jointe. Outcom accepte les descripteurs Outlook Unicode et ANSI ; si le flux virtuel n'est pas accessible, il enregistre directement la sélection Outlook en MSG Unicode. Au prochain envoi, l'extraction locale s'exécute en arrière-plan : PDF, EML, MSG et documents Office modernes sont convertis en texte structuré pour Codex, tandis que l'original est copié sans modification dans l'espace isolé de la conversation. Les images PNG, JPEG, GIF et WebP sont jointes directement comme images locales.
6. Utiliser **Nouvelle conversation** pour créer un fil indépendant sans effacer les autres. La liste **Activités Codex** permet de passer d'une conversation à une autre, y compris pendant plusieurs réponses simultanées. Chaque conversation conserve son propre brouillon de saisie, son transcript et ses documents de contexte.
7. Après une réponse liée à un courrier, **Créer un brouillon** enregistre et ouvre un nouveau message en texte brut. Aucun destinataire n'est ajouté et Outcom n'appelle jamais l'envoi ; l'utilisateur reste seul responsable de la relecture, des destinataires et de l'envoi manuel.

Le Contexte Codex est conservé dans `%LOCALAPPDATA%\Outcom\Context\global-context.dat`. Son contenu, y compris le modèle et la profondeur de raisonnement, est chiffré avec DPAPI pour le compte Windows courant : il ne peut pas être repris tel quel par un autre utilisateur ou sur un autre poste. Seules les sections de directives non vides sont ajoutées aux instructions au démarrage de chaque nouveau fil ; les protections fixes d'Outcom restent prioritaires.

Chaque fenêtre principale Outlook possède son propre volet et son propre ensemble de conversations éphémères. Fermer une fenêtre détruit son volet et annule toutes ses conversations actives ; fermer Outlook annule tous les traitements en cours. Le volet affiche le markdown comme du texte sélectionnable et ne fournit volontairement à Codex ni terminal, ni accès général au modèle objet Outlook.

Après une annulation, un délai dépassé ou une erreur de transport, Outcom ne réutilise pas silencieusement un fil dont l'état est incertain. Le volet conserve le transcript visible, signale la rupture de contexte et crée un nouveau fil lors du prochain envoi explicite.

### Proposer une réponse depuis Outlook

1. Pour une réponse déjà en cours, utiliser **Proposer une réponse** dans le groupe **Outcom** de l'onglet **Message**. Outlook affiche cet onglet dans la fenêtre de rédaction isolée ou sous **Outils de composition** lorsque la réponse est intégrée au volet de lecture. La commande n'est pas affichée dans l'onglet permanent **Accueil (Courrier)**.
2. Pour partir directement d'un courrier sélectionné, utiliser **Répondre et proposer** ou **Répondre à tous et proposer** dans **Accueil (Courrier) > Outcom** ou dans **Outcom > Réponses assistées**. Outcom crée et ouvre le brouillon correspondant, puis lance la proposition sans bloquer Outlook.
   Les mêmes commandes sont disponibles dans le groupe **Outcom** de l'onglet **Outils de message > Message** lorsqu'un courrier reçu est ouvert dans sa propre fenêtre.
3. Facultatif : rédiger en haut du message une ébauche, des points à reprendre ou des instructions destinées à Codex. Outcom délimite automatiquement cette zone à l'aide de la signature native de l'éditeur Outlook ; à défaut, il s'arrête au début de la section du courrier cité. Cette partie est placée seule dans un bloc d'orientations explicitement reconnu comme une demande de l'utilisateur ; le mode de génération impose d'identifier et d'appliquer chaque consigne actionnable pour produire le message complet, sans recopier dans le courrier final les consignes qui ne sont pas destinées au correspondant. La langue de ces orientations n'impose pas celle de la réponse : celle-ci suit le message le plus récent du fil, sauf consigne linguistique explicite. Lorsqu'une section citée est visible dans le composeur, elle devient systématiquement le message principal auquel répondre, même si l'API Conversation d'Outlook renvoie un autre élément.
   Dans un transfert, Outcom détecte l'action Outlook ou les préfixes usuels `TR`, `FW` et `FWD`, puis transmet les valeurs visibles des champs **À** et **Cc**. Le texte généré s'adresse à ces destinataires et présente le courrier transféré comme information, action ou suite à traiter. Sa langue suit d'abord une consigne explicite, puis celle de l'ébauche et enfin celle de l'interface utilisateur ; la langue du courrier transféré ne l'impose pas automatiquement.
4. L'état de la demande et son éventuelle annulation sont accessibles dans la zone **Activité Codex** du volet Outcom.
5. Si le même brouillon est toujours actif, non envoyé et que la zone de réponse n'a pas changé, le message complet généré remplace cette zone. La signature détectée par Outlook, sa mise en forme et les citations restent en place. Outcom inspecte le début de cette signature : si elle contient déjà une formule de politesse, Codex ne doit pas en ajouter ; sinon, il doit en produire exactement une dans la langue de la réponse. Juste avant l'insertion, Outcom supprime aussi une formule ou une signature finale identique au début de la signature existante. Si la zone a changé et que le repli est activé dans l'onglet **Comportement**, la proposition est ajoutée au début sans effacer les modifications. Outcom n'enregistre ni n'envoie le message automatiquement. Une fenêtre fermée, une réponse intégrée remplacée ou un message envoyé interdit toujours toute insertion tardive.

Cette commande utilise le modèle, la profondeur de raisonnement et les directives du **Contexte Codex** général. Si Outlook ne permet pas d'énumérer la conversation, Codex reçoit uniquement le corps actuel du brouillon, qui contient généralement l'historique cité de la réponse.

L'authentification est isolée de l'installation Codex personnelle de l'utilisateur :

- `CODEX_HOME` est fixé à `%LOCALAPPDATA%\Outcom\CodexHome` ; la configuration, les extensions et les serveurs MCP personnels ne sont donc pas chargés par Outcom.
- Les identifiants sont demandés au magasin sécurisé du système (`keyring`), c'est-à-dire au gestionnaire d'identifiants Windows, et ne sont jamais lus ni journalisés par le complément.
- Les variables d'environnement de clés API et de fournisseurs sont retirées du processus enfant : la liaison ne peut pas basculer silencieusement vers une facturation API séparée.
- Le répertoire de travail est un sous-dossier de session neuf et vide sous `%LOCALAPPDATA%\Outcom\CodexWorkspace` ; il est supprimé à l'arrêt s'il est resté vide.
- L'ancienne configuration `%LOCALAPPDATA%\Outcom\openai.settings.xml` est ignorée. Elle n'est pas supprimée automatiquement afin de ne pas effacer une donnée utilisateur existante ; elle peut être supprimée manuellement si elle n'est plus utile.

### Trouver l'exécutable Codex

Outcom cherche `codex.exe` dans cet ordre :

1. le chemin complet défini par la variable d'environnement `OUTCOM_CODEX_PATH` ;
2. le runtime épinglé placé dans `CodexRuntime\codex.exe` à côté de l'assembly Outcom ;
3. le runtime épinglé `%LOCALAPPDATA%\Outcom\CodexRuntime\codex.exe` ;
4. la commande `codex.exe` disponible dans `PATH` ;
5. le runtime de la dernière extension VS Code ou VS Code Insiders `openai.chatgpt-*` installée.

La version destinée à la distribution est épinglée à **`codex-cli 0.144.6`** pour **Windows x64** dans `Outcom.AddIn/CodexRuntime.json`. Outcom vérifie la version et le SHA-256 du runtime distribué avant de le lancer. `OUTCOM_CODEX_PATH`, le `PATH` et l'extension VS Code restent des solutions de développement explicites, mais ne constituent pas le runtime reproductible du paquet.

### Préparer le runtime de distribution

Le binaire n'est pas stocké dans Git. La commande suivante télécharge l'archive de la release officielle OpenAI, vérifie le SHA-256 publié, extrait le binaire, vérifie une seconde empreinte et sa signature Authenticode OpenAI, contrôle l'architecture PE x64 et la version, puis réalise un handshake `initialize` avec exactement les options `app-server` sécurisées d'Outcom :

```powershell
.\tools\Prepare-CodexRuntime.ps1
```

Le résultat est placé dans `artifacts\CodexRuntime` et contient :

- `codex.exe`, renommé depuis l'artefact officiel Windows x64 ;
- `CodexRuntime.json`, manifeste de provenance et d'intégrité ;
- `codex-runtime.validation.json`, preuve datée du contrôle local ;
- `Codex-CLI-LICENSE.txt`, licence Apache 2.0 du tag distribué.

Pour installer cette même version dans le profil Windows courant :

```powershell
.\tools\Prepare-CodexRuntime.ps1 -InstallForCurrentUser
```

Le dossier cible est alors `%LOCALAPPDATA%\Outcom\CodexRuntime`. Un installateur Outcom doit reprendre les quatre fichiers préparés, conserver la licence Codex CLI et installer `codex.exe` dans ce dossier ou dans un sous-dossier `CodexRuntime` à côté de l'assembly du complément. Les mentions de tiers sont regroupées dans `THIRD-PARTY-NOTICES.md`.

### Périmètre de sécurité actuel

La connexion, sa gestion et son test ne transmettent aucun contenu Outlook. Dans le volet, seul le texte saisi est envoyé par défaut. Un document n'est ajouté qu'après son dépôt explicite dans la zone de contexte. Lors du prochain clic sur **Envoyer**, Outcom extrait localement une représentation textuelle bornée, l'inclut dans la demande, puis copie également le fichier complet dans un sous-dossier isolé propre à la conversation. Les originaux PDF, EML, MSG et Office ne sont donc jamais altérés. L'extraction prend en charge les couches texte PDF, les métadonnées et corps des courriels, DOCX, PPTX, XLSX, HTML et les principaux formats texte ; un PDF numérisé sans couche texte reste signalé comme tel et disponible sous sa forme native. La limite d'extraction est de 60 000 caractères par document et 160 000 caractères par envoi. La limite des fichiers reste de 50 Mo par fichier et 100 Mo par conversation ; la matérialisation initiale des fichiers virtuels fournis par Outlook reste limitée à 25 Mo par fichier. Ces copies sont supprimées au retrait, à la réinitialisation ou à la fermeture de la conversation, ainsi qu'à l'arrêt du client Codex.

Le clic explicite sur **Proposer une réponse**, **Répondre et proposer** ou **Répondre à tous et proposer** transmet l'objet, les valeurs visibles des champs actuels **À** et **Cc**, et la zone de réponse située en haut du brouillon en texte brut, limités respectivement à 12 000 caractères par liste de destinataires et à 50 000 caractères pour la zone de réponse. Le champ **Cci** n'est pas transmis. Outcom joint également le message source visible ou jusqu'à 20 messages récents du fil accessibles à Outlook, avec un maximum cumulé de 80 000 caractères. La limite de cette zone provient du signet de signature de l'éditeur Word d'Outlook ou, à défaut, de l'en-tête de la section du courrier cité ; aucun nom de personne n'est recherché pour cette délimitation. La détection accepte les espaces ordinaires ou insécables ainsi que les en-têtes Outlook français, anglais, portugais, espagnols, allemands et italiens. La section citée visible est prioritaire et transmise comme message principal en texte brut. Seule la zone de réponse est marquée comme instruction explicite de l'utilisateur ; les destinataires, l'objet et le contenu du fil restent des données non fiables mais constituent des sources documentaires obligatoires qui déterminent le public et le fond sans pouvoir neutraliser les orientations. Pour les éléments structurés issus de l'API Conversation, seuls le nom affiché de l'expéditeur, la date et le corps en texte brut sont copiés. Les pièces jointes, en-têtes techniques et HTML restent exclus. La génération est asynchrone : Outcom capture la zone directement dans l'éditeur Word, surveille l'envoi et, pour une fenêtre isolée, sa fermeture, puis vérifie de nouveau l'état envoyé, l'identité du brouillon actif et le texte visible de cette même zone juste avant le remplacement. Un message envoyé, fermé ou remplacé n'est jamais altéré par une réponse tardive. Si la zone de réponse a réellement changé, Outcom abandonne l'insertion ou, lorsque l'option de repli est active, ajoute la proposition au début sans écraser le contenu existant. Le journal indique uniquement si des orientations et un message source ont été détectés, sans enregistrer leur texte. Si aucun des deux n'est disponible, Outcom refuse de lancer une réponse hors contexte.

Le contenu des documents est traité comme une donnée non fiable et non comme une instruction. Ajouter ou retirer un document déjà transmis réinitialise la conversation après confirmation, afin de ne pas mélanger silencieusement plusieurs contextes. Les conversations sont éphémères et conservées uniquement en mémoire par le processus local ; elles ne sont pas reprises automatiquement après un redémarrage de `app-server`.

Les sessions préparées pour les futures générations sont éphémères et utilisent un bac à sable en lecture seule, un répertoire vide, aucune approbation automatique et aucun accès réseau depuis les outils. Les outils shell, modification de fichiers, navigateur, applications, plugins, MCP, génération d'images et sous-agents sont désactivés ; toute demande d'exécution, de permission ou d'outil reçue du serveur est refusée. `codex app-server` conserve naturellement son propre accès HTTPS à OpenAI pour l'authentification et le service de modèle.

Le journal local ne contient ni URL d'authentification, ni jeton, ni requête JSON, ni demande utilisateur, ni réponse Codex, ni contenu Outlook, ni texte du Contexte Codex. Il reste limité aux événements techniques nécessaires au diagnostic.

La création d'un brouillon est une action Outlook locale distincte de la génération. Elle produit un message en texte brut, sans destinataire, l'enregistre et l'affiche, mais ne l'envoie jamais. Codex ne reçoit aucun outil lui permettant d'appeler cette action lui-même.

## Compiler et exécuter

1. Ouvrir `Outcom.AddIn\Outcom.AddIn.slnx` dans Visual Studio.
2. Sélectionner la configuration `Debug` et **Générer > Générer la solution**.
3. Lancer avec `F5` : Visual Studio inscrit temporairement le complément et démarre Outlook classique.
4. Vérifier l'onglet **Outcom**, effectuer la connexion navigateur, puis ouvrir le volet Codex et envoyer une demande sans courrier.
5. Sélectionner ensuite un courrier de test, le joindre explicitement, vérifier l'indicateur affiché, puis valider la création d'un brouillon sans destinataire.

La présence de Codex n'est pas nécessaire pour compiler le complément. Si le diagnostic affiche un avertissement Codex, installer Codex CLI, l'extension Codex pour VS Code, ou renseigner `OUTCOM_CODEX_PATH` avant de tester la connexion dans Outlook.

### Certificat de signature du manifeste

Le projet VSTO signe ses manifestes. Le fichier PFX de développement est volontairement ignoré par Git : une clé privée et son mot de passe ne doivent jamais être publiés dans le dépôt. Sur un nouveau poste ou dans une CI, il faut donc soit créer un certificat de test depuis **Propriétés du projet > Signature**, soit fournir un certificat injecté par le système de secrets.

Les valeurs locales du projet peuvent être remplacées sans modifier le fichier `.csproj` avec `OUTCOM_MANIFEST_KEY_FILE` (chemin du PFX) et `OUTCOM_MANIFEST_CERTIFICATE_THUMBPRINT` (empreinte du certificat), ou avec les propriétés MSBuild `ManifestKeyFile` et `ManifestCertificateThumbprint`. Une version distribuée doit utiliser un certificat de signature de code géré par l'organisation ; le certificat temporaire sert uniquement au développement.

## Structure du dépôt

```text
Outcom.AddIn/
  Outcom.AddIn.slnx
  Outcom.AddIn.csproj
  Assets/Outcom.ico
  Codex*.cs
  DocumentContextExtractor.cs
  OutcomRibbon.cs
  OutcomRibbon.xml
  ThisAddIn.cs
tools/
  Check-DevelopmentEnvironment.ps1
```

Les fichiers `ThisAddIn.Designer.cs` et `ThisAddIn.Designer.xml` sont générés par VSTO et ne doivent pas être modifiés manuellement.

## Étapes suivantes conseillées

- [x] Ajouter une journalisation locale minimale du démarrage et de l'arrêt.
- [x] Ajouter un bouton Ribbon sans traitement métier.
- [x] Ajouter une connexion ChatGPT configurable et testable via `codex app-server`.
- [x] Ajouter un volet Codex multi-tour dans chaque fenêtre principale Outlook.
- [x] Limiter la lecture Outlook au clic explicite et aux champs de courrier documentés.
- [x] Ajouter la création locale d'un brouillon sans destinataire ni envoi automatique.
- [x] Épingler et valider le runtime Codex destiné à la distribution.
- [ ] Tester les cas où Outlook désactive un complément lent.
- [ ] Préparer un MSI signé, distinct si un support Office 32 bits devient nécessaire.

## Documentation Microsoft

- [Configurer un ordinateur pour développer des solutions Office](https://learn.microsoft.com/visualstudio/vsto/how-to-configure-a-computer-to-develop-office-solutions)
- [Bien démarrer avec les compléments VSTO](https://learn.microsoft.com/visualstudio/vsto/getting-started-programming-vsto-add-ins)
- [Architecture des compléments VSTO](https://learn.microsoft.com/visualstudio/vsto/architecture-of-vsto-add-ins)
- [Déployer une solution VSTO avec Windows Installer](https://learn.microsoft.com/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)

## Documentation OpenAI

- [Codex App Server](https://developers.openai.com/codex/app-server)
- [Authentification Codex](https://developers.openai.com/codex/auth)
- [Codex CLI](https://developers.openai.com/codex/cli)
- [Sécurité et approbations Codex](https://developers.openai.com/codex/security)
