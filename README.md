# ffvl_club.ps1

Script pour faire des statistiques sur les licenciés/adhérents d'un club. Notamment dans un objectif sécurité.

Note : le site de la ffvl fourni déjà des stats intéressantes, via les sous-onglets "Statistiques" et "Licenciés qualifiés", mais ces sections souffrent de manque de cohérence rendant les calculs complexes. Voir dernier chapitre pour plus d'infos.

Ce script permet d'obtenir :
- La liste des pilotes par groupe de meilleure qualification obtenue
- La liste des pilotes sans brevet, trié par première année de license ffvl
- Le nombre de license par type
- Le nombre de packs Individuelle Accident (IA) contractés
- Le nombre de pilotes sans pack IA
- La liste des biplaceurs associatif sans IA passager
- Le nombre d'options choisies pour :
  - Assurance matérielle
  - Carte compétiteur
  - Protection juridique
  - Extension sports de pleine nature
- La liste des : accompagnateurs, animateurs, juges de précision d'atterrissage
- La liste des qualifications obtenues dans l'année, groupé par qualification
- La liste des pilotes primo-pratiquants ayant obtenu un brevet

Il est possible d'afficher les résultats uniquement pour les licenciés principaux du club (défaut), ou bien pour tous les adhérents dont ceux déjà inscrit à un autre club ffvl (`-addOtherMembers $true`).

## Utilisation

### 0 - Pré-requis
* Avoir un accès à la liste des licenciés du club sur l'intranet ffvl (membre du bureau)
* Fonctionne avec Powershell 5 ou Powershell core
* Installer le module powershell PowerHTML :  
`Install-Module PowerHTML -Scope CurrentUser -ErrorAction Stop`  
Ce module permet de rechercher du contenu via requête XPath dans les pages HTML téléchargées.

### 1 - Récupérer un cookie d'authentification

- Connectez-vous avec votre navigateur favori sur le site officiel : https://intranet.ffvl.fr/
- Récupérer les informations du cookie d'authentification :
  - Typiquement: F12 > Application (ou Stockage sous Firefox) > Cookies > intranet.ffvl.fr
  - Rechercher le cookie dont le nom commence par `SSESS`
  - Copier son nom et valeur, qui seront à spécifier en paramétre du script (cf section suivante)

Note #1 : ne pas communiquer ces infos à autrui ! (C'est comme si vous donniez votre login/password)  
Note #2 : la date d'expiration du cookie est visible dans la colonne appropriée, il est donc valable plusieurs jours  

Il vous faudra aussi l'identifiant de la structure. Très simple : allez sur "Ma structure" puis regarder la dernière partie de l'url.

### 2 - Lancer le script
Dans un terminal powershell :
``` powershell
$year = 2024 # Année à analyser
$structure = 377 # Identifiant club, ce n'est PAS le numéro du club, mais l'identifiant dans l'url
$cookieSessionName = "SSESS<hash>" # Nom du cookie de session
$cookieSessionValue = "xxxxxxxxxxxxxxxxxx" # Valeur du cookie de session
.\ffvl_club.ps1 -year $year -structure $structure -cookieSessionName $cookieSessionName -cookieSessionValue $cookieSessionValue
```

Paramètres possibles :
* `-structure <int>`: Identifiant club
* `-year <int>`: Année à analyser ; par défaut l'année en cours si non spécifié
* `-cookieSessionName <string>`: Nom du cookie de session
* `-cookieSessionValue <string>`: Valeur du cookie de session
* `-filterBestQualification <bool>`: Afficher uniquement la meilleure qualification (défaut $true) ; à $false le pilote apparaît dans chaque qualification validée
* `-addOtherMembers <bool>`: Prendre en compte les adhérents du club qui sont déjà licenciés dans un autre club (défaut $false)

## En cas de problème

Me contacter ici (issues) ou via : thibault.rohmer@gmail.com

## Fonctionnement

Le script récupère la liste des licenciés, comme quand on va manuellement sur l'onglet "Licenciés" de l'intranet ffvl.  
Note: aucune modification n'est effectuée sur le site ffvl.
Puis, le script va récupérer la fiche de chaque pilote pour y extraire les informations utiles.  
Enfin, le script calcule les groupes et affiche les résultats.

## Limitations actuelles du site ffvl
- Certains pilotes ont validé les partie pratique et théorique d'un brevet mais n'apparaissent pas dans la bonne catégorie de l'onglet "licenciés qualifiés"
- Dans "licenciés qualifiés", des pilotes apparaissent alors qu'ils n'ont pas fini de prendre leur licence (et ne font donc pas parti des licenciés effectifs)
- Dans "licenciés qualifiés", un pilote ayant bpi + bp + bpc + biplace apparaît dans chaque section; pas de vue synthétique pour le club
- Mélange licenciés / adhérents dans certains menus et pas d'autres
