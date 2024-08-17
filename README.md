# AccessBackup
Un gestionnaire de sauvegarde de base de données Access (ou autre fichier)
---

Comment faire régulièrement des copies de sauvegarde d'une base MS-Access partagée sur un lecteur réseau ? il suffit de créer une tâche planifiée de copie sur le serveur, c'est tout ! mais que se passe-t-il si la base est en cours d'utilisation ? comment conserver les versions successives des copies ? AccessBackUp adresse précisément ce genre de tâche, et il peut archiver en fait tout type de fichier. De plus, c'est un programme DotNet qui tourne tel quel sur un Windows Serveur sans installation particulière à faire, pas d'impact sur vôtre configuration donc. Voici la liste des fonctionnalités assurées par AccessBackUp :
- Gestion de deux niveaux de sécurité :
   1°) Niveau fréquent (par exemple toutes les heures) : copie instantanée temporaire de la base MS-Access ;
   2°) Niveau + rare (par exemple une fois par jour) : copie instantanée en vue d'archivage définitif de la base MS-Access ;
- Gestion des bases en cours d'utilisation (copie incertaine ou sinon copie fiable si personne n'est connecté) ;
- Gestion du mot de passe des bases MS-Access ;
- Gestion du compactage (avec réparation, le cas échéant) des bases MS-Access (DAO.DBEngine.CompactDatabase fonctionne en DotNet standard via l'adaptateur DAO) ;
- Gestion de la compression des sauvegardes au format standard .zip ;
- Gestion des roulements (pour conserver par exemple les 10 dernières versions) ;
- Gestion de la numérotation des archives définitives ;
- Paramétrage complet en ligne de commande ;
- Fichier log pour laisser la trace des sauvegardes effectuées ;
- Et bien sûr il fonctionne sans session ouverte sur le serveur.

Mots clés : MS-Access, DotNet, VB .Net, Sauvegardes, Compactage, Compression, Zip, Versions, Roulement, Archives, Snapshot, Sauvegarde en continue, Sauvegarde à chaud, CDP : Continuous data protection : sauvegarde de données en continue, Hot backup.

## Table des matières
- [Paramètres](#paramètres)
- [Limitations](#limitations)
- [Versions](#versions)
- [Liens](#liens)

## Paramètres
Syntaxe : Paramètre Valeur : séparés par un espace, ainsi que les autres paramètres.
(s'il n'y a qu'un argument, alors c'est le chemin de la base à sauvegarder, c'est le seul paramètre obligatoire)

Résumé :
- CheminSrc
- DossierSauvegardes
- DossierSauvegardesIncert
- SuffixeArchive
- SuffixeCopie
- SuffixeBdOuverte
- SuffixeBdOuverteCompactee
- SuffixeBdFermee
- FormatVersionsRoulement
- FormatVersionsArch
- PeriodeArchJours
- NbVersionsRoulement
- CheminTrace
- MotDePasse
- CompactRepair

- CheminSrc : Chemin du fichier source à sauvegarder, par exemple :
	C:\Tmp\MaBase.mdb ou
	\\Serveur\Donnees\MaBasePartagee.mdb
	"\\Serveur\Mon dossier\Ma base partagée.mdb"
	\\Serveur\Donnees\MonDocument.doc

- DossierSauvegardes : Sous-dossier où placer les sauvegardes (relatif à l'emplacement de la source), ou sinon chemin complet du dossier des sauvegardes. On peut aussi laisser vide. Exemples :
	\Bak (ou Bak)
	C:\Sauvegardes
	\\Serveur\Sauvegardes
	"\\Serveur\Mes sauvegardes"

- DossierSauvegardesIncert : On peut choisir de séparer les sauvegardes incertaines dans un dossier distinct, sinon le même dossier sera utilisé
- SuffixeArchive : Suffixe des fichiers d'archive à conserver, par exemple : Arch
- SuffixeCopie : Suffixe pour la copie (autre que base MS-Access), par exemple : _Copie
- SuffixeBdOuverte : Suffixe pour la base ouverte non compactée (utile pour conserver sa date d'écriture), par exemple : Arch
- SuffixeBdOuverteCompactee : Suffixe pour la base ouverte compactée, par exemple : _CopieIncert ou _CopieNonSure
- SuffixeBdFermee : Suffixe pour la base fermée (et compactée), par exemple : _CopieFiable
- FormatVersionsRoulement : Format numérique pour conserver les n versions précédentes, par exemple
	0 pour conserver au maximum les 9 versions précédentes
	00 pour conserver au maximum les 99 versions précédentes
- FormatVersionsArch : Format numérique pour conserver les n versions d'archives définitives
	000 pour conserver au moins 999 versions (en cas de dépassement, le tri des fichiers ne sera pas parfait, mais cela ne pose pas d'autre problème)
- PeriodeArchJours : Période d'archivage : Nombre de jours d'intervalle pour faire une nouvelle archive définitive
- NbVersionsRoulement : Nombre de versions distinctes récentes conservées (par exemple les 5 dernières versions)
- CheminTrace : Chemin du fichier de conservation des traces d'exécution, le cas échéant (laisser vide pour ne pas activer le traçage). On peut aussi indiquer seulement un sous-dossier où se trouve AccessBackup. Utile lorsque AccessBackup est lancé depuis une tâche planifiée sans ouverture de session. Exemples :
	Trace.txt
	\Trace\Trace.txt
	C:\Tmp\Log.txt

- MotDePasse : Mot de passe de la base MS-Access : Si la base n'a pas de mot de passe, cela fonctionnera aussi, mais attention car la base compactée elle sera protégée !

- CompactRepair : pour pouvoir réparer une base corrompue sans requérir une version complète de MS-Access (syntaxe : CompactRepair C:\MaBd.mdb, CompactRepair remplace CheminSrc dans ce cas). Cet argument doit être utilisé si AccessBackup indique que la base est ouverte en mode exclusif ("Base ouverte exclusivement"), et que la base est inaccessible (depuis une application par exemple). MS-Access affiche le message suivant lorsque l'on tente d'ouvrir la base : "La base de données ‘xxx.mdb’ doit être réparée ou n'est pas un fichier de base de données. Vous ou un autre utilisateur avez peut-être quitté Microsoft Access de manière inattendue alors qu'une base de données Microsoft Access était ouverte. Essayer de réparer la base de données ? Oui Non".

## Limitations
Au moment de faire la copie directe du fichier de base de données, le fichier est verrouillé, pas au sens d'une base de données, mais au sens de l'accès disque à un fichier quelconque : à ce moment là, une tentative d'écriture, de verrouillage ou d'ouverture en mode exclusif (par exemple pour effectuer un compactage) peut échouer, au quel cas il faudra prévoir un traitement d'erreur : de toute façon, toute application de base de données doit gérer les erreurs, quelque soit le type d'accès demandé, ce n'est pas vraiment une exigence particulière qui est recommandée ici.

Si on ouvre une base MS-Access en mode exclusif (option de la boite de dialogue d'ouverture d'une base sous MS-Access), alors AccessBackup ne peut pas la compacter ; mais, si on ouvre cette fois la base en mode lecture seule et exclusif, cette fois AccessBackup peut la compacter. Cependant, il ne peut pas terminer le processus car on ne peut pas remplacer la base d'origine par la copie compactée (mais on peut faire une copie fiable de la base, contrairement au mode exclusif tout court dans lequel on ne peut faire aucune copie).

Dans tous les cas, aucune boite de dialogue n'est affichée sur le serveur, car le programme doit fonctionner sans qu'une session soit ouverte : si une erreur survient, on peut consulter la trace des messages dans le fichier dédié à cela, si le traçage est activé. Sinon, si une session est ouverte, un minuteur permet d'avoir le temps de lire les messages.

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

Documentation d'origine complète : [AccessBackup.html](http://patrice.dargenton.free.fr/CodesSources/AccessBackup.html)