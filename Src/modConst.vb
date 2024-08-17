
' Fichier modConst.vb
' -------------------

Module modConst

    Public ReadOnly sNomAppli$ = My.Application.Info.Title ' AccessBackup
    Public Const sTitreMsg$ = "AccessBackup - Gestionnaire de sauvegarde"
    Public Const sDateVersionAppli$ = "17/08/2024"

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

    ' Pour pouvoir localiser la ligne ayant provoqué une erreur, mettre bTrapErr = False
    ' Intercepter les erreurs (peu utilisé)
    Public Const bTrapErr As Boolean = True ' Mode Release = distribuable
    'Public Const bTrapErr As Boolean = False ' Ne pas intercepter les erreurs

    ' Afficher toutes les boîtes de dialogue des erreurs
    Public Const bPromptErrGlob As Boolean = False

    ' Ignorer les dates en mode debug : 
    '  sauvegarder même si la date d'écriture n'a pas changée
    Public Const bIgnorerDateMAJ As Boolean = False
    ' Tester le backup toutes les 10 secondes
    Public Const bBoucleInfinie As Boolean = False

    Public Const iDelaiLectureMsgMilliSec% = 5000

    ' Suffixe pour la copie temporaire (autre que base Access)
    Public Const sSuffixeCopieTmp$ = "_CopieTmp"
    ' Suffixe pour la base compactée temporaire
    Public Const sSuffixeCompactTmp$ = "_CompactTmp"


    ' Valeurs par défaut :

    ' Suffixe pour la base ouverte non compactée (utile pour conserver sa date d'écriture)
    Public Const sSuffixeBdOuverteDef$ = "_CopieIncertOrig"
    ' Suffixe pour la base ouverte compactée
    Public Const sSuffixeBdOuverteCompacteeDef$ = "_CopieIncert" '"CopieNonSure"
    ' Suffixe pour la base fermée (et compactée)
    Public Const sSuffixeBdFermeeDef$ = "_CopieFiable"
    ' Suffixe pour la copie (autre que base Access)
    Public Const sSuffixeCopieDef$ = "_Copie"

    ' Chemin du fichier source à sauvegarder 
    '  s'il n'y a pas d'argument dans la ligne de commande
    Public Const sCheminSrcDef$ = ""

    ' Sous-dossier où placer les sauvegardes (relatif à l'emplacement
    '  de la source à sauvegarder) ou sinon chemin complet du dossier 
    '  des sauvegardes. On peut aussi laisser vide
    Public Const sDossierSauvegardesDef$ = "\Sauvegardes"

    ' Chemin du fichier de conservation des traces d'exécution, le cas échéant 
    ' (laisser vide pour ne pas activer le traçage). Utile lorsque AccessBackup 
    ' est lancé depuis une tâche planifiée sans ouverture de session
    Public Const sCheminTraceDef$ = ""

    ' Nombre de versions distinctes récentes conservées 
    ' (par exemple les 5 dernières versions)
    Public Const iNbVersionsRoulementDef% = 5
    ' Format pour conserver au moins 9 versions précédentes
    'Public Const sFormatVersionsRoulementDef$ = "0"
    Public Const sFormatVersionsRoulementDef$ = "00" ' 99 versions précédentes

    ' Suffixe pour les fichiers d'archive définitive
    Public Const sSuffixeArchiveDef$ = "Arch"
    ' Période d'archivage : Nombre de jours d'intervalle
    '  pour faire une nouvelle archive définitive
    Public Const iPeriodeArchJoursDef% = 7
    ' Format pour conserver au moins 999 versions d'archives 
    ' (en cas de dépassement, le tri des fichiers ne sera pas parfait)
    Public Const sFormatVersionsArchDef$ = "000"

End Module