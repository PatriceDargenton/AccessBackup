
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

    ' Pour pouvoir localiser la ligne ayant provoqu� une erreur, mettre bTrapErr = False
    ' Intercepter les erreurs (peu utilis�)
    Public Const bTrapErr As Boolean = True ' Mode Release = distribuable
    'Public Const bTrapErr As Boolean = False ' Ne pas intercepter les erreurs

    ' Afficher toutes les bo�tes de dialogue des erreurs
    Public Const bPromptErrGlob As Boolean = False

    ' Ignorer les dates en mode debug : 
    '  sauvegarder m�me si la date d'�criture n'a pas chang�e
    Public Const bIgnorerDateMAJ As Boolean = False
    ' Tester le backup toutes les 10 secondes
    Public Const bBoucleInfinie As Boolean = False

    Public Const iDelaiLectureMsgMilliSec% = 5000

    ' Suffixe pour la copie temporaire (autre que base Access)
    Public Const sSuffixeCopieTmp$ = "_CopieTmp"
    ' Suffixe pour la base compact�e temporaire
    Public Const sSuffixeCompactTmp$ = "_CompactTmp"


    ' Valeurs par d�faut :

    ' Suffixe pour la base ouverte non compact�e (utile pour conserver sa date d'�criture)
    Public Const sSuffixeBdOuverteDef$ = "_CopieIncertOrig"
    ' Suffixe pour la base ouverte compact�e
    Public Const sSuffixeBdOuverteCompacteeDef$ = "_CopieIncert" '"CopieNonSure"
    ' Suffixe pour la base ferm�e (et compact�e)
    Public Const sSuffixeBdFermeeDef$ = "_CopieFiable"
    ' Suffixe pour la copie (autre que base Access)
    Public Const sSuffixeCopieDef$ = "_Copie"

    ' Chemin du fichier source � sauvegarder 
    '  s'il n'y a pas d'argument dans la ligne de commande
    Public Const sCheminSrcDef$ = ""

    ' Sous-dossier o� placer les sauvegardes (relatif � l'emplacement
    '  de la source � sauvegarder) ou sinon chemin complet du dossier 
    '  des sauvegardes. On peut aussi laisser vide
    Public Const sDossierSauvegardesDef$ = "\Sauvegardes"

    ' Chemin du fichier de conservation des traces d'ex�cution, le cas �ch�ant 
    ' (laisser vide pour ne pas activer le tra�age). Utile lorsque AccessBackup 
    ' est lanc� depuis une t�che planifi�e sans ouverture de session
    Public Const sCheminTraceDef$ = ""

    ' Nombre de versions distinctes r�centes conserv�es 
    ' (par exemple les 5 derni�res versions)
    Public Const iNbVersionsRoulementDef% = 5
    ' Format pour conserver au moins 9 versions pr�c�dentes
    'Public Const sFormatVersionsRoulementDef$ = "0"
    Public Const sFormatVersionsRoulementDef$ = "00" ' 99 versions pr�c�dentes

    ' Suffixe pour les fichiers d'archive d�finitive
    Public Const sSuffixeArchiveDef$ = "Arch"
    ' P�riode d'archivage : Nombre de jours d'intervalle
    '  pour faire une nouvelle archive d�finitive
    Public Const iPeriodeArchJoursDef% = 7
    ' Format pour conserver au moins 999 versions d'archives 
    ' (en cas de d�passement, le tri des fichiers ne sera pas parfait)
    Public Const sFormatVersionsArchDef$ = "000"

End Module