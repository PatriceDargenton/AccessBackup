
' Fichier frmAccessBackup.vb
' --------------------------

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (bool�en vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacit� du VB6.Long)
' l pour Long : &
' r pour nombre R�el (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' a pour Array (tableau) : ()
' o pour Object : objet instanci� localement
' refX pour reference � un objet X pr�existant qui n'est pas sens� �tre ferm�
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Public Class frmAccessBackup

    Private m_sCheminTrace$, m_sCheminSrc$
    Private m_sDossierSauvegardes$, m_sDossierSauvegardesIncert$
    Private m_sFormatVersionsRoulement$, m_sFormatVersionsArch$
    Private m_iPeriodeArchJours%, m_iNbVersionsRoulement%

    Private m_sSuffixeArchive$, m_sSuffixeCopie$
    Private m_sSuffixeBdOuverte$, m_sSuffixeBdOuverteCompactee$
    Private m_sSuffixeBdFermee$
    Private m_sMotDePasse$ = "" ' 01/11/2009
    Private m_bCompactRepair As Boolean ' 16/03/2013

#Region " Propri�t�s de la classe "

    Public WriteOnly Property sCheminSrc$()
        'Get
        '    sCheminSrc = m_sCheminSrc
        'End Get
        Set(sCheminSrc0$)
            m_sCheminSrc = sCheminSrc0
        End Set
    End Property

    Public WriteOnly Property sCheminTrace$()
        Set(sCheminTrace0$)
            m_sCheminTrace = sCheminTrace0
        End Set
    End Property

    Public WriteOnly Property sDossierSauvegardes$()
        Set(sDossierSauvegardes0$)
            m_sDossierSauvegardes = sDossierSauvegardes0
        End Set
    End Property

    Public WriteOnly Property sDossierSauvegardesIncert$()
        Set(sDossierSauvegardesIncert0$)
            m_sDossierSauvegardesIncert = sDossierSauvegardesIncert0
        End Set
    End Property

    Public WriteOnly Property sSuffixeArchive$()
        Set(sSuffixeArchive0$)
            m_sSuffixeArchive = sSuffixeArchive0
        End Set
    End Property

    Public WriteOnly Property sSuffixeCopie$()
        Set(sSuffixeCopie0$)
            m_sSuffixeCopie = sSuffixeCopie0
        End Set
    End Property

    Public WriteOnly Property sSuffixeBdOuverte$()
        Set(sSuffixeBdOuverte0$)
            m_sSuffixeBdOuverte = sSuffixeBdOuverte0
        End Set
    End Property

    Public WriteOnly Property sSuffixeBdOuverteCompactee$()
        Set(sSuffixeBdOuverteCompactee0$)
            m_sSuffixeBdOuverteCompactee = sSuffixeBdOuverteCompactee0
        End Set
    End Property

    Public WriteOnly Property sSuffixeBdFermee$()
        Set(sSuffixeBdFermee0$)
            m_sSuffixeBdFermee = sSuffixeBdFermee0
        End Set
    End Property

    Public WriteOnly Property sFormatVersionsRoulement$()
        Set(sFormatVersionsRoulement0$)
            m_sFormatVersionsRoulement = sFormatVersionsRoulement0
        End Set
    End Property
    Public WriteOnly Property sFormatVersionsArch$()
        Set(sFormatVersionsArch0$)
            m_sFormatVersionsArch = sFormatVersionsArch0
        End Set
    End Property

    Public WriteOnly Property iPeriodeArchJours%()
        Set(iPeriodeArchJours0%)
            m_iPeriodeArchJours = iPeriodeArchJours0
        End Set
    End Property

    Public WriteOnly Property iNbVersionsRoulement%()
        Set(iNbVersionsRoulement0%)
            m_iNbVersionsRoulement = iNbVersionsRoulement0
        End Set
    End Property

    Public WriteOnly Property sMotDePasse$() ' 01/11/2009
        Set(sMotDePasse0$)
            m_sMotDePasse = sMotDePasse0
        End Set
    End Property

    Public WriteOnly Property bCompactRepair() As Boolean ' 16/03/2013
        Set(bCompactRepair0 As Boolean)
            m_bCompactRepair = bCompactRepair0
        End Set
    End Property

#End Region

    Private Sub frmAccessBackup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 16/03/2013
        Dim sVersionAppli$ = My.Application.Info.Version.Major &
            "." & My.Application.Info.Version.Minor &
            My.Application.Info.Version.Build
        Dim sTxt$ = sTitreMsg & " - Version " & sVersionAppli & " (" & sDateVersionAppli & ")"
        If bDebug Then sTxt &= " - Debug"
        Me.Text = sTxt

        ' Attention, lorsqu'AccessBackUp est lanc� depuis le gestionnaire 
        '  de t�ches de Windows, il n'a pas le focus, ne pas utiliser 
        '  frmActivated, mais bien frmLoad

        ' Attendre une demi-seconde avant de lancer le backup 
        '  pour laisser le temps d'afficher le frm
        Me.timerDebut.Interval = 500
        Me.timerDebut.Start()

    End Sub

    Private Sub timerDebut_Tick(sender As Object, e As EventArgs) Handles timerDebut.Tick

        Me.timerDebut.Stop()
        Backup(m_sCheminSrc)

        If Not bBoucleInfinie Then Exit Sub
        ' Tester le backup toutes les 10 secondes
        Me.timerDebut.Interval = 10000 ' en millisec., soit 10 sec.
        Me.timerDebut.Start()

    End Sub

    Private Sub timerFin_Tick(sender As Object, e As EventArgs) Handles timerFin.Tick
        Me.Close() : Exit Sub
    End Sub

    Private Sub Backup(sCheminSrc$)

        ' Gestion des sauvegardes

        Static bEnCours As Boolean
        ' Eviter la r�entrance dans la fonction (en cas de debug avec le timer)
        If bEnCours Then Exit Sub
        bEnCours = True

        Dim sMsgErr$ = ""
        Dim sMsgResultat$ = ""
        Dim bSucces As Boolean = False

        Sablier()
        Application.DoEvents()
        If sCheminSrc = "" Then
            sMsgResultat = "Rien � faire !"
            AfficherMsg(sMsgResultat)
            bSucces = True
            GoTo Fin
        End If
        If Not bFichierExiste2(sCheminSrc, sMsgErr) Then GoTo Fin

        Dim sCheminCourant$ = IO.Path.GetDirectoryName(sCheminSrc)
        Dim sCheminDossierCopie$ =
            sDeduireChemin(m_sDossierSauvegardes, sCheminCourant)
        Dim sCheminDossierCopieIncert$ =
            sDeduireChemin(m_sDossierSauvegardesIncert, sCheminCourant)
        Dim sNomFichier$ = IO.Path.GetFileNameWithoutExtension(sCheminSrc)
        Dim sExt$ = IO.Path.GetExtension(sCheminSrc) ' Extension du fichier � archiver

        ' Traitement des bases Access
        Dim sTypeCopie$ = m_sSuffixeCopie
        Dim bBaseAccess As Boolean = False
        If sExt.ToLower = ".mdb" Or sExt.ToLower = ".mde" Then bBaseAccess = True
        Dim bBaseAccessDejaCompactee As Boolean = False
        Dim bBaseAccessFermee As Boolean = False
        Dim bBaseOuverteExclusivement As Boolean = False
        Dim bBaseOuverteExclusivLectureSeule As Boolean = False
        Dim iNbU% = 0
        If bBaseAccess Then

            ' Pour une r�paration, il ne faut pas v�rifier les utilisateurs connect�s
            '  car iNbUtilisateurs renvoie bBaseOuverteExclusivement avec l'erreur -2147467259 
            If m_bCompactRepair Then ' 16/03/2013
                AfficherMsg("Compactage/r�paration de la base en cours... :" & vbCrLf & sCheminSrc)
                If Not bCompacterMdb(sCheminSrc, sMsgErr, , , m_sMotDePasse) Then GoTo Fin
                sMsgResultat = "Compactage/r�paration de la base effectu� avec succ�s : " & sCheminSrc
                AfficherMsg("Compactage/r�paration de la base termin� :" & vbCrLf & sCheminSrc)
                bSucces = True : GoTo Fin
            End If

            Dim sInfoUtilisateurs$ = ""
            Dim bBaseFiable As Boolean
            ' Nombre d'utilisateurs en cours de la base Access
            iNbU = iNbUtilisateurs(sCheminSrc, sMsgErr, , sInfoUtilisateurs,
                bBaseOuverteExclusivement, bBaseOuverteExclusivLectureSeule,
                bBaseFiable, , m_sMotDePasse)
            'Dim sMsg$ = "Base MDB : " & vbCrLf & sCheminSrc & vbCrLf & _
            '    sInfoUtilisateurs & vbCrLf & _
            '    "Base fiable : " & CStr(IIf(bBaseFiable, "Oui", "Non"))
            'MsgBox(sMsg, vbInformation, sTitreMsg)

            ' Si la base est ouverte en mode exclusif, on ne peut pas 
            '  faire de sauvegarde maintenant
            If bBaseOuverteExclusivement Then
                sMsgResultat = "Base ouverte exclusivement :   " & sCheminSrc
                AfficherMsg("Base ouverte exclusivement :" & vbCrLf & sCheminSrc)
                ' Succ�s dans le sens que la proc�dure se d�roule sans erreur
                bSucces = True
                GoTo Fin
            End If

            ' La copie directe d'une base Access n'est consid�r�e comme fiable que quand
            '  il n'y a personne de connect� dessus
            sTypeCopie = m_sSuffixeBdOuverte
            If iNbU = -1 Then ' 01/11/2009
                sMsgResultat = "Erreur"
                GoTo Fin
            ElseIf iNbU = 0 Then
                sTypeCopie = m_sSuffixeBdFermee
                bBaseAccessFermee = True
            Else
                sCheminDossierCopie = sCheminDossierCopieIncert
            End If

            ' Si la base est dans un �tat interm�diaire (en cours d'�criture ?) 
            '  ne pas faire de sauvegarde maintenant
            If Not bBaseFiable Then
                sMsgResultat = "Base suspecte :                " & sCheminSrc
                AfficherMsg("Base suspecte :" & vbCrLf & sCheminSrc)
                ' Succ�s dans le sens que la proc�dure se d�roule sans erreur
                bSucces = True
                GoTo Fin
            End If
        End If

        Dim sCheminCopie$ = sCheminDossierCopie & "\" & sNomFichier & sTypeCopie & sExt

        ' Si la copie du fichier existe, v�rifier les derni�res dates d'�criture
        AfficherMsg("V�rification des derni�res dates d'�criture :" & vbCrLf & sCheminSrc)
        Dim dDateSrc As Date = IO.File.GetLastWriteTime(sCheminSrc)
        If bFichierExiste2(sCheminCopie, sMsgErr) Then
            Dim dDateCopie As Date = IO.File.GetLastWriteTime(sCheminCopie)
            If dDateSrc <= dDateCopie Then
                sMsgResultat = "A jour (" & dDateSrc & ") : " & sCheminSrc
                AfficherMsg("Pas de mise � jour � faire du fichier :" & vbCrLf &
                    sCheminSrc & vbCrLf & "(date de la derni�re �criture : " & dDateSrc & ")")
                If Not bIgnorerDateMAJ Then bSucces = True : GoTo Fin
            End If
        End If

        ' Compacter si c'est une base Access ferm�e, seulement apr�s la v�rification des dates
        '  car le compactage change logiquement la date d'�criture
        If bBaseAccessFermee And Not bBaseOuverteExclusivLectureSeule Then
            AfficherMsg("Compactage de la base en cours... :" & vbCrLf & sCheminSrc)
            If Not bCompacterMdb(sCheminSrc, sMsgErr, , , m_sMotDePasse) Then GoTo Fin
            'AfficherMsg("Compactage de la base termin� :" & vbCrLf & sCheminSrc)
            bBaseAccessDejaCompactee = True
        End If

        ' Copier le nouveau fichier � la place de l'ancien : proc�der en 2 �tapes
        '  pour �viter de supprimer la derni�re copie valable en cas de probl�me
        If Not bVerifierCreerDossier2(sCheminDossierCopie, sMsgErr) Then GoTo Fin
        AfficherMsg("Sauvegarde du fichier en cours... :" & vbCrLf & sCheminSrc)
        Dim sCheminTmp$ = sCheminDossierCopie & "\" & sNomFichier & sSuffixeCopieTmp & sExt
        Dim sSrc$ = sCheminSrc
        Dim sDest$ = sCheminTmp
        If Not bSupprimerFichier2(sDest, sMsgErr) Then GoTo Fin
        If Not bCopierFichier2(sSrc, sDest, sMsgErr) Then GoTo Fin
        ' Remplacer la copie tmp par la copie finale
        sSrc = sCheminTmp
        sDest = sCheminCopie
        If Not bRenommerFichier2(sSrc, sDest, sMsgErr) Then GoTo Fin
        'AfficherMsg("Sauvegarde du fichier termin� :" & vbCrLf & sCheminSrc)
        sMsgResultat = "Copi�  (" & dDateSrc & ") : " & sCheminSrc

        If bBaseAccess And Not bBaseAccessDejaCompactee And
           Not bBaseOuverteExclusivLectureSeule Then
            ' Compacter la copie de la base ouverte, en conservant la base d'origine 
            '  non compact�e pour pouvoir contr�ler les dates
            Dim sCheminCopieOrig$ = sCheminCopie
            sTypeCopie = m_sSuffixeBdOuverteCompactee
            'If bBaseOuverteExclusivLectureSeule Then
            'sTypeCopie = m_sSuffixeBdFermee
            ' il faudrait aussi renommer la copie fiable en copie tmp : pas grave
            'End If
            sCheminCopie = sCheminDossierCopie & "\" & sNomFichier & sTypeCopie & sExt
            AfficherMsg("Compactage de la base en cours... :" & vbCrLf & sCheminCopieOrig)
            If Not bCompacterMdb(sCheminCopieOrig, sMsgErr, sCheminCopie, ,
                m_sMotDePasse) Then GoTo Fin
            'AfficherMsg("Compactage de la base termin� :" & vbCrLf & sCheminCopieOrig)
        End If

        ' Compresser la copie (compact�e si c'est une base Access) 
        '  en utilisant un num�ro temporaire ~00.zip qui sera renomm� en ~.zip � la fin
        AfficherMsg("Compression du fichier en cours... :" & vbCrLf & sCheminCopie)
        Dim sCheminZipTmp$ = sCheminDossierCopie & "\" & sNomFichier & sTypeCopie & "00.zip"
        If Not bZipper(sCheminZipTmp, sCheminCopie, sMsgErr) Then GoTo Fin
        'AfficherMsg("Compression du fichier termin� :" & vbCrLf & sCheminCopie)

        ' Faire une copie de roulement du fichier source (n derni�res versions)
        Dim sCheminZipFin$ = ""
        If Not bRoulement(sNomFichier, sTypeCopie, sCheminZipTmp, sCheminZipFin, sMsgErr) Then GoTo Fin

        ' Faire un archivage d�finitif d'un fichier source
        Dim sCheminArch$ = ""
        If Not bArchiver(sNomFichier, sTypeCopie & m_sSuffixeArchive,
            sCheminZipFin, sCheminArch, sMsgErr) Then GoTo Fin
        bSucces = True

Fin:
        If Not bSucces Then AfficherMsg(sMsgErr)
        If Me.m_sCheminTrace <> "" Then
            Dim sMsg$ = DateTime.Now.ToShortDateString() &
                " - " & DateTime.Now.ToLongTimeString()
            sMsg &= " : " & sMsgResultat
            If bBaseAccess Then
                If bBaseAccessFermee Then
                    sMsg &= " - Base ferm�e"
                ElseIf bBaseOuverteExclusivement Then
                    sMsg &= " - Base ouverte"
                ElseIf iNbU > 0 Then ' 01/11/2009
                    sMsg &= " - Base ouverte (" & iNbU & " ut.)"
                End If
            End If
            If sMsgErr <> "" Then sMsg &= vbCrLf & sMsgErr
            TracerExecution(sMsg)
        End If

        bEnCours = False
        Sablier(bDesactiver:=True)

        ' Recommencer ind�finiment
        If bBoucleInfinie Then Exit Sub

        'm_bQuitter = True
        ' Laisser le temps de lire le statut de sauvergarde avant de quitter
        Me.timerFin.Interval = iDelaiLectureMsgMilliSec
        Me.timerFin.Start()

    End Sub

    Private Sub TracerExecution(sMsg$)

        Try
            Dim sCheminTrace$ = sDeduireChemin(m_sCheminTrace, Application.StartupPath)
            Dim sDossierTrace$ = IO.Path.GetDirectoryName(sCheminTrace)
            Dim sMsgErr$ = ""
            If Not bVerifierCreerDossier2(sDossierTrace, sMsgErr) Then
                AfficherMsg("Erreur lors de l'�criture de la trace d'ex�cution : " & vbCrLf &
                    "Chemin : " & sCheminTrace & vbCrLf & sMsgErr)
                Exit Sub
            End If
            Dim fs As IO.FileStream, sw As IO.StreamWriter
            fs = New IO.FileStream(sCheminTrace, IO.FileMode.Append, IO.FileAccess.Write)
            sw = New IO.StreamWriter(fs)
            sw.WriteLine(sMsg)
            sw.Close()
        Catch ex As Exception
            Dim sMsg0$ = "Erreur lors de l'�criture de la trace d'ex�cution : " & vbCrLf &
                "Chemin : " & m_sCheminTrace
            Dim sMsgErr$ = sMsg0 & vbCrLf & ex.Message
            AfficherMsg(sMsgErr)
            If bPromptErrGlob Then _
                AfficherMsgErreur2(ex, "bVerifierCreerDossier", sMsg0)
        End Try

    End Sub

    Private Sub AfficherMsg(sInfo$)
        Me.lblInfo.Text = sInfo
        Application.DoEvents()
    End Sub

    Public Sub Sablier(Optional bDesactiver As Boolean = False)
        If bDesactiver Then
            'Cursor.Current = Cursors.Default
            Me.Cursor = Cursors.Default
        Else
            'Cursor.Current = Cursors.WaitCursor
            Me.Cursor = Cursors.WaitCursor
        End If
    End Sub

    Private Function bRoulement(sNomFichierOrig$, sTypeCopie$,
        sCheminSrc$, ByRef sCheminDest$, ByRef sMsgErr$) As Boolean

        ' Faire une copie de roulement d'un fichier source pour conserver 
        '  les n derni�res versions d'archive temporaire
        '  et renvoyer le chemin du fichier Zip le plus r�cent : sCheminDest

        Dim i%, sSrc$, sDest$
        AfficherMsg("Copie de roulement de la base attach�e en cours... :" & vbCrLf & sCheminSrc)
        ' Chemin du dossier courant du fichier source
        Dim sCheminCourant$ = IO.Path.GetDirectoryName(sCheminSrc)
        Dim sExt$ = IO.Path.GetExtension(sCheminSrc) ' Extension du fichier � archiver

        ' Parcourir les versions � l'envers (en supprimant la derni�re)
        For i = m_iNbVersionsRoulement - 1 To -1 Step -1
            ' Remplacer l'archive n�i+1 par l'archive n�i
            sSrc = sCheminCourant & "\" & sNomFichierOrig & sTypeCopie &
                sFormater(i, m_sFormatVersionsRoulement) & sExt
            sDest = sCheminCourant & "\" & sNomFichierOrig & sTypeCopie &
                sFormater(i + 1, m_sFormatVersionsRoulement) & sExt
            If i = 0 Then
                ' Avant dernier fichier : remplacer le pr�c�dent fichier principal par l'archive n�1
                sSrc = sCheminCourant & "\" & sNomFichierOrig & sTypeCopie & sExt
            ElseIf i = -1 Then
                ' Dernier fichier : remplacer le fichier temporaire source par 
                '  le nouveau fichier principal
                sSrc = sCheminSrc
                sDest = sCheminCourant & "\" & sNomFichierOrig & sTypeCopie & sExt
                sCheminDest = sDest
            End If
            If bFichierExiste2(sSrc, sMsgErr) AndAlso Not bRenommerFichier2(sSrc, sDest, sMsgErr) Then Return False
        Next i

        'AfficherMsg("Copie de roulement de la base attach�e termin� :" & vbCrLf & sCheminDest)
        Return True

    End Function

    Private Function bArchiver(sNomFichierOrig$, sTypeCopie$, sCheminSrc$,
        ByRef sCheminDest$, ByRef sMsgErr$) As Boolean

        ' Faire un archivage d�finitif d'un fichier source si le pr�c�dent archivage est ancien :
        '  dans ce cas, le num�ro de version augmente de 1 et le fichier de destination est retourn�

        ' Chemin du dossier courant du fichier � archiver
        Dim sCheminCourant$ = IO.Path.GetDirectoryName(sCheminSrc)
        Dim sExt$ = IO.Path.GetExtension(sCheminSrc) ' Extension du fichier � archiver
        ' Filtre de recherche des archives du fichier
        Dim sFiltre$ = sNomFichierOrig & sTypeCopie & "*" & sExt
        Dim sDernierBackup$ = "" ' M�morisation du dernier fichier d'archive
        Dim di As New IO.DirectoryInfo(sCheminCourant)
        Dim fi As IO.FileInfo() = di.GetFiles(sFiltre) ' Liste des fichiers d'archives
        Dim iNbFichiers% = fi.GetLength(0)

        ' M�thode + s�re : rechercher le n� max. des fichiers, car il peut en manquer
        Dim i%
        Dim sRacine$ = sNomFichierOrig & sTypeCopie
        Dim iLenRacine% = sRacine.Length
        Dim iNumMaxFichier% = 1
        For i = 0 To iNbFichiers - 1
            Dim sFichier$ = IO.Path.GetFileNameWithoutExtension(fi(i).Name)
            Dim sNumFichier$ = sFichier.Substring(iLenRacine)
            Dim iNumFichier% = iConvertir(sNumFichier, 0)
            If iNumFichier > iNumMaxFichier Then iNumMaxFichier = iNumFichier
        Next i

        iNbFichiers = iNumMaxFichier

        Do
            ' Trouver le prochain num�ro de fichier d'archive
            sCheminDest = sCheminCourant & "\" & sNomFichierOrig & sTypeCopie &
                sFormater(iNbFichiers, m_sFormatVersionsArch) & sExt
            If Not IO.File.Exists(sCheminDest) Then Exit Do
            sDernierBackup = sCheminDest
            iNbFichiers += 1
        Loop

        If sDernierBackup <> "" Then
            ' V�rifier la date du dernier fichier d'archive
            Dim dDateSrc As Date = IO.File.GetLastWriteTime(sCheminSrc)
            Dim dDateBak As Date = IO.File.GetLastWriteTime(sDernierBackup)
            ' Arrondir au nombre de jour le plus proche en passant par les heures
            Dim iNbJours% = CInt(DateDiff(DateInterval.Hour, dDateBak, dDateSrc) / 24)
            If iNbJours < m_iPeriodeArchJours Then
                AfficherMsg("Archive assez r�cente :" & vbCrLf &
                    sCheminSrc & vbCrLf &
                    "(date actuelle : " & dDateSrc.ToShortDateString &
                    ", date archive : " & dDateBak.ToShortDateString &
                    " : nb. jours : " & iNbJours & " < " &
                    m_iPeriodeArchJours & ")")
                bArchiver = True
                Exit Function
            End If
        End If

        AfficherMsg("Archivage du fichier en cours... :" & vbCrLf & sCheminSrc)
        If Not bCopierFichier2(sCheminSrc, sCheminDest, sMsgErr) Then Return False
        AfficherMsg("Archivage du fichier termin� :" & vbCrLf & sCheminDest)
        Return True

    End Function

End Class