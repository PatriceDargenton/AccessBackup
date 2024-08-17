
' Les fonctions qui existent déjà dans modUtilFichier sont numérotées avec 2
Module modUtilFichierAvecGestionErrParMsg

    Const sCauseErrPoss$ =
        "Le fichier est peut-être protégé en écriture ou bien verrouillé par une autre application"

    Public Function bFichierExiste2(sCheminFichier$,
            ByRef sMsgErr$, Optional bPrompt As Boolean = False) As Boolean

        ' Retourne True si un fichier correspondant est trouvé
        ' Ne fonctionne pas avec un filtre, par ex. du type C:\*.txt
        bFichierExiste2 = IO.File.Exists(sCheminFichier)

        If bFichierExiste2 Then Exit Function

        sMsgErr = "Impossible de trouver le fichier :" & vbCrLf & sCheminFichier
        If bPrompt Then MsgBox(sMsgErr, MsgBoxStyle.Critical, sTitreMsg & " - Fichier introuvable")

    End Function

    Public Function bSupprimerFichier2(sCheminFichier$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        ' Vérifier si le fichier existe
        If Not bFichierExiste2(sCheminFichier, sMsgErr) Then Return True

        ' Supprimer le fichier
        Try
            IO.File.Delete(sCheminFichier)
            Return True
        Catch ex As Exception
            Dim sMsg$ = "Impossible de supprimer le fichier :" & vbCrLf &
            sCheminFichier & vbCrLf & sCauseErrPoss
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
            AfficherMsgErreur2(ex, "bSupprimerFichier2", sMsg)
            Return False
        End Try

    End Function

    Public Function bCopierFichier2(sSrc$, sDest$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        Try
            IO.File.Copy(sSrc, sDest)
            Return True
        Catch ex As Exception
            Dim sMsg$ = "Impossible de copier le fichier source :" & vbCrLf &
            sSrc & vbCrLf & "vers le fichier de destination :" & vbCrLf &
            sDest & vbCrLf & sCauseErrPoss
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
            AfficherMsgErreur2(ex, "bCopierFichier2", sMsg)
            Return False
        End Try

    End Function

    Public Function bRenommerFichier2(sSrc$, sDest$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        If Not bFichierExiste2(sSrc, sMsgErr) Then Return False
        If Not bSupprimerFichier2(sDest, sMsgErr) Then Return False
        Try
            IO.File.Move(sSrc, sDest)
            Return True
        Catch ex As Exception
            Dim sMsg$ = "Impossible de renommer le fichier source :" & vbCrLf &
                sSrc & vbCrLf & "vers le fichier de destination :" & vbCrLf &
                sDest & sCauseErrPoss
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then AfficherMsgErreur2(ex, "bRenommerFichier2", sMsg)
            Return False
        End Try

    End Function

    Public Function bVerifierCreerDossier2(ByRef sCheminDossier$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        ' Vérifier et créer le dossier

        If sCheminDossier = "" Then Return True
        Dim sMsgErr0$ = "Impossible de créer le dossier :" & vbCrLf & sCheminDossier

        Dim di As IO.DirectoryInfo
        Try
            di = New IO.DirectoryInfo(sCheminDossier)
            If Not di.Exists Then
                di.Create()
                di = New IO.DirectoryInfo(sCheminDossier)
            End If
            If Not di.Exists Then
                sMsgErr = sMsgErr0
                If bPromptErr Or bPromptErrGlob Then _
                MsgBox(sMsgErr, MsgBoxStyle.Critical,
                    sTitreMsg & " - bVerifierCreerDossier")
                Return False
            End If
            Return True
        Catch ex As Exception
            sMsgErr = sMsgErr0 & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
            AfficherMsgErreur2(ex, "bVerifierCreerDossier2", sMsgErr0)
            Return False
        End Try

    End Function

    Public Function sDossierRacine$(sCheminOuDossier$)

        ' Extraire le dossier racine d'un chemin

        sDossierRacine = IO.Path.GetPathRoot(sCheminOuDossier)

        ' Ne pas considerer \ comme un dossier racine, 
        '  mais plutot le chemin comme un sous-dossier relatif
        If sDossierRacine = "\" Then sDossierRacine = ""

    End Function

    Public Function sDeduireChemin$(sDossier$, sCheminCourant$)

        ' Déuire le dossier en fonction du dossier ou chemin d'origine et
        '  du chemin courant (chemin de référence)

        Dim sLecteur$ = sDossierRacine(sDossier)
        If sLecteur = "" Then
            If sDossier.Chars(0) = "\" Then
                sDeduireChemin = sCheminCourant & sDossier
            Else
                sDeduireChemin = sCheminCourant & "\" & sDossier
            End If
        Else
            sDeduireChemin = sDossier
        End If

    End Function

End Module