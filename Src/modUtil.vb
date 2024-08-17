
' Fichier modUtil.vb
' ------------------

Module modUtil

    Public Function iConvertir%(sVal$, iValDef%)

        ' Convertir en un entier une chaîne représentant un entier, sans provoquer 
        '  d'erreur, mais plutôt en fixant une valeur par défaut dans ce cas

        Try
            iConvertir = CInt(sVal)
        Catch 'ex As Exception
            iConvertir = iValDef
        End Try

    End Function

    Public Function bAppliDejaOuverte(bMemeExe As Boolean) As Boolean

        ' Détecter si l'application est déja lancée :
        ' - depuis n'importe quelle copie de l'exécutable, ou bien seulement
        ' - depuis le même emplacement du fichier exécutable sur le disque dur

        Dim sExeProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
        Dim sNomProcessAct$ = IO.Path.GetFileNameWithoutExtension(sExeProcessAct)

        If Not bMemeExe Then
            ' Détecter si l'application est déja lancée depuis n'importe quel exe
            If Process.GetProcessesByName(sNomProcessAct).Length > 1 Then Return True
            Return False
        End If

        ' Détecter si l'application est déja lancée depuis le même exe
        Dim sCheminProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.FileName
        Dim aProcessAct As Diagnostics.Process() = Process.GetProcessesByName(sNomProcessAct)
        Dim processAct As Diagnostics.Process
        Dim iNbApplis% = 0
        For Each processAct In aProcessAct
            Dim sCheminExe$ = processAct.MainModule.FileName
            If sCheminExe = sCheminProcessAct Then iNbApplis += 1
        Next
        If iNbApplis > 1 Then Return True
        Return False

    End Function

    Public Function sFormater$(iVal%, sFormat$)

        ' Formater un entier selon un format précisé (même syntaxe que VB6)
        sFormater = iVal.ToString(sFormat)

    End Function

    Public Function bCreerObjet(ByRef oObjetQcq As Object, sClasse$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        ' Créer une instance d'une classe OLE COM ActiveX, et renvoyer sa référence

        Try
            oObjetQcq = CreateObject(sClasse)
            Return True
        Catch ex As Exception
            oObjetQcq = Nothing
            Dim sMsg$ = "bCreerObjet : L'objet de classe [" & sClasse &
                "] ne peut pas être créé"
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
                AfficherMsgErreur2(ex, "bCreerObjet", sMsg)
            Return False
        End Try

    End Function

    Public Sub AfficherMsgErreur(ByRef Erreur As Microsoft.VisualBasic.ErrObject,
            Optional sTitreFct$ = "", Optional sInfo$ = "", Optional sDetailMsgErr$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If Erreur.Number > 0 Then
            sMsg &= vbCrLf & "Err n°" & Erreur.Number.ToString & " :"
            sMsg &= vbCrLf & Erreur.Description
        End If
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        MsgBox(sMsg, MsgBoxStyle.Critical, sTitreMsg)

    End Sub

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
            Optional sTitreFct$ = "", Optional sInfo$ = "", Optional sDetailMsgErr$ = "",
            Optional bCopierMsgPressePapier As Boolean = True, Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
                sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier", bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Function asArgLigneCmd(sLigneCmd$,
            Optional bSupprimerEspaces As Boolean = True) As String()

        ' Retourner les arguments de la ligne de commande

        ' Parser les noms longs (fonctionne quel que soit le nombre de fichiers)
        ' Chaque nom long de fichier est entre guillemets : ""
        '  une fois le nom traité, les guillemets sont enlevés
        ' S'il y a un non court parmi eux, il n'est pas entre guillemets

        ' Réutilisation de cette fonction pour parser les "" :
        ' --------------------------------------------------
        ' Cette fonction ne respecte pas le nombre de colonne, elle parse seulement les "" correctement
        '  (on pourrait cependant faire une option pour conserver les colonnes vides)
        ' Cette fonction ne sait pas non plus parser correctement une seconde ouverture de "" entre ;
        '  tel que : xxx;"x""x";xxx ou "xxx";"x""x";"xxx"
        ' En dehors des guillemets, le séparateur est l'espace et non le ;
        ' --------------------------------------------------

        Dim asArgs$() = Nothing
        If String.IsNullOrEmpty(sLigneCmd) Then
            ReDim asArgs(0)
            asArgs(0) = ""
            asArgLigneCmd = asArgs
            Exit Function
        End If

        ' Parser les noms cours : facile
        'asArgs = Split(Command, " ")

        Dim lstArgs As New List(Of String) ' 16/10/2016
        Const sGm$ = """" ' Un seul " en fait
        'sGm = Chr$(34) ' Guillemets
        Dim sFichier$, sSepar$
        Dim sCmd$, iLongCmd%, iFin%, iDeb%, iDeb2%
        Dim bFin As Boolean, bNomLong As Boolean
        Dim iCarSuiv% = 1
        sCmd = sLigneCmd
        iLongCmd = Len(sCmd)
        iDeb = 1
        Do

            bNomLong = False : sSepar = " "

            ' Chaîne vide : ""
            Dim s2Car$ = Mid(sCmd, iDeb, 2)
            If s2Car = sGm & sGm Then
                bNomLong = True : sSepar = sGm
                iFin = iDeb + 1
                GoTo Suite
            End If

            ' Si le premier caractère est un guillement, c'est un nom long
            Dim sCar$ = Mid(sCmd, iDeb, 1)
            'Dim iCar% = Asc(sCar) ' Pour debug
            If sCar = sGm Then bNomLong = True : sSepar = sGm

            iDeb2 = iDeb
            ' Supprimer les guillemets dans le tableau de fichiers
            If bNomLong AndAlso iDeb2 < iLongCmd Then iDeb2 += 1 ' Gestion chaîne vide
            iFin = InStr(iDeb2 + 1, sCmd, sSepar)

            ' 16/10/2016 On tolère que un " peut remplacer un espace
            iCarSuiv = 1
            Dim iFinGM% = InStr(iDeb2 + 1, sCmd, sGm)
            If iFinGM > 0 AndAlso iFin > 0 AndAlso iFinGM < iFin Then
                iFin = iFinGM : bNomLong = True : sSepar = sGm : iCarSuiv = 0
            End If

            ' Si le séparateur n'est pas trouvé, c'est la fin de la ligne de commande
            If iFin = 0 Then bFin = True : iFin = iLongCmd + 1

            sFichier = Mid(sCmd, iDeb2, iFin - iDeb2)
            If bSupprimerEspaces Then sFichier = Trim(sFichier)

            If sFichier.Length > 0 Then lstArgs.Add(sFichier)

            If bFin OrElse iFin = iLongCmd Then Exit Do

Suite:
            iDeb = iFin + iCarSuiv ' 1

            ' 16/10/2016 On tolère que un " peut remplacer un espace, plus besoin
            'If bNomLong Then iDeb = iFin + 2

            If iDeb > iLongCmd Then Exit Do ' 09/10/2014 Gestion chaîne vide

        Loop

        asArgs = lstArgs.ToArray()
        Const iCodeGuillemets% = 34
        For iNumArg As Integer = 0 To UBound(asArgs)
            Dim sArg$ = asArgs(iNumArg)
            ' S'il y avait 2 guillemets, il n'en reste plus qu'un
            '  on le converti en chaîne vide
            Dim iLong0% = Len(sArg)
            If iLong0 = 1 AndAlso Asc(sArg.Chars(0)) = iCodeGuillemets Then asArgs(iNumArg) = ""
        Next iNumArg

        asArgLigneCmd = asArgs

    End Function

End Module