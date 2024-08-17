
' Fichier modUtilDAO.vb
' ---------------------

Module modUtilDAO

    Public Function iNbUtilisateurs%(sCheminMbd$,
            ByRef sMsgErr$,
            Optional ByRef sListeUtilisateurs$ = "",
            Optional ByRef sInfoUtilisateurs$ = "",
            Optional ByRef bBaseOuverteExclusivement As Boolean = False,
            Optional ByRef bBaseOuverteExclusivLectureSeule As Boolean = False,
            Optional ByRef bBaseFiable As Boolean = True,
            Optional bPromptErr As Boolean = False,
            Optional sMotDePasse$ = "")

        ' Trouver le nombre d'utilisateur en cours d'une base de données

        ' Avantage : pas besoin de DLL (la méthode avec la dll MSLDBUSR.DLL
        '  ne marche pas en DotNet)
        ' Inconvénients : 
        ' - cette fonction ouvre une connexion ; 
        ' - en cours de développement sur une base Access, celle-ci peut être 
        '   verrouillée : impossible alors de lire la table (pas grave)
        ' - Il faut compiler en mode 32 bits, car cela ne marche pas en 64 bits :
        '   Mettre <PlatformTarget>x86</PlatformTarget> dans AccessBackup.vbproj (25/05/2013)

        bBaseFiable = True

        ' Liaison précoce ou anticipée : à la compilation
        Dim oConnADODB As ADODB.Connection
        Dim oRq As ADODB.Recordset

        If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0

        oConnADODB = New ADODB.Connection
        oRq = New ADODB.Recordset

        oConnADODB.Provider = "Microsoft.Jet.OLEDB.4.0"
        ' http://msdn.microsoft.com/en-us/library/ms676505(VS.85).aspx Open
        ' http://msdn.microsoft.com/en-us/library/ms675810(VS.85).aspx ConnectionString
        If sMotDePasse.Length > 0 Then ' 01/11/2009
            ' How to open a secured Access database in ADO through OLE DB
            ' http://support.microsoft.com/kb/191754/en-us
            oConnADODB.Open("Data Source=" & sCheminMbd &
                ";Jet OLEDB:Database Password=" & sMotDePasse)
        Else
            oConnADODB.Open("Data Source=" & sCheminMbd)
        End If

        ' Test d'ouverture d'une autre connexion
        'Dim oConn2 As New ADODB.Connection
        'oConn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
        '    & "Data Source=" & sCheminMbd

        ' The user roster is exposed as a provider-specific schema rowset
        ' in the Jet 4 OLE DB provider.  You have to use a GUID to
        ' reference the schema, as provider-specific schemas are not
        ' listed in ADO's type library for schema rowsets

        'Const adSchemaProviderSpecific& = -1 '(&HFFFFFFFF)
        Const JET_SCHEMA_USERROSTER$ =
          "{947bb102-5d43-11d1-bdbf-00c04fb92675}"
        oRq = oConnADODB.OpenSchema(
        ADODB.SchemaEnum.adSchemaProviderSpecific, , JET_SCHEMA_USERROSTER)

        ' Output the list of all users in the current database.

        'Debug.Print(oRq.Fields(0).Name, "", oRq.Fields(1).Name, _
        '    "", oRq.Fields(2).Name, oRq.Fields(3).Name)

        ' COMPUTER_NAME LOGIN_NAME CONNECTED SUSPECT_STATE
        ' X             Admin      Vrai      Null

        ' NOTES: Fields as follows
        ' 0 - COMPUTER_NAME:   Workstation
        ' 1 - LOGIN_NAME:      Name used to Login to DB
        ' 2 - CONNECTED:       True if Lock in LDB File
        ' 3 - SUSPECTED_STATE: True if user has left database in a suspect state (else Null)
        Const iColOrdi% = 0
        Const iCol_bSusp% = 3

        Dim iNumUtilisateur%, sOrdi$, sPremierUtilisateur$, sMemListeUtilisateurs$
        Dim bBaseSuspecte As Boolean
        Dim bBaseFiable0 As Boolean
        Const iNbUtilisateursAffMax% = 5
        sPremierUtilisateur = ""
        sMemListeUtilisateurs = ""

        ' Ne fonctionne pas en ADO : -1 ???
        'oRq.MoveLast()
        'iNbUtilisateurs = oRq.RecordCount
        'oRq.MoveFirst()

        While Not oRq.EOF
            bBaseSuspecte = CBool(oNz(oRq.Fields(iCol_bSusp).Value, False))
            If bBaseSuspecte Then bBaseFiable = False
            sMemListeUtilisateurs = sListeUtilisateurs
            iNumUtilisateur = iNumUtilisateur + 1
            If iNumUtilisateur > iNbUtilisateursAffMax Then
                sListeUtilisateurs = sListeUtilisateurs & "..."
            Else
                sOrdi = CStr(oNz(oRq.Fields(iColOrdi).Value, "?"))
                sOrdi = sOrdi.TrimEnd
                ' Suppression du dernier caractère
                sOrdi = sOrdi.Substring(0, sOrdi.Length - 1)
                If sPremierUtilisateur = "" Then sPremierUtilisateur = sOrdi
                If sListeUtilisateurs = "" Then
                    sListeUtilisateurs = "Utilisateur n°" &
                    iNumUtilisateur & " : " & sOrdi
                Else
                    sListeUtilisateurs &= vbCrLf & "Utilisateur n°" &
                    iNumUtilisateur & " : " & sOrdi
                End If
            End If
            oRq.MoveNext()
        End While
        If iNumUtilisateur = 0 Then
            ' Si on a réussi à ouvrir une connexion mais qu'elle n'est pas comptabilisée
            '  alors c'est que la base est ouverte en mode exclusif + lecture seule
            bBaseOuverteExclusivLectureSeule = True
            ' On considère qu'il n'y aucun utilisateur, car il ne peut y faire de modification
            '  et on peut faire une copie fiable de la base de données
            iNbUtilisateurs = 0
        Else
            ' -1 pour la connexion qui sert dans cette fonction
            iNbUtilisateurs = iNumUtilisateur - 1
            sListeUtilisateurs = sMemListeUtilisateurs
        End If

Fin:
        If Not (oRq Is Nothing) AndAlso
        oRq.State = ADODB.ObjectStateEnum.adStateOpen Then oRq.Close()
        If Not (oConnADODB Is Nothing) AndAlso
        oConnADODB.State = ADODB.ObjectStateEnum.adStateOpen Then oConnADODB.Close()

        If iNbUtilisateurs = 1 Then
            sInfoUtilisateurs = "1 seul utilisateur en cours de la base : " &
            sPremierUtilisateur
        ElseIf iNbUtilisateurs > 0 Then
            sInfoUtilisateurs = iNbUtilisateurs &
            " utilisateurs en cours de la base :" & vbCrLf &
            sListeUtilisateurs
        ElseIf iNbUtilisateurs = 0 Then
            sInfoUtilisateurs = "Base fermée."
        Else
            sInfoUtilisateurs =
            "Utilisateurs en cours de la base : ? (réessayer plus tard)"
        End If
        Exit Function

Erreur:
        iNbUtilisateurs = -1
        Dim sMsg$ =
        "Impossible d'obtenir la liste des utilisateurs connectés à la base :" &
        sCheminMbd
        ' Base ouverte en mode exclusif
        If Err.Number = -2147467259 Then bBaseOuverteExclusivement = True : Resume Fin
        sMsgErr = sMsg & vbCrLf & Err.Description
        If bPromptErr Or bPromptErrGlob Then
            AfficherMsgErreur(Err, "iNbUtilisateurs", sMsg)
            AfficherErreursADO(oConnADODB)
        End If
        Resume Fin

    End Function

    Private Sub AfficherErreursADO(ByRef oConnexion As ADODB.Connection)

        If oConnexion Is Nothing Then Exit Sub
        'If oConnexion.State <> ADODB.ObjectStateEnum.adStateOpen Then Exit Sub

        Dim sMsg$ = ""
        Dim errDB As ADODB.Error
        For Each errDB In oConnexion.Errors
            sMsg &= "Erreur ADO : " & errDB.Description & vbCrLf
            sMsg &= "Numéro : " & errDB.Number & " (" &
                Hex(errDB.Number) & "), Erreur Jet : " & errDB.SQLState & vbCrLf
            MsgBox(sMsg, MsgBoxStyle.Critical, sTitreMsg)
        Next errDB

    End Sub

    Public Function oNz(oVal As Object, Optional oDef As Object = 0) As Object

        ' Implementation de la fonction Nz d'Access en VB7 :
        ' Non Zero : renvoyer 0 (ou une autre valeur par défaut) 
        '  si la valeur du champ de bd est null
        '  ou sinon renvoyer simplement une copie de la valeur : ByVal

        ' Mieux vaut passer les objets en valeur : copie, au lieu de ref : le pointeur 
        '  sur la valeur, par ex. pour lire une valeur d'un enreg : si on garde ByRef, 
        '  on obtient une err comme quoi l'objet ne peut être mis à jour : "Informations 
        '  supplémentaires : Le jeu d'enregistrements suivant ne prend pas en charge la 
        '  mise à jour. Il s'agit peut-être d'une limitation du fournisseur ou du type 
        '  de verrou sélectionné."

        If IsDBNull(oVal) Then oNz = oDef : Exit Function
        If oVal Is System.DBNull.Value Then oNz = oDef : Exit Function
        If oVal Is Nothing Then oNz = oDef : Exit Function ' Pour les chaînes vides
        oNz = oVal

    End Function

End Module