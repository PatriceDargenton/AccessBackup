
' Fichier modUtilDAO2.vb
' ----------------------

Option Strict Off ' Pour DAO.DBEngine.CompactDatabase

Module modUtilDAO2

    Const sClsDAOEngineCompactage$ = "DAO.DBEngine.36" ' C'est bien la derni�re version dispo.

    Public Function bCompacterMdb(sCheminBaseSrc$, ByRef sMsgErr$,
            Optional sCheminBaseDest$ = "",
            Optional bPromptErr As Boolean = False,
            Optional sMotDePasse$ = "") As Boolean

        Dim sCheminCourant$ = IO.Path.GetDirectoryName(sCheminBaseSrc)
        Dim sNomFichier$ = IO.Path.GetFileNameWithoutExtension(sCheminBaseSrc)
        ' Extension du fichier � archiver
        Dim sExt$ = IO.Path.GetExtension(sCheminBaseSrc)
        Dim bRetablirNom As Boolean = False
        If sCheminBaseDest = "" Then
            ' Si on ne pr�cise pas la base de destination, il faudra r�tablir le  
            '  nom d'origine apr�s le compactage, � partir d'un nom temporaire
            sCheminBaseDest = sCheminCourant & "\" & sNomFichier & sSuffixeCompactTmp & sExt
            bRetablirNom = True
        End If
        If Not bSupprimerFichier2(sCheminBaseDest, sMsgErr) Then Return False

        ' On a une exception ici, mais l'objet oDBE est tout de m�me cr��, le compactage fonctionne
        ' Assistant D�bogage manag� 'BindingFailure' 
        ' Message=Assistant D�bogage manag� 'BindingFailure' : 
        ' L'assembly avec le nom complet 'dao' n'a pas pu se charger dans le contexte de liaison 'LoadFrom' 
        '  de l'AppDomain ayant l'ID 1. La cause de l'erreur �tait�: System.IO.FileNotFoundException: 
        ' Impossible de charger le fichier ou l'assembly
        '  'dao, Version=10.0.4504.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35' 
        '   ou une de ses d�pendances. Le fichier sp�cifi� est introuvable.'
        ' La dll est pourtant l� : "C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll"
        Dim oDBE As Object = Nothing
        If Not bCreerObjet(oDBE, sClsDAOEngineCompactage, sMsgErr) Then Return False

        Try
            ' http://msdn.microsoft.com/en-us/library/bb220986.aspx
            If sMotDePasse.Length > 0 Then ' 01/11/2009
                ' Si la bd n'a pas de mot de passe, cela fonctionnera aussi,
                '  mais attention car la base compact�e elle sera prot�g�e !
                ' http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_20268832.html
                oDBE.CompactDatabase(sCheminBaseSrc, sCheminBaseDest,
                ";pwd=" & sMotDePasse, , ";pwd=" & sMotDePasse)
            Else
                oDBE.CompactDatabase(sCheminBaseSrc, sCheminBaseDest)
            End If

        Catch ex As Exception
            Dim sMsg$ = "Echec du compactage de la base :" & vbCrLf &
            sCheminBaseSrc & vbCrLf
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
            AfficherMsgErreur2(ex, "bCompacterMdb", sMsg)
            Return False
        Finally
            oDBE = Nothing
        End Try

        If bRetablirNom Then
            ' R�tablir le nom d'origine du fichier
            If Not bRenommerFichier2(sCheminBaseDest, sCheminBaseSrc, sMsgErr) Then
                sMsgErr = "Echec du compactage de la base :" & vbCrLf &
                sCheminBaseSrc & vbCrLf & sMsgErr
                ' En cas d'�chec, supprimer la version compact�e
                Dim sMsgErr0$ = ""
                If Not bSupprimerFichier2(sCheminBaseDest, sMsgErr0) Then Return False
                Return False
            End If
        End If
        Return True

    End Function

End Module