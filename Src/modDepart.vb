
' Fichier modDepart.vb
' --------------------

Module modDepart

    Public m_sTitreMsg$ = sNomAppli

    Public Sub DefinirTitreApplication(sTitreMsg As String)
        m_sTitreMsg = sTitreMsg
    End Sub

    Public Sub Main()

        ' modUtilFichier peut maintenant être compilé dans une dll
        DefinirTitreApplication(sTitreMsg)

        ' Laisser la possibilité de lancer plusieurs backup avec le même exe
        '  mais avec des bases distinctes
        'If bAppliDejaOuverte(bMemeExe:=True) Then Exit Sub

        ' On peut démarrer l'application sur la feuille, ou bien sur la procédure 
        '  main() si on veut pouvoir détecter l'absence de la dll sans plantage
        Dim sMsgErr$ = ""
        If Not bFichierExiste2(Application.StartupPath & "\ICSharpCode.SharpZipLib.dll", sMsgErr,
            bPrompt:=True) Then Exit Sub
        If Not bFichierExiste2(Application.StartupPath & "\adodb.dll", sMsgErr,
            bPrompt:=True) Then Exit Sub

        ' Extraire les options passées en argument de la ligne de commande
        ' Ne fonctionne pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command

        Dim sCheminSrc$ = sCheminSrcDef
        Dim sCheminTrace$ = sCheminTraceDef
        Dim sDossierSauvegardes$ = sDossierSauvegardesDef
        Dim sDossierSauvegardesIncert$ = ""
        Dim iPeriodeArchJours% = iPeriodeArchJoursDef
        Dim iNbVersionsRoulement% = iNbVersionsRoulementDef
        Dim sFormatVersionsRoulement$ = sFormatVersionsRoulementDef
        Dim sFormatVersionsArch$ = sFormatVersionsArchDef
        Dim sSuffixeArchive$ = sSuffixeArchiveDef
        Dim sSuffixeCopie$ = sSuffixeCopieDef
        Dim sSuffixeBdOuverte$ = sSuffixeBdOuverteDef
        Dim sSuffixeBdOuverteCompactee$ = sSuffixeBdOuverteCompacteeDef
        Dim sSuffixeBdFermee$ = sSuffixeBdFermeeDef
        Dim sMotDePasse$ = ""
        Dim bCompactRepair As Boolean = False

        If sArg0 <> "" Then
            Dim asArgs$() = asArgLigneCmd(sArg0)
            Dim iNbArg% = 1 + UBound(asArgs)
            If iNbArg = 1 Then
                ' 01/11/2009 Si un seul argument, on convient que c'est la bd à traiter
                sCheminSrc = asArgs(0)
                GoTo Suite
            End If
            Dim iNbPairsArg% = iNbArg \ 2
            Dim iNumArg1%
            For iNumArg1 = 0 To iNbPairsArg - 1
                Dim sCle$ = asArgs(iNumArg1 * 2)
                If iNumArg1 * 2 + 1 > iNbArg Then
                    MsgBox("Erreur : Nombre impair d'arguments !" & vbCrLf &
                        "Pensez à mettre entre guillemets les chemins contenant des espaces",
                        MsgBoxStyle.Critical, sTitreMsg)
                    Exit Sub
                End If
                Dim sVal$ = asArgs(iNumArg1 * 2 + 1)
                Select Case sCle.ToLower
                    Case "CheminSrc".ToLower
                        sCheminSrc = sVal
                    Case "DossierSauvegardes".ToLower
                        sDossierSauvegardes = sVal
                    Case "DossierSauvegardesIncert".ToLower
                        sDossierSauvegardesIncert = sVal
                    Case "CheminTrace".ToLower
                        sCheminTrace = sVal
                    Case "SuffixeArchive".ToLower
                        If sVal <> "" Then sSuffixeArchive = sVal
                    Case "SuffixeCopie".ToLower
                        If sVal <> "" Then sSuffixeCopie = sVal
                    Case "SuffixeBdOuverte".ToLower
                        If sVal <> "" Then sSuffixeBdOuverte = sVal
                    Case "SuffixeBdOuverteCompactee".ToLower
                        If sVal <> "" Then sSuffixeBdOuverteCompactee = sVal
                    Case "SuffixeBdFermee".ToLower
                        If sVal <> "" Then sSuffixeBdFermee = sVal
                    Case "FormatVersionsRoulement".ToLower
                        If sVal <> "" Then sFormatVersionsRoulement = sVal
                    Case "FormatVersionsArch".ToLower
                        If sVal <> "" Then sFormatVersionsArch = sVal
                    Case "PeriodeArchJours".ToLower
                        iPeriodeArchJours = iConvertir(sVal, iPeriodeArchJoursDef)
                    Case "NbVersionsRoulement".ToLower
                        iNbVersionsRoulement = iConvertir(sVal, iNbVersionsRoulementDef)
                    Case "MotDePasse".ToLower ' 01/11/2009
                        sMotDePasse = sVal
                    Case "CompactRepair".ToLower ' 16/03/2013
                        bCompactRepair = True
                        sCheminSrc = sVal
                End Select
            Next iNumArg1
        End If

Suite:
        If sDossierSauvegardesIncert = "" Then _
            sDossierSauvegardesIncert = sDossierSauvegardes

        Try
            Dim oFrm As frmAccessBackup
            oFrm = New frmAccessBackup
            oFrm.sCheminTrace = sCheminTrace
            oFrm.sCheminSrc = sCheminSrc
            oFrm.sDossierSauvegardes = sDossierSauvegardes
            oFrm.sDossierSauvegardesIncert = sDossierSauvegardesIncert
            oFrm.iPeriodeArchJours = iPeriodeArchJours
            oFrm.iNbVersionsRoulement = iNbVersionsRoulement
            oFrm.sFormatVersionsRoulement = sFormatVersionsRoulement
            oFrm.sFormatVersionsArch = sFormatVersionsArch
            oFrm.sSuffixeArchive = sSuffixeArchive
            oFrm.sSuffixeCopie = sSuffixeCopie
            oFrm.sSuffixeBdOuverte = sSuffixeBdOuverte
            oFrm.sSuffixeBdOuverteCompactee = sSuffixeBdOuverteCompactee
            oFrm.sSuffixeBdFermee = sSuffixeBdFermee
            oFrm.sMotDePasse = sMotDePasse
            oFrm.bCompactRepair = bCompactRepair
            ' Surtout pas ShowDialog : cela ne fonctionne pas si aucune session 
            '  n'est ouverte : le code erreur de retour est affiché dans le 
            '  planifieur de tâche : 0xe0434f4d au lieu de 0x0
            '  ou sinon une boîte de dialogue peut s'afficher même sans session ouverte, 
            '  avec le code d'erreur 0x800405a6 lié au deboguer JIT
            '  (ce bug fut difficile à trouver...)
            'oFrm.ShowDialog()
            Application.Run(oFrm)
        Catch Ex As Exception
            If bDebug Then MsgBox("Erreur : " & Ex.Message & vbCrLf & Ex.Source,
                MsgBoxStyle.Critical, sTitreMsg)
            'Catch ' (on ne peut pas mettre à la fois Catch Ex et Catch seul selon VB8)
            '    If bDebug Then MsgBox("Erreur non managée !", MsgBoxStyle.Critical, sTitreMsg)
        End Try

    End Sub

End Module