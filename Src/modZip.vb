
' Fichier modZip.vb
' -----------------

Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Module modZip

    Public Function bZipper(sCheminZip$, sCheminFichier$,
            ByRef sMsgErr$, Optional bPromptErr As Boolean = False) As Boolean

        ' Compresser le fichier sCheminFichier dans le fichier sCheminZip

        Const iTailleBuffer% = 4096
        Dim aOctets(iTailleBuffer) As Byte

        Try
            Dim zosFluxZip As New ZipOutputStream(File.Create(sCheminZip))

            zosFluxZip.SetLevel(5) ' Niveau de compression max.
            ' Possibilité de mettre un commentaire dans le fichier zip
            'zosFluxZip.SetComment("AccessBackup")

            If File.Exists(sCheminFichier) Then
                ' Ouverture en lecture du fichier à zipper 
                Dim fsFlux As FileStream
                fsFlux = File.OpenRead(sCheminFichier)

                ' Enregistrement dans le zip de la référence du fichier d'entrée 
                Dim zeFichier As New ZipEntry(Path.GetFileName(sCheminFichier))
                zosFluxZip.PutNextEntry(zeFichier)

                ' Lecture et zip du fichier par blocs de 4096 bytes 
                Dim iNbOctetsLus% = fsFlux.Read(aOctets, 0, iTailleBuffer)
                While (iNbOctetsLus > 0)
                    zosFluxZip.Write(aOctets, 0, iNbOctetsLus)
                    iNbOctetsLus = fsFlux.Read(aOctets, 0, iTailleBuffer)
                End While
                fsFlux.Flush()
                fsFlux.Close()
            End If

            zosFluxZip.Close()

            Return True

        Catch ex As Exception

            Dim sMsg$ = "Impossible de compresser le fichier :" & vbCrLf &
                sCheminFichier & vbCrLf &
                "dans le fichier Zip :" & vbCrLf & sCheminZip
            sMsgErr = sMsg & vbCrLf & ex.Message
            If bPromptErr Or bPromptErrGlob Then _
            AfficherMsgErreur2(ex, "bZipper", sMsg)
            Return False

        End Try

    End Function

End Module