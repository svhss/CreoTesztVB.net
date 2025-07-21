Imports System.Collections.Specialized.BitVector32
Imports pfcls

Module Module1

    Dim session As IpfcBaseSession = Nothing
    Dim conn As IpfcAsyncConnection = Nothing



    Sub Main()
        Try
            Connect(session, conn)
        Catch ex As Exception
            If ex.Message = "pfcExceptions::XToolkitAmbiguous" Then
                MsgBox("Egynél több Creo fut, kérlezk zárd be az egyiket.")
            End If
            Exit Sub
        End Try

        Try
            Connect(session, conn)
            ShowCreoMessage("A kapcsolat működik")
        Catch ex As Exception
            Console.WriteLine($"Error: {ex.Message}")
        Finally
            conn.Disconnect(1)
        End Try

    End Sub

    Private Sub Connect(ByRef session As IpfcBaseSession, ByRef conn As IpfcAsyncConnection)

        Dim cac As New CCpfcAsyncConnection
        conn = cac.Connect("", "", ".", 5)
        session = conn.Session

    End Sub

    Private Sub ShowCreoMessage(msg As String)

        session.UIShowMessageDialog(msg, Nothing)

    End Sub

End Module
