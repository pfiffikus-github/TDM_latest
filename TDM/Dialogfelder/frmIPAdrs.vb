
Public Class frmIPAdrs

    Private myIP As String
    Const mySubTitle As String = " - PING-Anfrage"

    Public Function ShowDiaAndEditObject(ByVal locString As String, Optional ByVal machineNameForTitle As String = "Maschine") As String
        Using Me
            With Me
                .txtIP.Text = locString
                .myIP = .txtIP.Text
                .Text = machineNameForTitle & " - IP-Adresse"
                .ShowDialog()
                Return myIP
            End With
        End Using
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        txtIP_TextChanged(Nothing, Nothing)
    End Sub

    Protected Overrides Sub OnClosing(e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        If Me.DialogResult = Windows.Forms.DialogResult.OK Then
            Try
                My.Computer.Network.Ping(myIP, 10)
            Catch ex As Exception
                ShowInvalidIPAddressError()
                e.Cancel = True
            End Try
            myIP = txtIP.Text
        Else
            myIP = Nothing
        End If
    End Sub

    Private Sub BtnPing_Click(sender As System.Object, e As System.EventArgs) Handles BtnPing.Click
        Dim MsgBoxTitle As String = myIP & mySubTitle
        Try
            MsgBox("Der Host """ & Me.txtIP.Text & """ ist" & IIf(My.Computer.Network.Ping(myIP, 250), " ", " NICHT ") & "erreichbar!", MsgBoxStyle.Information, myIP & mySubTitle)
        Catch ex As Exception
            ShowInvalidIPAddressError()
        End Try
    End Sub

    Private Sub txtIP_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtIP.TextChanged
        With Me
            .myIP = .txtIP.Text.Replace(" ", "")
            .BtnPing.Text = "PING " & .myIP
        End With
    End Sub

    Private Sub ShowInvalidIPAddressError()
        MsgBox("ungültige IP-Adresse!", MsgBoxStyle.Critical, myIP & mySubTitle)
    End Sub

End Class

