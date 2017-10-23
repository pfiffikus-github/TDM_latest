Imports MDM.AppSharedLayer

Public Class frmPropertyGrid

    Private myLocObject As Object

    Public Function ShowDiaAndEditObject(ByVal [object] As Object) As Object
        Using Me
            ' Try

            With Me
                .myLocObject = CType(ADObjectCloner.DeepCopy([object]), Object)
                .PropertyGrid.SelectedObject = myLocObject
                .ShowDialog()
                Return myLocObject
            End With
            'Catch ex As Exception
            '    GlobalErrManagement.GlobalErrHandler(ex)
            '    Return [object]
            'End Try
        End Using


    End Function

    Private Sub OK_Button_Click(sender As System.Object, e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As System.Object, e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        PropertyGrid_PropertyValueChanged(Nothing, Nothing)
    End Sub

    Protected Overrides Sub OnClosing(e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        If Me.DialogResult = Windows.Forms.DialogResult.OK Then
            myLocObject = PropertyGrid.SelectedObject
        Else
            myLocObject = Nothing
        End If
    End Sub

    Private Sub PropertyGrid_PropertyValueChanged(s As System.Object, e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles PropertyGrid.PropertyValueChanged
        Me.Text = myLocObject.ToString & " - Konfiguration"
    End Sub

End Class


