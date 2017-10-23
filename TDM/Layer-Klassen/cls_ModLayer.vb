Imports System.ComponentModel
Imports System.IO
Imports System.Drawing.Design
Imports System.ComponentModel.Design

Namespace ModLayer

    Friend Class UITypeEditorFrmPropertyGrid
        Inherits UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            If Not context Is Nothing AndAlso Not context.Instance Is Nothing Then
                Return UITypeEditorEditStyle.Modal
            End If
            Return UITypeEditorEditStyle.None
        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object
            Dim locDialog As New frmPropertyGrid
            Dim DialogResult As Object = locDialog.ShowDiaAndEditObject(value)
            If DialogResult IsNot Nothing Then
                Return DialogResult
            Else
                Return value
            End If
        End Function

    End Class

    Friend Class UITypeEditorEditIP
        Inherits UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            If Not context Is Nothing AndAlso Not context.Instance Is Nothing Then
                Return UITypeEditorEditStyle.Modal
            End If
            Return UITypeEditorEditStyle.None
        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object
            Dim locDialog As New frmIPAdrs
            Dim locDialogResult As String = locDialog.ShowDiaAndEditObject(value)
            If locDialogResult IsNot Nothing Then
                Return locDialogResult
            Else
                Return value
            End If
        End Function
    End Class

    Friend Class UITypeEditorSelectFile
        Inherits UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            If Not context Is Nothing AndAlso Not context.Instance Is Nothing Then
                Return UITypeEditorEditStyle.Modal
            End If
            Return UITypeEditorEditStyle.None
        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object
            Dim myDialog As New OpenFileDialog
            If myDialog.ShowDialog = DialogResult.OK Then
                Return myDialog.FileName
            Else
                Return Nothing
            End If
        End Function
    End Class

    Friend Class UITypeEditorSelectFolder
        Inherits UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            If Not context Is Nothing AndAlso Not context.Instance Is Nothing Then
                Return UITypeEditorEditStyle.Modal
            End If
            Return UITypeEditorEditStyle.None
        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object
            Dim myDialog As New FolderBrowserDialog
            myDialog.Description = "Bitte den Ordner für " & context.PropertyDescriptor.DisplayName & _
                                   " von " & context.Instance.ToString & " wählen..."
            If Directory.Exists(value) Then myDialog.SelectedPath = value.ToString
            If myDialog.ShowDialog = DialogResult.OK Then
                Return myDialog.SelectedPath
            Else
                Return Nothing
            End If
        End Function

    End Class

    Friend Class MachinesConverter
        Inherits StringConverter

        Friend Shared OptionStringArray() As String

        Public Overrides Function GetStandardValuesSupported(context As System.ComponentModel.ITypeDescriptorContext) As Boolean
            Return True
        End Function

        Public Overrides Function GetStandardValuesExclusive(context As System.ComponentModel.ITypeDescriptorContext) As Boolean
            Return True
        End Function

        Public Overrides Function GetStandardValues(context As System.ComponentModel.ITypeDescriptorContext) As System.ComponentModel.TypeConverter.StandardValuesCollection
            Return New StandardValuesCollection(MachinesConverter.OptionStringArray)
        End Function
    End Class

    Friend Class CollectionEditorWithDescription
        Inherits CollectionEditor

        Public Sub New(ByVal type As Type)
            MyBase.New(type)
        End Sub

        Protected Overrides Function CreateCollectionForm() As CollectionForm
            Dim form As CollectionForm = MyBase.CreateCollectionForm()
            With form
                .HelpButton = False
                .Size = New Size(800, 600)
                AddHandler .Shown, Sub()
                                       ShowDescription(form)
                                   End Sub
            End With
            Return form
        End Function

        Protected Overrides Function CanSelectMultipleInstances() As Boolean
            Return False
        End Function

        Private Shared Sub ShowDescription(control As Control)
            Dim grid As PropertyGrid = TryCast(control, PropertyGrid)
            If grid IsNot Nothing Then grid.HelpVisible = True
            For Each child As Control In control.Controls
                ShowDescription(child)
            Next
        End Sub

    End Class

    Friend Class IPAddressStringConverter
        Inherits StringConverter

        Public Overrides Function GetStandardValuesExclusive(context As System.ComponentModel.ITypeDescriptorContext) As Boolean
            Return True
        End Function

    End Class

End Namespace



