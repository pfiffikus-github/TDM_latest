Imports System.Collections.ObjectModel
Imports MDM.DataLayer



Namespace LogicLayer

    <Serializable()> _
    Public Class TouchProbeCollection
        Inherits Collection(Of TouchProbe)

        Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Sub InsertItem(index As Integer, item As TouchProbe)

            For Each element In Me
                If element.ToolNumber = item.ToolNumber Then Exit Sub
            Next

            MyBase.InsertItem(index, item)
        End Sub

        Public Overrides Function ToString() As String
            Return MyClass.ToString()
        End Function


        'Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    End Class

    <Serializable()> _
    Public Class MachineCollection
        Inherits Collection(Of Machine)

        Public Event ItemsChanged(ByVal Sender As Object)

        Public Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Sub InsertItem(index As Integer, item As Machine)
            For Each element In Me
                If element.Name_ID = item.Name_ID Then Exit Sub
            Next
            MyBase.InsertItem(index, item)
            RaiseEvent ItemsChanged(Me)
        End Sub

        Protected Overrides Sub RemoveItem(index As Integer)
            MyBase.RemoveItem(index)
            RaiseEvent ItemsChanged(Me)
        End Sub

        Protected Overrides Sub ClearItems()
            MyBase.ClearItems()
            RaiseEvent ItemsChanged(Me)
        End Sub

        Public Function AllItemsAsArray() As String()
            Dim myItems(Me.Count - 1) As String
            For i As Integer = 0 To Me.Count - 1
                myItems(i) = Item(i).Name_ID
            Next
            Return myItems
        End Function


    End Class


End Namespace