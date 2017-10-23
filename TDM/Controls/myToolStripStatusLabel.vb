Imports System.ComponentModel

Public Class myToolStripStatusLabel
    Inherits ToolStripStatusLabel

    Public Event InfoLevelChanged(ByVal Sender As Object)

    Private myInfoLevel As InfoLevelValues

    Sub New()
        MyBase.New()
        InfoLevel = InfoLevelValues.Information
    End Sub



    <Category("Darstellung"), _
    DefaultValue(GetType(InfoLevelValues), "0"), _
    Description("Der InfoLevel steuert die BackColor")> _
    Public Property InfoLevel() As InfoLevelValues
        Get
            Return myInfoLevel
        End Get
        Set(ByVal value As InfoLevelValues)
            myInfoLevel = value
            Select Case value
                Case InfoLevelValues.Error
                    Me.ForeColor = Color.Red
                Case InfoLevelValues.Warning
                    Me.ForeColor = Color.YellowGreen
                Case InfoLevelValues.Information
                    Me.ForeColor = Color.Green
            End Select
            RaiseEvent InfoLevelChanged(Me)
        End Set
    End Property

    Public Enum InfoLevelValues
        Information
        Warning
        [Error]
    End Enum


End Class
