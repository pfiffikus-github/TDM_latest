Imports MDM.LogicLayer
Imports MDM.DataLayer
Imports MDM.AppSharedLayer
Imports System.IO.Path


Public Class frmMDM

    Private myLocalSettings As LocalSettings
    Private myDefaultMachine As Machine



    Private Sub frmTDM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With Me

            'Init Members

            myLocalSettings = New LocalSettings()



            If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(Combine(Application.StartupPath, "XMLSave.xml")) Then
                myLocalSettings = AppSharedLayer.XMLSerialization.DeSerialization(Combine(Application.StartupPath, "XMLSave.xml"), myLocalSettings)
                PropertyGrid1.SelectedObject = myLocalSettings
            Else
                myLocalSettings = New LocalSettings()
            End If






            myDefaultMachine = myLocalSettings.Machines(0)
            .Text = Application.ProductName & " V" & Application.ProductVersion
            .dateTimeTimer_Tick(Nothing, Nothing)

            AddHandler My.Computer.Network.NetworkAvailabilityChanged, _
                Sub()
                    dateTimeTimer_Tick(Nothing, Nothing)
                End Sub

            With ToolStripStatusTempInfo
                .Text = Application.ProductName & " bereit"
                .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Information
            End With

            With ToolStripStatusTempProcess
                .Text = "Laufender Prozess=NONE"
                .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Information
            End With


        End With



    End Sub

    Private Sub dateTimeTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dateTimeTimer.Tick

        With Me
            .ToolStripStatusDateTime.Text = Date.Now.ToString("dd.MM.yy HH:mm")

            With ToolStripStatusNetworkState
                Dim subStr As String = "Netzwerk ="
                Select Case My.Computer.Network.IsAvailable
                    Case False
                        .Text = subStr & " Offline"
                        .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Warning
                    Case True
                        .Text = subStr & " Online"
                        .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Information
                End Select
            End With

            With ToolStripStatusMachineState
                Dim subStr As String = myDefaultMachine.Name_ID & " (" & myDefaultMachine.IPAddress & ") ="

                Select Case My.Computer.Network.Ping(myDefaultMachine.IPAddress)
                    Case False
                        .Text = subStr & " Offline"
                        .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Warning
                    Case True
                        .Text = subStr & " Online"
                        .InfoLevel = myToolStripStatusLabel.InfoLevelValues.Information
                End Select

            End With

        End With
        'End Using


    End Sub

    Private Sub ToolStripMenuItem_Beenden_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem_Beenden.Click
        End
    End Sub

    Private Sub ToolStripMenuItem_Konfiguration_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem_Konfiguration.Click
        Dim locDialog As New frmPropertyGrid
        Dim DialogResult As Object = locDialog.ShowDiaAndEditObject(myLocalSettings)
        If DialogResult IsNot Nothing Then
            myLocalSettings = DialogResult
            If AppSharedLayer.XMLSerialization.Serialization(Combine(Application.StartupPath, "XMLSave.xml"), myLocalSettings) = False Then
            End If
            PropertyGrid1.SelectedObject = myLocalSettings

        End If

    End Sub


End Class


