Imports System.ComponentModel
Imports MDM.LogicLayer
Imports System.IO
Imports System.Drawing.Design
Imports MDM.ModLayer
Imports MDM.AppSharedLayer

Namespace DataLayer

    <Serializable()> _
    Public Class Machine
        Private myName As String
        Private myInvalidCharactersInColumnName As String
        Private myIPAddress As String
        Private myMarkerPLCState As Integer : Private myMarkerPLCStateTrueValue As Boolean
        Private myMarkerTTEditState As Integer : Private myMarkerTTEditStateTrueValue As Boolean
        Private myCheckManuallyInput As Boolean
        Private myParentConnectionLogbook As String
        Private myInvalidCharactersAutoCorrection As Boolean
        Private myControl As ControlName

        Public Sub New()
        End Sub


        Public Sub InitializeMembers()
            MemberInit.initMembers(Me)
            TouchProbes.Add(New TouchProbe)
            myParentConnectionLogbook = Application.StartupPath
        End Sub


        <Category("Entwurf"), _
         DefaultValue(GetType(String), "Maschine"), _
         Description("Der Name der Maschine")> _
        Public Property Name_ID() As String
            Get
                Return myName
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing And value <> "" Then
                    myName = value.Trim.Substring(0, Math.Min(16, value.Length))
                End If
            End Set
        End Property

        <Category("Kommunikation"), _
         DefaultValue(GetType(String), "/\{}[]()+!&%$§<>;:^°|²³äüö"), _
         Description("Alle Zeichen, die in der Werkzeugtabellenspalte 'Name' ungültig sind")> _
        Public Property InvalidCharactersInColumnName() As String
            Get
                Return myInvalidCharactersInColumnName
            End Get
            Set(ByVal value As String)
                myInvalidCharactersInColumnName = value.Trim.Substring(0, Math.Min(64, value.Length))
            End Set
        End Property

        <Category("Entwurf"), _
         Editor(GetType(CollectionEditorWithDescription), GetType(UITypeEditor)), _
         Description("Alle Taster der Maschine")> _
        Public Property TouchProbes As New TouchProbeCollection

        <Category("Kommunikation"), _
         DefaultValue(GetType(String), "127.0.0.1"), _
         Editor(GetType(UITypeEditorEditIP), GetType(UITypeEditor)), _
         TypeConverter(GetType(IPAddressStringConverter)), _
         Description("Die IP-Adresse der Maschine")> _
        Public Property IPAddress As String
            Get
                Return myIPAddress
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing And value <> "" Then
                    myIPAddress = value
                End If
            End Set
        End Property

        <Category("Kommunikation"), _
         DefaultValue(GetType(Integer), "9959"), _
         Description("Der PLC-Status-Merker der Maschine")> _
        Public Property MarkerPLCState() As Integer
            Get
                Return myMarkerPLCState
            End Get
            Set(ByVal value As Integer)
                If value < 10000 And value > 0 Then
                    myMarkerPLCState = value
                End If
            End Set
        End Property

        <Category("Kommunikation"), _
         DefaultValue(GetType(Boolean), "True"), _
         Description("Der PLC-Status-Merker-TrueValue der Maschine")> _
        Public Property MarkerPLCStateTrueValue() As Boolean
            Get
                Return myMarkerPLCStateTrueValue
            End Get
            Set(ByVal value As Boolean)
                myMarkerPLCStateTrueValue = value
            End Set
        End Property

        <Category("Kommunikation"), _
         DefaultValue(GetType(Integer), "9978"), _
         Description("Der TT-Edit-Status-Merker der Maschine")> _
        Public Property MarkerTTEditState() As Integer
            Get
                Return myMarkerTTEditState
            End Get
            Set(ByVal value As Integer)
                If value < 10000 And value > 0 Then
                    myMarkerTTEditState = value
                End If
            End Set
        End Property

        <Category("Kommunikation"), _
         DefaultValue(GetType(Boolean), "True"), _
         Description("Der Werkzeugtabellen-Editierungsstatus-Merker-TrueValuee der Maschine")> _
        Public Property MarkerTTEditStateTrueValue() As Boolean
            Get
                Return myMarkerTTEditStateTrueValue
            End Get
            Set(ByVal value As Boolean)
                myMarkerTTEditStateTrueValue = value
            End Set
        End Property


        <Category("Kommunikation"), _
         DefaultValue(GetType(Boolean), "False"), _
         Description("Merker-Abgleich durch manuelle Eingabe/Bestätigung ersetzen?")> _
        Public Property CheckManuallyInput() As Boolean
            Get
                Return myCheckManuallyInput
            End Get
            Set(ByVal value As Boolean)
                myCheckManuallyInput = value
            End Set
        End Property


        <Category("Kommunikation"), _
         DefaultValue(GetType(Boolean), "True"), _
         Description("Fehlerhafte Zeichen in Spalte 'Name' automatisch korrigieren?")> _
        Public Property InvalidCharactersAutoCorrection() As Boolean
            Get
                Return myInvalidCharactersAutoCorrection
            End Get
            Set(ByVal value As Boolean)
                myInvalidCharactersAutoCorrection = value
            End Set
        End Property

        <Category("Kommunikation"), _
         Editor(GetType(UITypeEditorSelectFolder), GetType(UITypeEditor)), _
         DefaultValue(GetType(String), ""), _
         Description("Das Verzeichnis des 'großen' Verbindungslogbuches")> _
        Public Property ParentConnectionLogbook() As String
            Get
                Return myParentConnectionLogbook
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing Or value <> "" Then
                    If Directory.Exists(value) = True Then
                        myParentConnectionLogbook = value
                    End If
                End If
            End Set
        End Property


        <Category("Steuerung"), _
         DefaultValue(GetType(ControlName), "1"), _
         Description("Die Heidenhain-NC-Steuerung der Maschine")> _
        Public Property [Control]() As ControlName
            Get
                Return myControl
            End Get
            Set(ByVal value As ControlName)
                myControl = value
            End Set
        End Property


        <Category("BackUp/Kommunikation"), _
         Editor(GetType(UITypeEditorFrmPropertyGrid), GetType(UITypeEditor)), _
         Description("Die Presettabellen-BackUp-Konfiguration")> _
        Public Property PresetTableBackUp As New PresetTableBackUp



        <Category("BackUp/Kommunikation"), _
         Editor(GetType(UITypeEditorFrmPropertyGrid), GetType(UITypeEditor)), _
         Description("Die Werkzeugtabellen-BackUp-Konfiguration")> _
        Public Property ToolTableBackUp As New ToolTableBackUp


        Enum ControlName As Int16
            [Heidenhain_iTNC530] = 1
            [Heidenhain_TNC420] = 2
            [Heidenhain_TNC620] = 4
            [Heidenhain_TNC640] = 8
        End Enum


        Public Overrides Function ToString() As String
            Return Name_ID & " @ " & IPAddress
        End Function



    End Class

    <Serializable()> _
    Public Class TouchProbe

        Private myToolNumber As Integer
        Private myName As String

        Public Sub New()
        End Sub

        Public Sub InitializeMembers()
            MemberInit.initMembers(Me)
        End Sub



        <Category("Entwurf"), _
         DefaultValue(GetType(Integer), "1"), _
         Description("Die Werkzeugnummer des 3D-Tasters in der iTNC-Werkzeugtabelle")>
        Public Property ToolNumber() As Integer
            Get
                Return myToolNumber
            End Get
            Set(ByVal value As Integer)
                If value < 65000 And value > 0 Then
                    myToolNumber = value
                End If
            End Set
        End Property

        <Category("Entwurf"), _
         DefaultValue(GetType(String), "3D-Taster"), _
         Description("Der Name des 3D-Tasters")>
        Public Property Name_ID() As String
            Get
                Return myName
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing And value <> "" Then
                    myName = value.Trim.Substring(0, Math.Min(16, value.Length))
                End If
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return Name_ID & " @ " & ToolNumber
        End Function

    End Class

    <Serializable()> _
    Public Class LocalSettings
        Private myAutoConnectMachine As Boolean
        Private myDefaultMachine As String
        Private myShowMachineSelectDialog As Boolean
        Private myCheckIPs As Boolean
        Private myWriteDetailedConnectionLogbook As Boolean
        Private myGlobalSettingsFolder As String
        Private myCloseAfterSync As Boolean
        Private myAskBeforeSync As Boolean
        Private myEnableTDM As Boolean
        Private myEnableTCx000 As Boolean
        Private myEnablePDM As Boolean

        Public Sub New()
        End Sub

        Public Sub InitializeMembers()
            MemberInit.initMembers(Me)
            Machines.Add(New Machine)
            myGlobalSettingsFolder = Application.StartupPath
        End Sub

        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Mit Standardmaschine, oder ausgewählte Maschine automatisch verbinden?")> _
        Public Property AutoConnectMachine() As Boolean
            Get
                Return myAutoConnectMachine
            End Get
            Set(ByVal value As Boolean)
                myAutoConnectMachine = value
            End Set
        End Property

        <Category("Entwurf"), _
        Description("Die Standard-Maschine der lokalen Anwendungsinstanz"), _
        DefaultValue(GetType(String), "Maschine"), _
        TypeConverter(GetType(MachinesConverter)), _
        Editor(GetType(CollectionEditorWithDescription), GetType(UITypeEditor))> _
        Public Property DefaultMachine() As String
            Get
                Return myDefaultMachine
            End Get
            Set(ByVal value As String)
                myDefaultMachine = value
            End Set
        End Property

        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Maschinen-Auswahl-Dialog beim Start anzeigen?")> _
        Public Property ShowMachineSelectDialog() As Boolean
            Get
                Return myShowMachineSelectDialog
            End Get
            Set(ByVal value As Boolean)
                myShowMachineSelectDialog = value
            End Set
        End Property

        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Die Anzahl der gesendetet und empfangen IP-Pakete vergleichen/prüfen?")> _
        Public Property CheckIPs() As Boolean
            Get
                Return myCheckIPs
            End Get
            Set(ByVal value As Boolean)
                myCheckIPs = value
            End Set
        End Property

        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Ausführliches Logbuch protokollieren?")> _
        Public Property WriteDetailedConnectionLogbook() As Boolean
            Get
                Return myWriteDetailedConnectionLogbook
            End Get
            Set(ByVal value As Boolean)
                myWriteDetailedConnectionLogbook = value
            End Set
        End Property


        <Category("Entwurf"), _
        Editor(GetType(UITypeEditorSelectFolder), GetType(UITypeEditor)), _
        DefaultValue(GetType(String), ""), _
        Description("Das Verzeichnis der lokalen Einstellungsdatei")> _
        Public Property GlobalSettingsFolder() As String
            Get
                Return myGlobalSettingsFolder
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing Or value <> "" Then
                    If Directory.Exists(value) = True Then
                        myGlobalSettingsFolder = value
                    End If
                End If
            End Set
        End Property


        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "False"), _
        Description("Programm nach Datenübertragung schließen?")> _
        Public Property CloseAfterSync() As Boolean
            Get
                Return myCloseAfterSync
            End Get
            Set(ByVal value As Boolean)
                myCloseAfterSync = value
            End Set
        End Property


        <Category("Entwurf"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Abfrage vor Datenübertragung?")> _
        Public Property AskBeforeSync() As Boolean
            Get
                Return myAskBeforeSync
            End Get
            Set(ByVal value As Boolean)
                myAskBeforeSync = value
            End Set
        End Property


        <Category("Programmfunktionen aktiviren/deaktiviren"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Werkzeugdatenmanagement aktiv?")> _
        Public Property EnableTDM() As Boolean
            Get
                Return myEnableTDM
            End Get
            Set(ByVal value As Boolean)
                myEnableTDM = value
            End Set
        End Property


        <Category("Programmfunktionen aktiviren/deaktiviren"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("TCx000 aktiv?")> _
        Public Property EnableTCx000() As Boolean
            Get
                Return myEnableTCx000
            End Get
            Set(ByVal value As Boolean)
                myEnableTCx000 = value
            End Set
        End Property

        <Category("Programmfunktionen aktiviren/deaktiviren"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("Presetdatenmanagement aktiv?")> _
        Public Property EnablePDM() As Boolean
            Get
                Return myEnablePDM
            End Get
            Set(ByVal value As Boolean)
                myEnablePDM = value
            End Set
        End Property

        <Category("Entwurf"), _
         Editor(GetType(CollectionEditorWithDescription), GetType(UITypeEditor)), _
         Description("Alle Maschinen")> _
        Property Machines As New MachineCollection

        Private WithEvents MachinesCache As MachineCollection = Machines 'Event Verwaltet das richtige Anzeigen einer gültigen DefaultMachine im DropDown
        Private Sub MachinesCache_ItemsChanged(ByVal Sender As Object) Handles MachinesCache.ItemsChanged
            MachinesConverter.OptionStringArray = Machines.AllItemsAsArray
        End Sub

    End Class

    <Serializable()> _
    Public Class BackUp
        Private myBackUpFileSyntax As String
        Private myBackUpFileCheckSyntax As String
        Private myPCBackUpFolder As String
        Private myTNCBackUpFolder As String
        Private myCreatingPCBackUp As Boolean
        Private myCreatingTNCBackUp As Boolean
        Private myPCBackUpLifeTime As Integer
        Private myTNCBackUpLifeTime As Integer


        Public Sub New()
        End Sub

        Public Sub InitializeMembers()
            MemberInit.initMembers(Me)
            myPCBackUpFolder = Path.Combine(Application.StartupPath, "BackUpFiles")
        End Sub


        <Category("BackUp-Syntax"), _
        DefaultValue(GetType(String), "%data%_%time:~0,2%-%time:~3,2%-%time:~6,2%.t"), _
        Description("Der Syntax der BackUp-Datei")> _
        Public Property BackUpFileSyntax() As String
            Get
                Return myBackUpFileSyntax
            End Get
            Set(ByVal value As String)
                myBackUpFileSyntax = value
            End Set
        End Property


        <Category("BackUp-Syntax"), _
        DefaultValue(GetType(String), "dd.MM.yyyy_HH-mm-ss"), _
        Description("Der VB.NET-Date/Time Syntax der das Alter der BackUps überprüft")> _
        Public Property BackUpFileCheckSyntax() As String
            Get
                Return myBackUpFileCheckSyntax
            End Get
            Set(ByVal value As String)
                myBackUpFileCheckSyntax = value
            End Set
        End Property

        <Category("BackUp-Einstellungen-PC"), _
        Editor(GetType(UITypeEditorSelectFolder), GetType(UITypeEditor)), _
        DefaultValue(GetType(String), ""), _
        Description("Das BackUp-Verzeichnis auf dem PC/dem Netzlaufwerk")> _
        Public Property PCBackUpFolder() As String
            Get
                Return myPCBackUpFolder
            End Get
            Set(ByVal value As String)
                If value IsNot Nothing Or value <> "" Then
                    If Directory.Exists(value) = True Then
                        myPCBackUpFolder = value
                    End If
                End If
            End Set
        End Property

        <Category("BackUp-Einstellungen-TNC"), _
        DefaultValue(GetType(String), "TNC:\BackUps"), _
        Description("Das BackUp-Verzeichnis auf der Maschine")> _
        Public Property TNCBackUpFolder() As String
            Get
                Return myTNCBackUpFolder
            End Get
            Set(ByVal value As String)
                myTNCBackUpFolder = value
            End Set
        End Property


        <Category("BackUp-Einstellungen-PC"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("PC-BackUp erstellen?")> _
        Public Property CreatingPCBackUp() As Boolean
            Get
                Return myCreatingPCBackUp
            End Get
            Set(ByVal value As Boolean)
                myCreatingPCBackUp = value
            End Set
        End Property

        <Category("BackUp-Einstellungen-TNC"), _
        DefaultValue(GetType(Boolean), "True"), _
        Description("TNC-BackUp erstellen?")> _
        Public Property CreatingTNCBackUp() As Boolean
            Get
                Return myCreatingTNCBackUp
            End Get
            Set(ByVal value As Boolean)
                myCreatingTNCBackUp = value
            End Set
        End Property

        <Category("BackUp-Einstellungen-PC"), _
        DefaultValue(GetType(Integer), "3"), _
        Description("Die Lebensdauer des PC-BackUp's (in Tagen), bevor es bei der nächsten Datenübertragung wieder gelöscht wird")> _
        Public Property PCBackUpLifeTime() As Integer
            Get
                Return myPCBackUpLifeTime
            End Get
            Set(ByVal value As Integer)
                If value < 8 And value > 0 Then
                    myPCBackUpLifeTime = value
                End If
            End Set
        End Property

        <Category("BackUp-Einstellungen-TNC"), _
        DefaultValue(GetType(Integer), "3"), _
        Description("Die Lebensdauer des TNC-BackUp's (in Tagen), bevor es bei der nächsten Datenübertragung wieder gelöscht wird")> _
        Public Property TNCBackUpLifeTime() As Integer
            Get
                Return myTNCBackUpLifeTime
            End Get
            Set(ByVal value As Integer)
                If value < 8 And value > 0 Then
                    myTNCBackUpLifeTime = value
                End If
            End Set
        End Property

    End Class

    <Serializable()> _
    Public Class ToolTableBackUp
        Inherits BackUp

        Private myToolTable As String

        Public Sub New()
        End Sub

        <Category("BackUp-Datei"), _
        DefaultValue(GetType(String), "TNC:\Tool.t"), _
        Description("Die TNC-Werkzeugtabelle (einschließlich Pfad) von welcher ein TNC- oder PC-BackUp erstellt werden soll")> _
        Public Property ToolTable() As String
            Get
                Return myToolTable
            End Get
            Set(ByVal value As String)
                myToolTable = value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return ToolTable
        End Function
    End Class

    <Serializable()> _
    Public Class PresetTableBackUp
        Inherits BackUp

        Private myPresetTable As String

        Public Sub New()
        End Sub


        <Category("BackUp-Datei"), _
        DefaultValue(GetType(String), "TNC:\Preset.pr"), _
        Description("YYDie TNC-Presettabelle (einschließlich Pfad) von welcher ein TNC- oder PC-BackUp erstellt werden soll")> _
        Public Property PresetTable() As String
            Get
                Return myPresetTable
            End Get
            Set(ByVal value As String)
                myPresetTable = value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return PresetTable
        End Function

    End Class

End Namespace
