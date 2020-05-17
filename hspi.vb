Imports System
Imports Scheduler
Imports HomeSeerAPI
Imports HSCF.Communication.Scs.Communication.EndPoints.Tcp
Imports HSCF.Communication.ScsServices.Client
Imports HSCF.Communication.ScsServices.Service
Imports System.Reflection
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Threading
Imports System.Text
Imports System.Xml
Imports System.Web

Public Class HSPI
    Inherits ScsService
    Implements IPlugInAPI     ' this API is required for ALL plugins

    Friend WithEvents MyControllerTimer As Timers.Timer
    Friend WithEvents MyAddNewDeviceTimer As Timers.Timer

    Private Const SQLFileName As String = "System.Data.SQLite.dll"
    Private Const LinuxSQLFileName As String = "LinuxSystem.Data.SQLite.dll"
    Private Const WindowsSQLFileName As String = "WindowsSystem.Data.SQLite.dll"
    Const TrialPhase As Boolean = False
    Const TrialLastDate = "May 31, 2015 12:00:00 AM"

    Public instance As String = ""
    Public isRoot As Boolean = True
    Private ConfigurationPage As PlugInConfig  ' a jquery web page
    Private UPnPViewerPage As UPnPDebugWindow

    Private InitDeviceFlag As Boolean = False
    'Private AddDeviceFlag As Boolean = False
    Private MyConfigDevice As PlayerControl
    Private MyPlayerControlWebPage As PlayerControl
    Private MyPlayerConfigWebPage As DLNADeviceConfig
    Private MyPostAnnouncementAction As PostAnnouncementAction = PostAnnouncementAction.paaForwardNoMatch
    Private MyNewDiscoveredDeviceQueue As Queue(Of String) = New Queue(Of String)()
    Private NewDeviceHandlerReEntryFlag As Boolean = False
    'Private MissedNewDeviceNotificationHandlerFlag As Boolean = False
    Private WaitTimeToAddNewDevice As Integer = 3000 'changed from 30000 in V29

    Const TriggersPageName As String = "Events"
    Const ActionsPageName As String = "Events"
    Private mvarActionAdvanced As Boolean

    Const SonosDeviceIn As Boolean = False  ' set this to true to allow control of Sonos Devices


    Const RootHSDescription As String = "Media Device Control"


#Region "Common Interface"

    Public Function Search(SearchString As String, RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements HomeSeerAPI.IPlugInAPI.Search
        ' Not yet implemented in the Sample
        '
        ' Normally we would do a search on plug-in actions, triggers, devices, etc. for the string provided, using
        '   the string as a regular expression if RegEx is True.
        '
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Search called for instance = " & instance & " and SearchString = " & SearchString & " and RegEx = " & RegEx.ToString, LogType.LOG_TYPE_INFO)
        Return Nothing
    End Function
    Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginFunction called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As MethodInfo = ty.GetMethod(proc)
            If mi Is Nothing Then
                Log("Method " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
                Return Nothing
            End If
            Return (mi.Invoke(Me, parms))
        Catch ex As Exception
            Log("Error in PluginProc: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return Nothing
    End Function
    Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginPropertyGet called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                Log("Method " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
                Return Nothing
            End If
            Return mi.GetValue(Me, Nothing)
        Catch ex As Exception
            Log("Error in PluginProc: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return Nothing
    End Function
    Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginPropertySet called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                Log("Property " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
            End If
            mi.SetValue(Me, value, Nothing)
        Catch ex As Exception
            Log("Error in PluginPropertySet: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
        Get
            'Log("Name Called for Instance = " & instance & " and returned = " & tIFACE_NAME, LogType.LOG_TYPE_INFO)
            Return tIFACE_NAME
        End Get
    End Property

    Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
        Get
            Return False
        End Get
    End Property

    Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("Capabilities called. Capabilities are IO and Music", LogType.LOG_TYPE_INFO)
        If gInterfaceStatus = ERR_NONE Then '  ' 	generate some event from all players to get ipad/iphone clients updated when they come back on-line
            CapabilitiesCalledFlag = True      ' the time procedure will pick up on this flag, send the events and reset the flag
        End If
        Return HomeSeerAPI.Enums.eCapabilities.CA_IO '+ HomeSeerAPI.Enums.eCapabilities.CA_Music 
    End Function

    Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
        ' return the access level for this plugin
        ' 1=everyone can access, no protection
        ' 2=level 2 plugin. Level 2 license required to run this plugin
        Return 1
    End Function

    Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus
        'Log("InterfaceStatus called for instance " & instance, LogType.LOG_TYPE_INFO)
        Dim es As New IPlugInAPI.strInterfaceStatus
        es.intStatus = IPlugInAPI.enumInterfaceStatus.OK
        If TrialPhase Then
            Try
                Dim d As DateTime
                d = DateTime.Parse(TrialLastDate) 'date assignment
                Dim CurrentDate As Date = Now.Date
                Dim DaysLeft As Integer = Date.Compare(d, CurrentDate)
                If DaysLeft < 0 Then
                    es.sStatus = "Time expired"
                    es.intStatus = IPlugInAPI.enumInterfaceStatus.FATAL
                Else
                    es.sStatus = "Expires: " & TrialLastDate
                    es.intStatus = IPlugInAPI.enumInterfaceStatus.WARNING
                End If
            Catch ex As Exception
                Log("Error in InterfaceStatus for Instance = " & MainInstance, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        Return es
    End Function

    Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
        Return True
    End Function

    Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstancesSingleEXE
        Return True
    End Function

    Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InstanceFriendlyName called for instance = " & instance, LogType.LOG_TYPE_INFO)
        If instance <> "" And (Not isRoot) Then
            Return GetDeviceGivenNameByUDN(instance)
        Else
            Return ""
        End If
    End Function

    Public Function InitIO(ByVal port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO

        Try
            InitIO = ""

            If TrialPhase Then
                Try
                    Dim d As DateTime
                    d = DateTime.Parse(TrialLastDate) 'date assignment
                    Dim CurrentDate As Date = Now.Date
                    If Date.Compare(CurrentDate, d) > 0 Then
                        InitIO = "Time expired, get latest version"
                        Exit Function
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO for Instance = " & instance & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            Try
                CurrentAppPath = Environment.CurrentDirectory
                ' if root directory, the currentAppPath will be ended by a | or / like in C:\, else no slash, so to make it consistent, I must remove this
                If CurrentAppPath <> "" Then
                    If CurrentAppPath(CurrentAppPath.Length - 1) = "/" Or CurrentAppPath(CurrentAppPath.Length - 1) = "\" Then
                        CurrentAppPath.Remove(CurrentAppPath.Length - 1, 1)
                    End If
                End If
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found CurrentAppPath = " & CurrentAppPath, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine the current directory path this plugin is running in.", LogType.LOG_TYPE_ERROR)
                CurrentAppPath = hs.GetAppPath
            End Try
            Try
                HSisRunningOnLinux = (hs.GetOSType() = eOSType.linux)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found HS running on Linux = " & HSisRunningOnLinux.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine what OS HS is running on.", LogType.LOG_TYPE_ERROR)
            End Try
            Try
                Log("InitIO Called for Instance = " & instance & " and running on OS = " & Environment.OSVersion.Platform.ToString, LogType.LOG_TYPE_INFO)
                ImRunningOnLinux = Type.GetType("Mono.Runtime") IsNot Nothing
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found this plugin running on Linux = " & ImRunningOnLinux.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine what OS this plugin is running on.", LogType.LOG_TYPE_ERROR)
            End Try
            If ImRunningOnLinux Then
                DBPath = "/html/" & sIFACE_NAME & "/Databases/"
                SearchResultFile = "/html/" & sIFACE_NAME & "/SearchResults/UPnPSearchResult"
                gPlaylistPath = "/html/" & sIFACE_NAME & "/Playlists/"
                'gRemoteImportPath = "/html/" & sIFACE_NAME & "/RemoteControl/Imports/"
                DebugLogFileName = "/" & tIFACE_NAME & "/Logs/MediaControllerDebug.txt"
                BinPath = "/html/" & sIFACE_NAME & "/bin/"
            End If
            If HSisRunningOnLinux Then
                FileArtWorkPath = tIFACE_NAME & "/Artwork/"
            End If
            If isRoot Then
                If ImRunningOnLinux Then
                    DebugLogFileName = CurrentAppPath & "/html" & DebugLogFileName
                Else
                    DebugLogFileName = CurrentAppPath & "\html" & DebugLogFileName
                End If
            End If
            Try
                PlugInIPAddress = hs.GetIPAddress
                PluginIPPort = hs.GetINISetting("Settings", "gWebSvrPort", "")
                ' added this code on 9/7/2019 in v3.1.0.54 to be in line with the Sonos PI functions
                Dim HSServerIPBinding = hs.GetINISetting("Settings", "gServerAddressBind", "")
                If HSServerIPBinding <> "" Then
                    If HSServerIPBinding.ToLower <> "(no binding)" Then
                        ' HS has a non default setting
                        If HSServerIPBinding = PlugInIPAddress Then
                            ' all cool here
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in InitIO for Instance = " & instance & " received (" & PlugInIPAddress & "), which is a different IP adress from it's server binding (" & HSServerIPBinding & ")", LogType.LOG_TYPE_WARNING)
                        End If
                    End If
                End If
                If ServerIPAddress <> "" Then
                    ImRunningLocal = CheckLocalIPv4Address(hs.GetIPAddress)
                    If Not ImRunningLocal Then
                        PlugInIPAddress = GetLocalIPv4Address()
                    End If
                End If
            Catch ex As Exception
                Log("Error in InitIO Called for Instance = " & instance & ". Unable to retrieve IP address info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO Called for Instance = " & instance, LogType.LOG_TYPE_INFO)

            If Not isRoot Then
                gIOEnabled = True
                gInterfaceStatus = ERR_NONE
                bShutDown = False
                InitMusicAPI()
                Exit Function
            End If

            If gIOEnabled Then Exit Function

            If ImRunningOnLinux Then
                Try
                    If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME) Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & " directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/Announcements") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/Announcements")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/Announcements directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/Playlists") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/Playlists")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/Playlists directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/SearchResults") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/SearchResults")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/SearchResults directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/Logs") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/Logs")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/Logs directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try ' this is new to support the different SQL file for Linux
                    If File.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName) Then
                        Try
                            File.Delete(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & SQLFileName)
                        Catch ex As Exception
                        End Try
                        Try
                            File.Move(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName, CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & SQLFileName)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to rename " & CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
            Else ' running under Windows
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME) Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & " directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Announcements") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Announcements")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Announcements directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Artwork") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Artwork")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Artwork directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Playlists") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Playlists")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Playlists directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\SearchResults") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\SearchResults")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\SearchResults directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try
                    If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Logs") Then
                        Try
                            Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Logs")
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Logs directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
                Try ' this is new to support the different SQL file for Linux
                    If File.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName) Then
                        Try
                            File.Delete(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName)
                        Catch ex As Exception
                            'Log("Error in InitIO. Unable to delete " & CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName & " with error = " & ex.ToString, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            File.Move(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName, CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to rename " & CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Catch ex As Exception
                End Try
            End If
            Try
                If HSisRunningOnLinux Then
                    If Not File.Exists(CurrentAppPath & "/Config/" & tIFACE_NAME & ".ini") Then
                        Try
                            WriteIntegerIniFile("Options", "piDebuglevel", DebugLevel.dlErrorsOnly)
                            WriteIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create /Config/" & tIFACE_NAME & ".ini file with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Else
                    If Not File.Exists(CurrentAppPath & "\Config\" & tIFACE_NAME & ".ini") Then
                        Try
                            WriteIntegerIniFile("Options", "piDebuglevel", DebugLevel.dlErrorsOnly)
                            WriteIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly)
                        Catch ex As Exception
                            Log("Error in InitIO. Unable to create \Config\" & tIFACE_NAME & ".ini file with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                End If
            Catch ex As Exception
            End Try

            gLogToDisk = GetBooleanIniFile("Options", "LogToDisk", False)
            If gLogToDisk Then OpenLogFile(DebugLogFileName)

            ReadIniFile()

            Try

                If MainInstance <> "" Then
                    ConfigurationPage = New PlugInConfig(ConfigPage & ":" & MainInstance)
                Else
                    ConfigurationPage = New PlugInConfig(ConfigPage)
                End If
                ConfigurationPage.RefToPlugIn = Me
                ' register the page with the HS web server, HS will post back to the WebPage class
                ' "pluginpage" is the URL to access this page
                ' comment this out if you are going to use the GenPage/PutPage API istead
                hs.RegisterPage(ConfigPage, sIFACE_NAME, MainInstance)

                ' register a configuration link that will appear on the interfaces page
                Dim wpd As New WebPageDesc
                wpd.link = ConfigPage
                If MainInstance <> "" Then
                    wpd.linktext = "Config for instance " & MainInstance
                Else
                    wpd.linktext = "Config"
                End If

                wpd.page_title = "Media Controller Plugin Config"
                wpd.plugInName = sIFACE_NAME
                wpd.plugInInstance = MainInstance
                'callback.RegisterConfigLink(wpd)

                ' register a normal page to appear in the HomeSeer menu
                wpd = New WebPageDesc
                wpd.link = ConfigPage
                If MainInstance <> "" Then
                    wpd.linktext = "Config for instance " & MainInstance
                Else
                    wpd.linktext = "Config"
                End If
                wpd.page_title = "Media Controller Config"
                wpd.plugInName = sIFACE_NAME
                wpd.plugInInstance = MainInstance
                callback.RegisterLink(wpd)

            Catch ex As Exception
                bShutDown = True
                Return "Error on InitIO: " & ex.Message
            End Try

            Try
                ' register a normal page to appear in the HomeSeer menu
                Dim Helpwpd As New WebPageDesc
                Helpwpd.link = sIFACE_NAME & "/Help/Help.htm"
                Helpwpd.linktext = "Media Controller help "
                Helpwpd.page_title = "Media Controller help"
                Helpwpd.plugInName = sIFACE_NAME
                Helpwpd.plugInInstance = MainInstance
                hs.RegisterHelpLink(Helpwpd)
            Catch ex As Exception
                Log("Error in InitIO intializing the help link with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Try
                If MainInstance <> "" Then
                    UPnPViewerPage = New UPnPDebugWindow(UPnPViewPage & ":" & MainInstance)
                Else
                    UPnPViewerPage = New UPnPDebugWindow(UPnPViewPage)
                End If
                'UPnPViewerPage.RefToPlugIn = Me
                ' register the page with the HS web server, HS will post back to the WebPage class
                ' "pluginpage" is the URL to access this page
                ' comment this out if you are going to use the GenPage/PutPage API istead
                hs.RegisterPage(UPnPViewPage, sIFACE_NAME, MainInstance)
            Catch ex As Exception
                Log("Error in InitIO intializing the UPnP Viewer link with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Dim Index As Integer
            For Index = 0 To MaxTOActionArray
                MyTimeoutActionArray(Index) = 0
            Next

            MyTimeoutActionArray(TORediscover) = TORediscoverValue
            MyTimeoutActionArray(TOCheckChange) = TOCheckChangeValue

            Try
                MyControllerTimer = New Timers.Timer
                MyControllerTimer.Interval = 1000
                MyControllerTimer.AutoReset = True
                MyControllerTimer.Enabled = True
            Catch ex As Exception
                Log("Error in InitIO. Unable to create the Timer array with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            InitDeviceFlag = True
            Log(IFACE_NAME & " Plugin Initialized", LogType.LOG_TYPE_INFO)
            gIOEnabled = True
            gInterfaceStatus = ERR_NONE
            bShutDown = False
            Return ""       ' return no error, or an error message

        Catch ex As Exception
            Log("Error in InitIO somewhere with error: " & ex.ToString, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO
        ' shutdown the I/O interface
        ' called when HS exits
        Log("ShutdownIO called for Instance = " & instance & " and isRoot = " & isRoot.ToString, LogType.LOG_TYPE_INFO)


        If isRoot Then
            hs.SetDeviceValueByRef(MasterHSDeviceRef, msDisconnected, True)
            'hs.SetDeviceString(MasterHSDeviceRef, "Disconnected", True)
            If MyControllerTimer IsNot Nothing Then MyControllerTimer.Enabled = False
            If MyAddNewDeviceTimer IsNot Nothing Then MyAddNewDeviceTimer.Enabled = False
            Try
                If ProxySpeakerActive Then
                    callback.UnRegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                End If
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " unregistering the SpeakerProxy Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Try
                StopUPnPListener()
            Catch ex As Exception

            End Try
            If MySSDPDevice IsNot Nothing Then
                Try
                    MySSDPDevice.Dispose()
                    MySSDPDevice = Nothing
                Catch ex As Exception
                    Log("Error in ShutdownIO for Instance = " & instance & " destroying the SSDP device with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            If UPnPViewerPage IsNot Nothing Then
                Try
                    UPnPViewerPage.Dispose()
                Catch ex As Exception
                End Try
                UPnPViewerPage = Nothing
            End If
            Try
                RemoveHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
            Catch ex As Exception
            End Try
            Try
                'RemoveHandler MySSDPDevice.MSearchEvent, AddressOf MSearchEvent
            Catch ex As Exception
            End Try
            Try
                DestroyUPnPControllers()
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " destroying controllers with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                If MyHSDeviceLinkedList IsNot Nothing Then
                    If MyHSDeviceLinkedList.Count > 0 Then
                        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                            If Not HSDevice Is Nothing Then
                                HSDevice.Close()
                            End If
                        Next
                        MyHSDeviceLinkedList.Clear()
                    End If
                End If
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " deleting weblinks with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            MyControllerTimer = Nothing
            MyAddNewDeviceTimer = Nothing
            CloseLogFile()
            gIOEnabled = False ' moved here 9/28/2018 v3.0.0.34
            bShutDown = True   ' moved here 9/28/2018 v3.0.0.34
        Else
            If MyEventTimer IsNot Nothing Then MyEventTimer.Enabled = False
            Try
                Disconnect(True)
            Catch ex As Exception
            End Try
            Try
                DeleteWebLink(DeviceUDN, MyUPnPDeviceName)
            Catch ex As Exception
            End Try
            Try
                DestroyPlayer(True)
            Catch ex As Exception
            End Try
            MyEventTimer = Nothing
        End If

        'gIOEnabled = False ' removed 9/17/2018 v3.0.0.32
        UPnPControllerConf = Nothing
        GC.Collect()
    End Sub

    Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
        Return True
    End Function

    Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOMulti called", LogType.LOG_TYPE_INFO)
        Dim CC As CAPIControl
        For Each CC In colSend
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOMulti set value: " & CC.ControlValue.ToString & "->ref:" & CC.Ref.ToString, LogType.LOG_TYPE_INFO)
            SetIOEx(CC)
        Next
    End Sub

    Public Sub SetIOEx(CC As CAPIControl)

        'Dim Cmd As String = ""

        If Not CC Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOEx called for Ref = " & CC.Ref.ToString & ", Index " & CC.CCIndex.ToString & ", controlFlag = " & CC.ControlFlag.ToString &
                 ", ControlString" & CC.ControlString.ToString & ", ControlType = " & CC.ControlType.ToString & ", ControlValue = " & CC.ControlValue.ToString &
                  ", Label = " & CC.Label.ToString, LogType.LOG_TYPE_INFO)
        Else
            Exit Sub    ' Not ours.
        End If

        If CC.Ref = MasterHSDeviceRef Then
            ' Check which action is required
            Select Case CC.ControlValue
                'Case msDoRediscovery 'Rediscover Players
                '   DoRediscover()
                'Case Else
            End Select
            Exit Sub
        End If
        ' [UPnP HSRef to UDN]
        Dim ccUDN As String = GetStringIniFile("UPnP HSRef to UDN", CC.Ref, "")
        If ccUDN = "" Then
            ' not found
            Log("ERROR in SetIOEx: Device not found for received event. Event = " & CC.ControlValue.ToString & " DeviceRef = " & CC.Ref.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim UPnPDevice As HSPI
        UPnPDevice = GetAPIByUDN(ccUDN)

        If UPnPDevice Is Nothing Then
            ' not found
            Log("Error in SetIOEx: Device not found for received event. Event = " & CC.ControlValue.ToString & " DeviceRef = " & CC.Ref.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        UPnPDevice.TreatSetIOEx(CC)
        UPnPDevice = Nothing

    End Sub

    Public Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent
        Log("HSEvent: " & EventType.ToString, LogType.LOG_TYPE_INFO)
        Select Case EventType
            Case Enums.HSEvent.VALUE_CHANGE
        End Select
    End Sub

    Public Function PollDevice(ByVal dvref As Integer) As IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice

    End Function

    Public Function GenPage(ByVal link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenPage called with link = " & link.ToString, LogType.LOG_TYPE_INFO)
        Return ""
    End Function

    Public Function PagePut(ByVal data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PagePut called with data = " & data.ToString, LogType.LOG_TYPE_INFO)
        Return ""
    End Function

    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin
        'If you have more than one web page, use pageName to route it to the proper GetPagePlugin
        GetPagePlugin = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("hspi.GetPagePlugin called for instance = " & instance & " and pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)
        Try
            If pageName.IndexOf(ConfigPage) = 0 Then
                Return ConfigurationPage.GetPagePlugin(pageName, user, userRights, queryString)
            ElseIf pageName.IndexOf(PlayerControlPage) = 0 Then
                If instance = "" Then
                    ' this is a problem
                    Log("Error in hspi.GetPagePlugin.PlayerControl missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString & ". Unable to retrieve UDN = " & instance, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerControlWebPage.GetPagePlugin(pageName, user, userRights, queryString, True)
            ElseIf pageName.IndexOf(PlayerConfig) = 0 Then
                If instance = "" Then
                    ' this is a problem
                    Log("Error in hspi.GetPagePlugin.PlayerConfig missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString & ". Unable to retrieve UDN = " & instance, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerConfigWebPage.GetPagePlugin(pageName, user, userRights, queryString, True)
            ElseIf pageName.IndexOf(UPnPViewPage) = 0 Then
                Return UPnPViewerPage.GetPagePlugin(pageName, user, userRights, queryString)
            Else
                Return ""
            End If
        Catch ex As Exception
            Log("Error in hspi.GetPagePlugin called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
        'If you have more than one web page, use pageName to route it to the proper postBackProc
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("hspi.PostBackProc called for instance = " & instance & " with pageName = " & pageName.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

        PostBackProc = ""
        Try
            If pageName.IndexOf(ConfigPage) = 0 Then
                Return ConfigurationPage.postBackProc(pageName, data, user, userRights)
            ElseIf pageName.IndexOf(PlayerControlPage) = 0 Then
                Dim ZoneUDN As String = ""
                Dim UDNParts As String() = pageName.Split(":")
                If UBound(UDNParts) > 0 Then
                    ZoneUDN = UDNParts(1) ' the structure is PlayerControl:RINCON_000E5859008A01400
                Else
                    ' this is a problem
                    Log("Error in hspi.PostBackProc.PlayerControl missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                If instance = "" Then
                    ' this is the root
                    Log("Error in hspi.PostBackProc.PlayerControl called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & ". Unable to retrieve UDN = " & ZoneUDN, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerControlWebPage.postBackProc(pageName, data, user, userRights)
            ElseIf pageName.IndexOf(PlayerConfig) = 0 Then
                Dim ZoneUDN As String = ""
                Dim UDNParts As String() = pageName.Split(":")
                If UBound(UDNParts) > 0 Then
                    ZoneUDN = UDNParts(1) ' the structure is PlayerControl:RINCON_000E5859008A01400
                Else
                    ' this is a problem
                    Log("Error in hspi.PostBackProc.PlayerConfig missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                If instance = "" Then
                    ' this is the root
                    Log("Error in hspi.PostBackProc.PlayerConfig called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & ". Unable to retrieve UDN = " & ZoneUDN, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerConfigWebPage.postBackProc(pageName, data, user, userRights)
            ElseIf pageName.IndexOf(UPnPViewPage) = 0 Then
                Return UPnPViewerPage.postBackProc(pageName, data, user, userRights)
            Else
                Return ""
            End If
        Catch ex As Exception
            Log("Error in hspi.PostBackProc called for instance = " & instance & " with pageName = " & pageName.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
        Return False
    End Function

    Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
        Return False
    End Function

    Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
        Return False
    End Function

#End Region

#Region "Action Properties"

    Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionCount called", LogType.LOG_TYPE_INFO)
        If Not isRoot Then Return 0
        Return 1
    End Function

    Public ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionName called with ActionNumber = " & ActionNumber.ToString, LogType.LOG_TYPE_INFO)
            Select Case ActionNumber
                Case 1
                    If MainInstance <> "" Then
                        Return "Media Controller " & MainInstance & " Actions"
                    Else
                        Return "Media Controller Actions"
                    End If
            End Select
            Return ""
        End Get
    End Property

    Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
        Set(ByVal value As Boolean)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionAdvancedMode Set called with Value = " & value.ToString, LogType.LOG_TYPE_INFO)
            mvarActionAdvanced = value
        End Set
        Get
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionAdvancedMode Get called and returned = " & mvarActionAdvanced, LogType.LOG_TYPE_INFO)
            Return mvarActionAdvanced
        End Get
    End Property

    Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
        Dim stb As New StringBuilder
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim PlayerList As New clsJQuery.jqDropList("PlayerListAction" & sUnique, ActionsPageName, True)
        Dim CommandList As New clsJQuery.jqDropList("CommandListAction" & sUnique, ActionsPageName, True)
        Dim ServerList As New clsJQuery.jqDropList("ServerListAction" & sUnique, ActionsPageName, True)
        Dim ObjectString As New clsJQuery.jqTextBox("ObjectList" & sUnique, "text", "", ActionsPageName, 1000, True)
        ObjectString.visible = False
        Dim sKey As String
        Dim action As New action

        Dim PlayerIndex As String = "" ' this is the selected UDN??
        Dim ServerIndex As String = "" ' this is the selected UDN??
        Dim CommandIndex As String = ""
        Dim InputIndex As String = ""   ' used for volume value
        Dim SelectionIndex As String = ""
        Dim NavigationSelection As String = ""
        Dim NavigationList As String = ""
        Dim AddInfoSelection As String = ""
        Dim ClearQueueSelection As String = ""
        Dim PlayNowSelection As String = ""
        Dim InputString As String = ""


        Try
            CommandList.autoPostBack = True
            CommandList.AddItem("--Please Select--", "", False)

            PlayerList.autoPostBack = True
            PlayerList.AddItem("--Please Select--", "", False)

            ServerList.autoPostBack = True
            ServerList.AddItem("--Please Select--", "", False)


            If Not (ActInfo.DataIn Is Nothing) Then
                DeSerializeObject(ActInfo.DataIn, action)
            Else 'new event, so clean out the trigger object
                action = New action
            End If

            For Each sKey In action.Keys
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found skey = " & sKey.ToString & " and PlayerUDN = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerIndex = action(sKey)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI found PlayerIndex with Actioninfo = " & PlayerIndex.ToString, LogType.LOG_TYPE_INFO)
                    Case InStr(sKey, "ServerListAction") > 0
                        ServerIndex = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        CommandIndex = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputIndex = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ObjectAction") > 0
                        NavigationSelection = action(sKey)
                    Case InStr(sKey, "ObjectList") > 0
                        NavigationList = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        ClearQueueSelection = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI found Command = " & CommandIndex & " and PlayerUDN = " & PlayerIndex & " and Text = " & InputIndex, LogType.LOG_TYPE_INFO)

            Dim InputBox As New clsJQuery.jqTextBox("InputBoxAction" & sUnique, "text", InputIndex, ActionsPageName, 40, True)
            CommandList.AddItem("Play Item", "Play Item", CommandIndex = "Play Item")
            CommandList.AddItem("Play URL", "Play URL", CommandIndex = "Play URL")
            CommandList.AddItem("Set Volume", "Set Volume", CommandIndex = "Set Volume")
            CommandList.AddItem("Set Track Position", "Set Track Position", CommandIndex = "Set Track Position")

            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.UPnPDeviceControllerRef IsNot Nothing Then
                    ' special case, just find the first player with a reference
                    Dim Playername As String = HSDevice.UPnPDeviceControllerRef.MyUPnPDeviceName
                    Dim PlayerUDN As String = HSDevice.UPnPDeviceControllerRef.DeviceUDN
                    If HSDevice.UPnPDeviceDeviceType = "DMR" Or HSDevice.UPnPDeviceDeviceType = "HST" Then
                        PlayerList.AddItem(Playername, PlayerUDN, PlayerIndex = PlayerUDN)
                    ElseIf HSDevice.UPnPDeviceDeviceType = "DMS" Then
                        ServerList.AddItem(Playername, PlayerUDN, ServerIndex = PlayerUDN)
                    End If
                End If
            Next

            stb.Append("Select Command:")
            stb.Append(CommandList.Build)

            Select Case CommandIndex
                Case "" ' start building, this is the first time called
                    Return stb.ToString
            End Select

            stb.Append("Select Player:")
            stb.Append(PlayerList.Build)

            Select Case CommandIndex
                Case "Play Item"
                    stb.Append("Select Server:")
                    stb.Append(ServerList.Build)
                    Dim ObjectListBox As New clsJQuery.jqDropList("ObjectAction" & sUnique, ActionsPageName, True)
                    ObjectListBox.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
                    stb.Append("</br>")
                    If PlayerIndex = "" Then
                        ObjectListBox.AddItem("Select Player First!", "", True)
                        stb.Append(ObjectListBox.Build)
                        Return stb.ToString
                    End If
                    If ServerIndex = "" Then
                        ObjectListBox.AddItem("Select Server First!", "", True)
                        stb.Append(ObjectListBox.Build)
                        Return stb.ToString
                    End If
                    If CommandIndex = "Play Item" Then
                        Dim ClearQueueSelectionList As New clsJQuery.jqDropList("ClearQueueAction" & sUnique, ActionsPageName, True)
                        ClearQueueSelectionList.AddItem("No", "No", "No" = ClearQueueSelection)
                        ClearQueueSelectionList.AddItem("Yes", "Yes", "Yes" = ClearQueueSelection)
                        Dim PlayNowSelectionList As New clsJQuery.jqDropList("PlayNowAction" & sUnique, ActionsPageName, True)
                        PlayNowSelectionList.AddItem("Last", "Last", "Last" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Now", "Now", "Now" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Next", "Next", "Next" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("No", "No", "No" = PlayNowSelection)
                        stb.Append("Clear Queue:" & ClearQueueSelectionList.Build)
                        stb.Append("Play Last/Now/Next/No:" & PlayNowSelectionList.Build)
                        stb.Append("</br>")
                    End If

                    Dim ServerApi As HSPI = Nothing
                    Try
                        ServerApi = MyReferenceToMyController.GetAPIByUDN(ServerIndex)
                    Catch ex As Exception
                    End Try
                    If ServerApi Is Nothing Then
                        ObjectListBox.AddItem("Server not Found!", "", True)
                        stb.Append(ObjectListBox.Build)
                        Return stb.ToString
                    End If

                    'stb.Append("</br>")
                    Try
                        Dim ObjectNavigationParts As String() = Nothing

                        Dim NewObjectString As String = ""
                        Dim NavigationObjects As Integer = 0
                        If NavigationSelection = "-- Go Up One Level --" Then ' remove the last entry
                            If NavigationList <> "" Then
                                ObjectNavigationParts = Split(NavigationList, ";--::")
                                If ObjectNavigationParts IsNot Nothing Then
                                    If ObjectNavigationParts.Count > 0 Then
                                        For Index = 1 To ObjectNavigationParts.Count - 1
                                            If NewObjectString <> "" Then NewObjectString &= ";--::"
                                            NewObjectString &= ObjectNavigationParts(Index - 1)
                                        Next
                                    End If
                                End If
                            End If
                            NavigationList = NewObjectString
                        End If
                        ' go look for the selection entries
                        Dim retNavigationPart As String = ""
                        Dim retEndOfTree As Boolean = False
                        Dim ar As DBRecord() = Nothing
                        If NavigationSelection = "-- Go Up One Level --" Or NavigationSelection = "-- End of Selection --" Then
                            ar = ServerApi.NavigateIntoContainer("", NavigationList, retNavigationPart, retEndOfTree)
                        Else
                            ar = ServerApi.NavigateIntoContainer(NavigationSelection, NavigationList, retNavigationPart, retEndOfTree)
                        End If

                        ' add the entry to the navigationlist
                        If retNavigationPart <> "" Then
                            If NavigationList <> "" Then NavigationList &= ";--::"
                            NavigationList &= retNavigationPart
                        End If
                        Dim LevelIndex As Integer = 1
                        If NavigationList <> "" Then
                            ObjectNavigationParts = Split(NavigationList, ";--::")
                            If ObjectNavigationParts IsNot Nothing Then
                                If ObjectNavigationParts.Count > 0 Then
                                    NavigationObjects = ObjectNavigationParts.Count
                                    For Each ObjectNavigationPart As String In ObjectNavigationParts
                                        If ObjectNavigationPart IsNot Nothing Then
                                            Dim ObjectParts As String() = Nothing
                                            ObjectParts = Split(ObjectNavigationPart, ";::--")
                                            If ObjectParts IsNot Nothing Then
                                                stb.Append("Selected Item Level " & LevelIndex.ToString & " : " & ObjectParts(0))
                                                stb.Append("</br>")
                                                LevelIndex += 1
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                        stb.Append("Select Item:")
                        ObjectListBox.AddItem("--Please Select--", "", True)
                        ObjectListBox.AddItem("-- End of Selection --", "-- End of Selection --", False)
                        If NavigationObjects > 1 Then
                            ObjectListBox.AddItem("-- Go Up One Level --", "-- Go Up One Level --", False)
                        End If
                        If ar IsNot Nothing And Not retEndOfTree Then
                            For Each arrecord As DBRecord In ar
                                ObjectListBox.AddItem(arrecord.Title, arrecord.Title, False)
                            Next
                        End If
                    Catch ex As Exception
                        Log("Error in ActionBuildUI1 with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try

                    ObjectString.defaultText = NavigationList
                    stb.Append(ObjectListBox.Build)
                    stb.Append(ObjectString.Build)

                    Return stb.ToString

                Case "Set Volume"
                    InputBox.size = 4
                    stb.Append("Set Volume:")
                    stb.Append(InputBox.Build)
                    Return stb.ToString
                Case "Set Track Position"
                    stb.Append("Set Track Position:")
                    stb.Append(InputBox.Build)
                    Return stb.ToString
                Case "Play URL"
                    Dim InputEditBox As New clsJQuery.jqTextBox("InputEditAction" & sUnique, "text", "", ActionsPageName, 1000, True)
                    stb.Append("Enter URL:")
                    stb.Append(InputEditBox.Build)
                    Return stb.ToString
            End Select
        Catch ex As Exception
            Log("Error in ActionBuildUI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return stb.ToString

    End Function

    Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim Configured As Boolean = False
        Dim sKey As String
        Dim itemsConfigured As Integer = 0
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return False
        End If
        Dim action As New action
        Dim PlayerUDN As String = ""
        Dim ServerUDN As String = ""
        Dim LinkList As String = ""
        Dim Command As String = ""
        Dim InputBox As String = ""
        Dim SelectionIndex As String = ""
        Dim ObjectSelection As String = ""
        Dim ObjectList As String = ""
        Dim AddInfoSelection As String = ""
        Dim PlayNowSelection As String = ""
        Dim InputString As String = ""

        Try
            DeSerializeObject(ActInfo.DataIn, action)
            For Each sKey In action.Keys
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured found sKey = " & sKey.ToString & " and Value = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        PlayerUDN = action(sKey)
                    Case InStr(sKey, "ServerListAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        ServerUDN = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0 AndAlso action(sKey) <> ""
                        Select Case Command
                            Case "Set Volume", "Set Track Position"
                                itemsConfigured += 1 ' only copy for correct command                                
                        End Select
                        InputBox = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track") Then itemsConfigured += 1
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ObjectAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Item") Then itemsConfigured += 1
                        ObjectSelection = action(sKey)
                    Case InStr(sKey, "ObjectList") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Item") Then itemsConfigured += 1
                        ObjectList = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Item") Then itemsConfigured += 1
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0 AndAlso action(sKey) <> ""
                        Select Case Command
                            Case "Play URL"
                                itemsConfigured += 1 ' only copy for correct command                                
                        End Select
                        InputString = action(sKey)
                End Select
            Next
            Select Case Command
                Case "Play Item"
                    If ObjectSelection = "-- Go Up One Level --" Then
                        Return False
                    ElseIf ObjectSelection = "-- End of Selection --" Then
                        Configured = True
                    Else
                        If itemsConfigured = 11 Then Configured = True
                    End If
                Case "Set Volume"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Track Position"
                    If itemsConfigured <> 3 Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns False", LogType.LOG_TYPE_INFO)
                        Return False
                    End If
                    If Val(InputBox) <> 0 Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns True", LogType.LOG_TYPE_INFO)
                        Return True ' valid integer
                    End If
                    Dim Index As Integer
                    Dim Counter As Integer = 0
                    For Index = 0 To InputBox.Count - 1
                        If InputBox(Index) = ":" Then Counter += 1
                    Next
                    If Counter <> 2 Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns False", LogType.LOG_TYPE_INFO)
                        Return False
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns True", LogType.LOG_TYPE_INFO)
                        Return True
                    End If
                Case "Play URL"
                    If itemsConfigured = 3 Then Configured = True
            End Select

        Catch ex As Exception
            Log("Error in ActionConfigured with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns " & Configured.ToString, LogType.LOG_TYPE_INFO)
        Return Configured

    End Function

    Public Function ActionReferencesDevice(ByVal ActInfo As IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionReferencesDevice called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString & " and dvRef = " & dvRef.ToString, LogType.LOG_TYPE_INFO)
        '
        ' Actions in the sample plug-in do not reference devices, but for demonstration purposes we will pretend they do, 
        '   and that ALL actions reference our sample devices.
        '
        If dvRef = -1 Then Return True
        Return False
    End Function

    Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI
        Dim stb As New StringBuilder
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return ""
        End If
        Dim PlayerUDN As String = ""
        Dim ServerUDN As String = ""
        Dim action As New action
        Try

            If Not (ActInfo.DataIn Is Nothing) Then
                DeSerializeObject(ActInfo.DataIn, action)
            Else 'new event, so clean out the trigger object
                action = New action
            End If

            Dim PlayerName As String = ""
            Dim ServerName As String = ""
            Dim Command As String = ""
            Dim InputBox As String = ""
            Dim SelectionIndex As String = ""
            Dim ObjectSelection As String = ""
            Dim ObjectList As String = ""
            Dim AddInfoSelection As String = ""
            Dim ClearQueueSelection As String = ""
            Dim PlayNowSelection As String = ""
            Dim InputString As String = ""


            For Each sKey In action.Keys
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found skey = " & sKey.ToString & " and PlayerUDN = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerUDN = action(sKey)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found PlayerIndex with Actioninfo = " & PlayerUDN.ToString, LogType.LOG_TYPE_INFO)
                        If PlayerUDN <> "" Then
                            PlayerName = GetDeviceGivenNameByUDN(PlayerUDN)
                        End If
                    Case InStr(sKey, "ServerListAction") > 0
                        ServerUDN = action(sKey)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found ServerIndex with Actioninfo = " & ServerUDN.ToString, LogType.LOG_TYPE_INFO)
                        If ServerUDN <> "" Then
                            ServerName = GetDeviceGivenNameByUDN(ServerUDN)
                        End If
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputBox = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ObjectAction") > 0
                        ObjectSelection = action(sKey)
                    Case InStr(sKey, "ObjectList") > 0
                        ObjectList = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        ClearQueueSelection = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next

            Dim CommandPrefix As String = ""
            If MainInstance <> "" Then
                CommandPrefix = "Media Controller Instance " & MainInstance
            Else
                CommandPrefix = "Media Controller"
            End If
            Select Case Command

                Case "Play Item"
                    stb.Append(CommandPrefix & " Action Play Item - " & " for player - " & PlayerName & " on Server - " & ServerName & ", Clear Queue: " & ClearQueueSelection & ", Play: " & PlayNowSelection)
                    stb.Append("</br>")
                    Dim LevelIndex As Integer = 1
                    If ObjectList <> "" Then
                        Dim ObjectNavigationParts As String() = Split(ObjectList, ";--::")
                        If ObjectNavigationParts IsNot Nothing Then
                            If ObjectNavigationParts.Count > 0 Then
                                For Each ObjectNavigationPart As String In ObjectNavigationParts
                                    If ObjectNavigationPart IsNot Nothing Then
                                        Dim ObjectParts As String() = Nothing
                                        ObjectParts = Split(ObjectNavigationPart, ";::--")
                                        If ObjectParts IsNot Nothing Then
                                            stb.Append("Selected Item Level " & LevelIndex.ToString & " : " & ObjectParts(0))
                                            stb.Append("</br>")
                                            LevelIndex += 1
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Case "Set Volume"
                    stb.Append(CommandPrefix & " Action Set Volume for player - " & PlayerName & " to - " & InputBox.ToString)
                Case "Set Track Position"
                    stb.Append(CommandPrefix & " Action Set Track Position for player - " & PlayerName & " - " & InputBox.ToString)
                Case "Play URL"
                    stb.Append(CommandPrefix & " Action Play URL for player - " & PlayerName & " - " & InputString.ToString)
            End Select

        Catch ex As Exception
            Log("Error in ActionFormatUI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return stb.ToString

    End Function

    Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, ByVal ActInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI
        Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI called", LogType.LOG_TYPE_INFO)

        Ret.sResult = ""
        ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        Ret.DataOut = ActInfoIN.DataIn
        Ret.TrigActInfo = ActInfoIN

        If PostData Is Nothing Then Return Ret
        If PostData.Count < 1 Then Return Ret
        Dim Action As New action
        If Not (ActInfoIN.DataIn Is Nothing) Then
            'DeSerializeObject(ActInfoIN.DataIn, Action)
        End If

        Dim parts As Collections.Specialized.NameValueCollection

        Dim sKey As String
        Dim Command As String = ""
        'Ret.DataOut = Nothing

        parts = PostData
        Try
            For Each sKey In parts.Keys
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI found sKey " & sKey.ToString, LogType.LOG_TYPE_INFO)
                If sKey Is Nothing Then Continue For
                If String.IsNullOrEmpty(sKey.Trim) Then Continue For
                Select Case True

                    Case InStr(sKey, "PlayerListAction") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "ServerListAction") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "ObjectAction") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "ObjectList") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = parts(sKey)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI found Command " & Command.ToString, LogType.LOG_TYPE_INFO)
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        Select Case Command
                            Case "Set Volume", "Set Track Position"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "SelectionAction") > 0
                        'Action.Add(CObj(parts(sKey)), sKey) ' dcor I keep this out for time being so it never gets stored!
                    Case InStr(sKey, "PlayNowAction") > 0
                        Select Case Command
                            Case "Play Item"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "InputEditAction") > 0
                        Select Case Command
                            Case "Play URL"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case Else
                        Action.Add(CObj(parts(sKey)), sKey)
                End Select
            Next
            If Not SerializeObject(Action, Ret.DataOut) Then
                Ret.sResult = sIFACE_NAME & " Error, Serialization failed. Signal Action not added."
                Return Ret
            End If
        Catch ex As Exception
            Ret.sResult = "ERROR, Exception in Action UI of " & sIFACE_NAME & ": " & ex.Message
            Return Ret
        End Try

        ' All OK
        Ret.sResult = ""
        Return Ret

    End Function

    Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
        HandleAction = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleAction called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return False
        End If
        Dim action As New action
        Try

            DeSerializeObject(ActInfo.DataIn, action)
            Dim sKey As String
            Dim PlayerUDN As String = ""
            Dim ServerIndex As String = ""
            Dim Command As String = ""
            Dim InputBox As String = ""
            Dim SelectionIndex As String = ""
            Dim AddInfoSelection As String = ""
            Dim ClearQueueSelection As Boolean = False
            Dim PlayNowSelection As String = ""
            Dim QueueAction As QueueActions = QueueActions.qaDontPlay
            Dim NavigationList As String = ""
            Dim InputString As String = ""

            For Each sKey In action.Keys
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleAction found sKey = " & sKey.ToString & " and Value = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerUDN = action(sKey)
                    Case InStr(sKey, "ServerListAction") > 0
                        ServerIndex = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputBox = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ObjectList") > 0
                        NavigationList = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        If action(sKey) = "Yes" Then
                            ClearQueueSelection = True
                        End If
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                        Select Case PlayNowSelection
                            Case "Now"
                                QueueAction = QueueActions.qaPlayNow
                            Case "Last"
                                QueueAction = QueueActions.qaPlayLast
                            Case "Next"
                                QueueAction = QueueActions.qaPlayNext
                            Case "No"
                                QueueAction = QueueActions.qaDontPlay
                        End Select
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next

            Dim MusicApi As HSPI = Nothing
            Try
                MusicApi = MyReferenceToMyController.GetAPIByUDN(PlayerUDN)
            Catch ex As Exception
            End Try
            If MusicApi Is Nothing Then Return False

            Select Case Command
                Case "Play Item"
                    MusicApi.PlayMusic(NavigationList, ServerIndex, ClearQueueSelection, QueueAction)
                    Return True
                Case "Set Volume"
                    If Len(InputBox.Trim) > 0 Then
                        Dim Vol As Integer
                        Vol = Val(InputBox.Trim)
                        Try
                            MusicApi.Volume = Vol
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                Case "Set Track Position"
                    If Len(InputBox) > 0 Then
                        Try
                            If InputBox.IndexOf(":") <> -1 Then
                                MusicApi.AVTSeek("REL_TIME", InputBox) ' this must be in the format hh:mm:ss or Sonos will error. Should already be in the right format
                            Else
                                MusicApi.AVTSeek("REL_TIME", ConvertSecondsToTimeFormat(InputBox)) ' this must be in the format hh:mm:ss or Sonos will error
                            End If
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                    Return True
                Case "Play URL"
                    MusicApi.StopPlay()
                    MusicApi.AVTSetAVTransportURI(InputString, "")
                    MusicApi.Play()
                    Return True
            End Select
        Catch ex As Exception

        End Try
    End Function

#End Region


#Region "Conditions Properties"

    Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
        ' <summary>
        ' Indicates (when True) that the Trigger is in Condition mode - it is for triggers that can also operate as a condition
        '    or for allowing Conditions to appear when a condition is being added to an event.
        ' </summary>
        ' <param name="TrigInfo">The event, group, and trigger info for this particular instance.</param>
        ' <value></value>
        ' <returns>The current state of the Condition flag.</returns>
        ' <remarks></remarks>
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Condition.get called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
            Return False
        End Get
        Set(ByVal value As Boolean)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Condition.set called for instance " & instance & " with Value = " & value.ToString & " and evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HasConditions.get called for instance " & instance & " with TriggerNumber = " & TriggerNumber.ToString, LogType.LOG_TYPE_INFO)
            Select Case TriggerNumber
                Case 1
                    Return False
                Case 2
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

#End Region

#Region "Trigger Properties"

    Public ReadOnly Property HasTriggers() As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HasTriggers called", LogType.LOG_TYPE_INFO)
            Return True
        End Get
    End Property

    Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerCount called for instance " & Instance & " ", LogType.LOG_TYPE_INFO)
            If Not isRoot Then
                Return 0
            Else
                Return 2
            End If
        End Get
    End Property

    Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerName called for instance " & Instance & " with TriggerNumber = " & TriggerNumber, LogType.LOG_TYPE_INFO)
            If MainInstance <> "" Then
                Select Case TriggerNumber
                    Case 1
                        Return "Media Controller Instance " & MainInstance & " Trigger"
                    Case 2
                        Return "Media Controller Instance " & MainInstance & " Condition"
                    Case Else
                        Return ""
                End Select
            Else
                Select Case TriggerNumber
                    Case 1
                        Return "Media Controller Trigger"
                    Case 2
                        Return "Media Controller Condition"
                    Case Else
                        Return ""
                End Select
            End If
        End Get
    End Property

    Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SubTriggerCount called for instance " & instance & " with TriggerNumber = " & TriggerNumber, LogType.LOG_TYPE_INFO)
            Return 0
        End Get
    End Property

    Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SubTriggerName called for instance " & instance & " with TriggerNumber = " & TriggerNumber, LogType.LOG_TYPE_INFO)
            Return ""
        End Get
    End Property

    Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
        Dim stb As New StringBuilder
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI called for instance " & instance & " with sUnique = " & sUnique.ToString & " and evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim PlayerList As New clsJQuery.jqDropList("PlayerListTrigger" & sUnique, TriggersPageName, True)
        Dim CommandList As New clsJQuery.jqDropList("CommandListTrigger" & sUnique, TriggersPageName, True)
        Dim trigger As New trigger
        Dim sKey As String

        CommandList.autoPostBack = True
        CommandList.AddItem("--Please Select--", "", False)

        PlayerList.autoPostBack = True
        PlayerList.AddItem("--Please Select--", "", False)

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        Else 'new event, so clean out the trigger object
            trigger = New trigger
        End If

        Dim PlayerIndex As String = "" ' this is the selected UDN??
        Dim CommandIndex As String = ""
        Dim InputIndex As String = ""
        For Each sKey In trigger.Keys
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found skey = " & sKey.ToString & " and PlayerUDN = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerIndex = trigger(sKey)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found PlayerIndex with triggerinfo = " & PlayerIndex.ToString, LogType.LOG_TYPE_INFO)
                Case InStr(sKey, "CommandListTrigger") > 0
                    CommandIndex = trigger(sKey)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found CommandIndex with triggerinfo = " & CommandIndex.ToString, LogType.LOG_TYPE_INFO)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputIndex = trigger(sKey)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found InputIndex with triggerinfo = " & CommandIndex.ToString, LogType.LOG_TYPE_INFO)
            End Select
        Next
        Dim InputBox As New clsJQuery.jqTextBox("InputBoxTrigger" & sUnique, "text", InputIndex, TriggersPageName, 40, True)
        Select Case TrigInfo.TANumber
            Case 1 '  Triggers
                CommandList.AddItem("Media Controller Track Change", "Media Controller Track Change", CommandIndex = "Media Controller Track Change")
                CommandList.AddItem("Media Controller Player Stop", "Media Controller Player Stop", CommandIndex = "Media Controller Player Stop")
                CommandList.AddItem("Media Controller Player Paused", "Media Controller Player Paused", CommandIndex = "Media Controller Player Paused")
                CommandList.AddItem("Media Controller Player Start Playing", "Media Controller Player Start Playing", CommandIndex = "Media Controller Player Start Playing")
                CommandList.AddItem("Media Controller Volume Up", "Media Controller Volume Up", CommandIndex = "Media Controller Volume Up")
                CommandList.AddItem("Media Controller Volume Down", "Media Controller Volume Down", CommandIndex = "Media Controller Volume Down")
                CommandList.AddItem("Media Controller Config Change", "Media Controller Player Config Change", CommandIndex = "Media Controller Player Config Change")
                CommandList.AddItem("Media Controller Device Online", "Media Controller Device Online", CommandIndex = "Media Controller Device Online")
                CommandList.AddItem("Media Controller Device Offline", "Media Controller Device Offline", CommandIndex = "Media Controller Device Offline")
                CommandList.AddItem("Media Controller Next Track Change", "Media Controller Next Track Change", CommandIndex = "Media Controller Next Track Change")
            Case 2 '  Conditions
                CommandList.AddItem("IsPlaying", "IsPlaying", CommandIndex = "IsPlaying")
                CommandList.AddItem("IsPaused", "IsPaused", CommandIndex = "IsPaused")
                CommandList.AddItem("IsStopped", "IsStopped", CommandIndex = "IsStopped")
                CommandList.AddItem("IsNotPlaying", "IsNotPlaying", CommandIndex = "IsNotPlaying")
                CommandList.AddItem("IsNotPaused", "IsNotPaused", CommandIndex = "IsNotPaused")
                CommandList.AddItem("IsNotStopped", "IsNotStopped", CommandIndex = "IsNotStopped")
                CommandList.AddItem("hasTrack", "hasTrack", CommandIndex = "hasTrack")
                CommandList.AddItem("hasAlbum", "hasAlbum", CommandIndex = "hasAlbum")
                CommandList.AddItem("hasArtist", "hasArtist", CommandIndex = "hasArtist")
                CommandList.AddItem("IsMutted", "IsMutted", CommandIndex = "IsMutted")
                CommandList.AddItem("IsNotMutted", "IsNotMutted", CommandIndex = "IsNotMutted")
                CommandList.AddItem("isOnline", "isOnline", CommandIndex = "isOnline")
                CommandList.AddItem("isOffline", "isOffline", CommandIndex = "isOffline")
                CommandList.AddItem("IsMusic", "IsMusic", CommandIndex = "IsMusic")
                CommandList.AddItem("IsVideo", "IsVideo", CommandIndex = "IsVideo")
                CommandList.AddItem("IsPicture", "IsPicture", CommandIndex = "IsPicture")

        End Select

        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.UPnPDeviceControllerRef IsNot Nothing Then
                ' special case, just find the first player with a reference
                Dim Playername As String = HSDevice.UPnPDeviceControllerRef.MyUPnPDeviceName
                Dim PlayerUDN As String = HSDevice.UPnPDeviceControllerRef.DeviceUDN
                PlayerList.AddItem(Playername, PlayerUDN, PlayerIndex = PlayerUDN)
            End If
        Next

        stb.Append("Select Command:")
        stb.Append(CommandList.Build)

        stb.Append("Select Player:")
        stb.Append(PlayerList.Build)

        Select Case CommandIndex
            Case "hasTrack"
                stb.Append("Specify Track:")
                stb.Append(InputBox.Build)
            Case "hasAlbum"
                stb.Append("Specify Album:")
                stb.Append(InputBox.Build)
            Case "hasArtist"
                stb.Append("Specify Artist:")
                stb.Append(InputBox.Build)
        End Select


        Return stb.ToString

    End Function

    Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
        Get
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
            Dim Configured As Boolean = False
            Dim sKey As String
            Dim itemsConfigured As Integer = 0
            Dim itemsToConfigure As Integer = 2
            Dim trigger As New trigger
            If Not (TrigInfo.DataIn Is Nothing) Then
                DeSerializeObject(TrigInfo.DataIn, trigger)
                For Each sKey In trigger.Keys
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
                    Select Case True
                        Case InStr(sKey, "PlayerListTrigger") > 0 AndAlso trigger(sKey) <> ""
                            itemsConfigured += 1
                        Case InStr(sKey, "CommandListTrigger") > 0 AndAlso trigger(sKey) <> ""
                            Select Case trigger(sKey)
                                Case "hasTrack", "hasAlbum", "hasArtist"
                                    itemsToConfigure = 3
                            End Select
                            itemsConfigured += 1
                        Case InStr(sKey, "InputBoxTrigger") > 0 AndAlso trigger(sKey) <> ""
                            itemsConfigured += 1
                    End Select
                Next
                If itemsConfigured = itemsToConfigure Then Configured = True
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured returns " & Configured.ToString, LogType.LOG_TYPE_INFO)
            Return Configured
        End Get

    End Property

    Public Function TriggerReferencesDevice(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
        '
        ' Triggers in the sample plug-in do not reference devices, but for demonstration purposes we will pretend they do, 
        '   and that ALL triggers reference our sample devices.
        '
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerReferencesDevice called for instance " & instance & " with TrigInfo = " & TrigInfo.ToString & " and dvRef = " & dvRef.ToString, LogType.LOG_TYPE_INFO)
        'If dvRef = -1 Then Return True
        Return True
    End Function

    Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerFormatUI called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim stb As New StringBuilder
        Dim sKey As String
        Dim PlayerUDN As String = ""
        Dim Command As String = ""
        Dim InputBox As String = ""
        Dim trigger As New trigger

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        Else
            Return "" ' nothing configured
        End If

        For Each sKey In trigger.Keys
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerFormatUI found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerUDN = trigger(sKey)
                Case InStr(sKey, "CommandListTrigger") > 0
                    Command = trigger(sKey)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputBox = trigger(sKey)
            End Select
        Next

        Dim PlayerName = GetDeviceGivenNameByUDN(PlayerUDN)
        Dim CommandPrefix As String
        If MainInstance <> "" Then
            CommandPrefix = "Media Controller Instance " & MainInstance
        Else
            CommandPrefix = "Media Controller"
        End If

        Select Case TrigInfo.TANumber
            Case 1 '  Trigger
                stb.Append(CommandPrefix & " trigger - " & Command & " for player - " & PlayerName)
            Case 2 '  Condition
                Select Case Command
                    Case "hasTrack"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case "hasAlbum"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case "hasArtist"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case Else
                        stb.Append(CommandPrefix & " Condition - " & Command & " for player - " & PlayerName)
                End Select

        End Select


        Return stb.ToString

    End Function

    Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection,
                                         ByVal TrigInfoIn As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerProcessPostUI called for instance " & instance & " with evRef = " & TrigInfoIn.evRef.ToString & " and SubTANumber = " & TrigInfoIn.SubTANumber.ToString & " and TANumber = " & TrigInfoIn.TANumber.ToString & " and UID = " & TrigInfoIn.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

        Ret.sResult = ""
        ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        Ret.DataOut = TrigInfoIn.DataIn
        Ret.TrigActInfo = TrigInfoIn

        If PostData Is Nothing Then Return Ret
        If PostData.Count < 1 Then Return Ret
        Dim trigger As New trigger
        If Not (TrigInfoIn.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfoIn.DataIn, trigger)
        End If

        Dim parts As Collections.Specialized.NameValueCollection

        Dim sKey As String

        parts = PostData
        Try
            For Each sKey In parts.Keys
                If sKey Is Nothing Then Continue For
                If String.IsNullOrEmpty(sKey.Trim) Then Continue For
                Select Case True
                    Case InStr(sKey, "PlayerListTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "CommandListTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "InputBoxTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                End Select
            Next
            If Not SerializeObject(trigger, Ret.DataOut) Then
                Ret.sResult = sIFACE_NAME & " Error, Serialization failed. Signal Action not added."
                Return Ret
            End If
        Catch ex As Exception
            Ret.sResult = "ERROR, Exception in Action UI of " & sIFACE_NAME & ": " & ex.Message
            Return Ret
        End Try

        ' All OK
        Ret.sResult = ""
        Return Ret
    End Function

    Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerTrue
        TriggerTrue = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerTrue called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If TrigInfo.TANumber <> 2 Then Return False ' this should not be!
        If TrigInfo.DataIn Is Nothing Then Return False ' we can't work without data
        Dim trigger As New trigger
        DeSerializeObject(TrigInfo.DataIn, trigger)

        Dim sKey As String
        Dim PlayerUDN As String = ""
        Dim Command As String = ""
        Dim InputBox As String = ""
        For Each sKey In trigger.Keys
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerTrue found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerUDN = trigger(sKey)
                Case InStr(sKey, "CommandListTrigger") > 0
                    Command = trigger(sKey)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputBox = trigger(sKey)
            End Select

        Next

        Dim MusicApi As HSPI = Nothing
        Try
            MusicApi = MyReferenceToMyController.GetAPIByUDN(PlayerUDN)
        Catch ex As Exception
        End Try
        If MusicApi Is Nothing Then Return False
        Select Case Command
            Case "IsPlaying"
                If MusicApi.PlayerState = player_state_values.Playing Then Return True Else Return False
            Case "IsPaused"
                If MusicApi.PlayerState = player_state_values.Paused Then Return True Else Return False
            Case "IsStopped"
                If MusicApi.PlayerState = player_state_values.Stopped Then Return True Else Return False
            Case "IsMutted"
                If MusicApi.RCGetMute("Master") Then Return True Else Return False
            Case "IsNotMutted"
                If MusicApi.RCGetMute("Master") Then Return False Else Return True
            Case "IsNotPlaying"
                If MusicApi.PlayerState <> player_state_values.Playing Then Return True Else Return False
            Case "IsNotPaused"
                If MusicApi.PlayerState <> player_state_values.Paused Then Return True Else Return False
            Case "IsNotStopped"
                If MusicApi.PlayerState <> player_state_values.Stopped Then Return True Else Return False
            Case "IsNotMutted"
                If MusicApi.RCGetMute("Master") Then Return False Else Return True
            Case "hasTrack"
                If Trim(MusicApi.Track.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "hasAlbum"
                If Trim(MusicApi.Album.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "hasArtist"
                If Trim(MusicApi.Artist.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "isOnline"
                If MusicApi.DeviceStatus.ToUpper = "ONLINE" Then Return True Else Return False
            Case "isOffline"
                If MusicApi.DeviceStatus.ToUpper = "OFFLINE" Then Return True Else Return False
            Case "IsMusic"
                If MusicApi.CurrentContentClass = UPnPClassType.ctMusic Then Return True Else Return False
            Case "IsVideo"
                If MusicApi.CurrentContentClass = UPnPClassType.ctVideo Then Return True Else Return False
            Case "IsPicture"
                If MusicApi.CurrentContentClass = UPnPClassType.ctPictures Then Return True Else Return False
        End Select

    End Function


#End Region

#Region "    Device Interface"

    Function ConfigDevicePost(ByVal ref As Integer, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost

        Dim ReturnValue As Integer = Enums.ConfigDevicePostReturn.DoneAndCancel

        Return ReturnValue
    End Function

    Function ConfigDevice(ByVal ref As Integer, ByVal user As String, ByVal userRights As Integer, newDevice As Boolean) As String Implements IPlugInAPI.ConfigDevice
        Return ""
    End Function

#End Region



#Region "    Plug-In Procedures    "

    Private Sub MyAddNewDeviceTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyAddNewDeviceTimer.Elapsed
        Try
            AddNewDiscoveredDevice()
            CreateUPnPControllers(True)
        Catch ex As Exception
            Log("Error in MyAddNewDeviceTimer_Elapsed. Unable to add a new device with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub


    Private Sub MyControllerTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyControllerTimer.Elapsed
        Dim Index As Integer
        For Index = 0 To MaxTOActionArray
            If MyTimeoutActionArray(Index) <> 0 Then HandleTimeout(Index)
        Next
        If CapabilitiesCalledFlag Then
            CapabilitiesCalledFlag = False
            SendEventForAllZones()
        End If
        If InitDeviceFlag Then
            InitDeviceFlag = False
            Try
                InitializeUPnPDevices()
            Catch ex As Exception
                Log("Error in HandleTimeout. Unable to initialize the UPnP Devices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            MyPIisInitialized = True
            SetSpeakerProxy()
        End If
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub HandleTimeout(ByVal TOIndex As Integer)
        Select Case TOIndex
            Case TORediscover
                ' every 5 minutes I want to see if any zones got added
                MyTimeoutActionArray(TORediscover) = MyTimeoutActionArray(TORediscover) - 1
                If MyTimeoutActionArray(TORediscover) <= 0 Then
                    If MyPIisInitialized Then DoRediscover() ' added 11/16/2019 as part of Sonos fixes
                    MyTimeoutActionArray(TORediscover) = TORediscoverValue
                End If
            Case TOCheckAnnouncement
                MyTimeoutActionArray(TOCheckAnnouncement) = MyTimeoutActionArray(TOCheckAnnouncement) - 1
                If MyTimeoutActionArray(TOCheckAnnouncement) <= 0 Then
                    DoCheckAnnouncementQueue()
                    MyTimeoutActionArray(TOCheckAnnouncement) = TOCheckAnnouncementValue
                End If
        End Select
    End Sub

    Public ReadOnly Property LocalMACAddress As String
        Get
            If MyLocalMACAddress = "" Then
                MyLocalMACAddress = GetLocalMacAddress()
            End If
            LocalMACAddress = MyLocalMACAddress
        End Get
    End Property

    Private Sub InitializeUPnPDevices()
        Dim dv As Scheduler.Classes.DeviceClass
        Dim NewStart As Boolean = GetBooleanIniFile("Options", "RefreshDevices", False)
        TCPListenerPort = GetIntegerIniFile("Options", "TCPListenerPort", 0)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeUPnPDevices called", LogType.LOG_TYPE_INFO)

        Try

            MasterHSDeviceRef = GetIntegerIniFile("Settings", "MasterHSDeviceRef", -1)

            If MasterHSDeviceRef <> -1 Then ' we already have a a masterHS device
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeUPnPDevices found MasterHSDeviceRef in the inifile. MasterHSDeviceRef = " & MasterHSDeviceRef.ToString, LogType.LOG_TYPE_INFO)
            Else
                NewStart = True
                Log("InitializeUPnPDevices is deleting all existing HS devices", LogType.LOG_TYPE_WARNING)
                hs.DeleteIODevices(sIFACE_NAME, MainInstance)
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
                MasterHSDeviceRef = hs.NewDeviceRef(HSDevices.Master.ToString)
                Log("InitializeUPnPDevices is creating a new Masterdevices with Ref = " & MasterHSDeviceRef, LogType.LOG_TYPE_INFO)
                If MasterHSDeviceRef = -1 Then   ' checks if valid ref
                    Log("Error in InitializeUPnPDevices. No More HS References available", LogType.LOG_TYPE_ERROR)
                    Exit Sub
                Else
                    Try
                        WriteIntegerIniFile("Settings", "MasterHSDeviceRef", MasterHSDeviceRef) ' saves our new base housecode
                    Catch ex As Exception
                        Log("Error in InitializeUPnPDevices while writing to the ini file : " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End Try
                End If
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
            End If

            Try
                dv = hs.GetDeviceByRef(MasterHSDeviceRef)

                If NewStart Then
                    dv.Interface(hs) = sIFACE_NAME
                    dv.Location2(hs) = "Master" ' IFACE_NAME & InstanceFriendlyName.ToString
                    dv.Location(hs) = tIFACE_NAME
                    dv.Device_Type_String(hs) = HSDevices.Master.ToString
                    dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
                    dv.Image(hs) = ImagesPath & "DLNA.png"
                    dv.ImageLarge(hs) = ImagesPath & "DLNA.png"
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeSonosDevices added image  " & ImagesPath & "DLNA.png", LogType.LOG_TYPE_INFO)
                    Dim DT As New DeviceTypeInfo
                    DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Root
                    DT.Device_SubType_Description = "Media Controller PlugIn Master Controller"
                    dv.DeviceType_Set(hs) = DT
                    dv.Status_Support(hs) = True
                    dv.Address(hs) = "Master"
                    dv.Relationship(hs) = Enums.eRelationship.Standalone
                    'dv.MISC_Set(hs, Enums.dvMISC.CONTROL_POPUP)
                    dv.InterfaceInstance(hs) = MainInstance
                    hs.SaveEventsDevices()
                    BuildButtonStringRef(MasterHSDeviceRef)
                End If

                hs.SetDeviceValueByRef(MasterHSDeviceRef, msInitializing, True)
                'hs.SetDeviceString(MasterHSDeviceRef, "Disconnected", True)

            Catch ex As Exception
                Log("Error in InitializeUPnPDevices creating the UPNP Master with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Try
                BuildUPnPHSDeviceInfoList() 'This procedure gets all the players out of the HS Database, look them up in the .ini file and puts them in the UPnPDevicesInfo array
            Catch ex As Exception
                Log("Error in InitializeUPnPDevices building UPnPDevice Info List with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If MySSDPDevice Is Nothing Then
                MySSDPDevice = New MySSDP
                If MySSDPDevice IsNot Nothing Then
                    UPnPViewerPage = New UPnPDebugWindow("UPnPViewer")
                    If UPnPViewerPage IsNot Nothing Then
                        UPnPViewerPage.RefToSSDPn = MySSDPDevice
                    End If
                End If
            End If
            Try
                DetectUPnPDevices(False) ' parameter = refresh = true for periodic check to see if new devices are discovered
            Catch ex As Exception
                Log("Error in InitializeUPnPDevices calling BuildHSUPnPDevices with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("Error in InitializeUPnPDevices with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        '
        If MyHSDeviceLinkedList.Count > 0 Then
            Try
                CreateUPnPControllers(True)
            Catch ex As Exception
                Log("Error in InitializeUPnPDevices creating the controllers with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        WriteBooleanIniFile("Options", "RefreshDevices", False) ' just in case this flag was set
        SetDeviceStringConnected()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeUPnPDevices: Done Initializing UPnPDevice Devices", LogType.LOG_TYPE_INFO)
        Try
            WriteIntegerIniFile("Options", "PreviousVersion", CurrentVersion)
        Catch ex As Exception
            Log("Error in InitializeUPnPDevices writing PreviousVersion with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub BuildButtonStringRef(Ref As Integer)
        hs.DeviceVSP_ClearAll(Ref, True)
        Try
            Dim Pair As VSPair

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = msDisconnected
            Pair.Status = "Disconnected"
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = msInitializing
            Pair.Status = "Initializing"
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = msConnected
            Pair.Status = "Connected"
            hs.DeviceVSP_AddPair(Ref, Pair)

            'Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            'Pair.PairType = VSVGPairType.SingleValue
            'Pair.Value = msDoRediscovery
            'Pair.Status = "Rediscover Devices"
            'Pair.Render = Enums.CAPIControlType.Button
            'hs.DeviceVSP_AddPair(Ref, Pair)

        Catch ex As Exception
            Log("Error in BuildButtonStringRef with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub SetDeviceStringConnected()
        hs.SetDeviceValueByRef(MasterHSDeviceRef, msConnected, True)
        'hs.SetDeviceString(MasterHSDeviceRef, "Connected", True)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetDeviceStringConnected called", LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub BuildUPnPHSDeviceInfoList()
        ' This procedure gets all the players out of the HS Database, look them up in the .ini file and puts them in the UPNPDeviceInfo array. Not all the HS devices will create record entries
        ' Some Controller instances will create multiple HS devices, only one of them as master and the other linked to the master
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildUPnPHSDeviceInfoList called", LogType.LOG_TYPE_INFO)
        Dim en As Scheduler.Classes.clsDeviceEnumeration
        Dim Device As Scheduler.Classes.DeviceClass
        Dim DeviceIndex As Integer = 0
        Dim UPnPDeviceInfo As MyUPnpDeviceInfo
        Dim DeviceUDN As String
        Dim IndexToTransport As Integer = 0
        Try
            en = hs.GetDeviceEnumerator
            While Not en.Finished
                Device = en.GetNext
                If Device.Interface(Nothing) = sIFACE_NAME And Device.InterfaceInstance(Nothing) = MainInstance Then
                    Dim DT As DeviceTypeInfo = Device.DeviceType_Get(Nothing)
                    If DT.Device_SubType_Description = RootHSDescription Then
                        DeviceUDN = GetStringIniFile("UPnP HSRef to UDN", Device.Ref(Nothing), "")
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "BuildUPnPHSDeviceInfoList found " & Device.dc & " with UDN = " & DeviceUDN.ToString)
                        If DeviceUDN <> "" Then
                            ' the info was stored
                            UPnPDeviceInfo = DMAdd()
                            If UPnPDeviceInfo Is Nothing Then
                                ' this should realy not happen
                                Log("Error in BuildUPnPHSDeviceInfoList. Unable to add an entry to the HSDeviceInfo array for  DeviceRef = " & Device.Ref(Nothing), LogType.LOG_TYPE_ERROR)
                                Exit Sub
                            End If
                            UPnPDeviceInfo.UPnPDeviceGivenName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diGivenName.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceUDN = DeviceUDN
                            UPnPDeviceInfo.UPnPDeviceModelName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceModelName.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceDeviceType = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceServiceTypes = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceServiceTypes.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceFriendlyName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diFriendlyName.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceIPAddress = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diIPAddress.ToString, "")  ' this is the last known IPAddress
                            UPnPDeviceInfo.UPnPDeviceIPPort = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diIPPort.ToString, "")   ' this is the last known IPPort
                            UPnPDeviceInfo.UPnPDeviceMusicAPIIndex = GetIntegerIniFile(DeviceUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, 0)
                            UPnPDeviceInfo.UPnPDeviceAdminStateActive = GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diAdminState.ToString, False)
                            UPnPDeviceInfo.UPnPDeviceIsAddedToHS = GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
                            UPnPDeviceInfo.UPnPDeviceIconURL = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIConURL.ToString, "")
                            UPnPDeviceInfo.UPnPDeviceHSRef = Device.Ref(Nothing)
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildUPnPHSDeviceInfoList found " & UPnPDeviceInfo.UPnPDeviceGivenName & " at Index " & DeviceIndex.ToString, LogType.LOG_TYPE_INFO)
                            DeviceIndex = DeviceIndex + 1
                        Else
                            ' this could be the master or one of the many additional controls
                            Dim DeviceInfo As String = GetStringIniFile("Settings", "MasterHSDeviceRef", "")
                            If DeviceInfo = "" Then
                                ' OK this can be additional HS devices created by the Controller Instances themselves
                                DeviceInfo = GetStringIniFile("UPnP HSRef to UDN", Device.Ref(Nothing), "")
                                If DeviceInfo <> "" Then
                                    ' this is good
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("BuildUPnPHSDeviceInfoList found linked devices in .ini file with deviceRef = " & Device.Ref(Nothing), LogType.LOG_TYPE_INFO)
                                Else
                                    ' this is not good
                                    Log("Error in BuildUPnPHSDeviceInfoList. Info not found in .ini file Devtype = UPnP Devices and DeviceRef = " & Device.Ref(Nothing), LogType.LOG_TYPE_ERROR)
                                End If
                            End If
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            Log("Error in BuildUPnPHSDeviceInfoList with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DetectUPnPDevices(ByVal Refresh As Boolean)
        Dim UPnPDeviceInfoArray As MyUPnpDeviceInfo() = Nothing
        ReDim UPnPDeviceInfoArray(0)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DetectUPnPDevices called with Refresh = " & Refresh.ToString, LogType.LOG_TYPE_INFO)

        Dim UPnPDevicesToGoDiscover As New System.Collections.Generic.Dictionary(Of String, String)()
        UPnPDevicesToGoDiscover = GetIniSection("UPnP Devices to discover") '  As Dictionary(Of String, String)

        If UPnPDevicesToGoDiscover Is Nothing Then
            Log("Error in DetectUPnPDevices. No Devices specified in the .ini file under ""UPnP Devices to discover""", LogType.LOG_TYPE_ERROR)
            UPnPDeviceInfoArray = Nothing
            Exit Sub
        End If
        Try
            FindUPnPDevice(UPnPDeviceInfoArray) ' go find the devices using UPNP discovery
        Catch ex As Exception
            Log("Error in DetectUPnPDevices while finding the UPnPDevices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DetectUPnPDevices added a total of " & (UBound(UPnPDeviceInfoArray)).ToString & " devices", LogType.LOG_TYPE_INFO)

        Dim UPnPDeviceUDN As String
        Dim UPnPDeviceInfo As MyUPnpDeviceInfo = Nothing

        For I = 1 To UBound(UPnPDeviceInfoArray)  ' This is the array of devices discovered by UPnP
            UPnPDeviceUDN = ""
            Try
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DetectUPnPDevices is looking for UPnPDeviceName = " & UPnPDeviceInfoArray(I).UPnPDeviceGivenName & " in UPnPDeviceInfo")
                UPnPDeviceUDN = UPnPDeviceInfoArray(I).UPnPDeviceUDN
            Catch ex As Exception
                Log("Error in DetectUPnPDevices while finding the UPnPDevices with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If UPnPDeviceUDN <> "" Then
                ' go find it in the array
                Try
                    UPnPDeviceInfo = FindUPnPDeviceInfo(UPnPDeviceUDN)
                Catch ex As Exception
                    Log("Error in DetectUPnPDevices while finding the DeviceInfo with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
                Dim InterestedServices As String = ""
                Try
                    InterestedServices = CheckForInterestedServices(UPnPDeviceInfoArray(I).Device)
                    If InterestedServices = "" Then GoTo NextElement
                Catch ex As Exception
                    Log("Error in DetectUPnPDevices while checking for interested Services for NewUDN = " & UPnPDeviceUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    GoTo NextElement
                End Try
                If Not UPnPDeviceInfo Is Nothing Then
                    ' this zone is already known by the plugin
                    ' there is a device code ... use it and put in object
                    Try
                        If UPnPDeviceInfoArray(I).Device IsNot Nothing Then
                            UPnPDeviceInfo.UPnPDeviceOnLine = UPnPDeviceInfoArray(I).Device.Alive 'CheckDeviceIsOnLine(UPnPDeviceInfoArray(I).UPnPDeviceIPAddress)
                        Else
                            UPnPDeviceInfo.UPnPDeviceOnLine = False '  CheckDeviceIsOnLine(UPnPDeviceInfoArray(I).UPnPDeviceIPAddress)
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DetectUPnPDevices found UPnPDeviceName = " & UPnPDeviceInfo.UPnPDeviceGivenName & " without a device ??? UDN = " & UPnPDeviceUDN & " and Type = " & UPnPDeviceInfo.UPnPDeviceDeviceType, LogType.LOG_TYPE_INFO)
                        End If
                        UPnPDeviceInfo.UPnPDeviceModelName = UPnPDeviceInfoArray(I).UPnPDeviceModelName
                        UPnPDeviceInfo.UPnPDeviceModelNumber = UPnPDeviceInfoArray(I).UPnPDeviceModelNumber
                        UPnPDeviceInfo.Device = UPnPDeviceInfoArray(I).Device
                        UPnPDeviceInfo.UPnPDeviceIPAddress = UPnPDeviceInfoArray(I).UPnPDeviceIPAddress
                        UPnPDeviceInfo.UPnPDeviceIPPort = UPnPDeviceInfoArray(I).UPnPDeviceIPPort
                        UPnPDeviceInfo.UPnPDeviceIconURL = UPnPDeviceInfoArray(I).UPnPDeviceIconURL
                        UPnPDeviceInfo.UPnPDeviceManufacturerName = UPnPDeviceInfoArray(I).UPnPDeviceManufacturerName
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DetectUPnPDevices found UPnPDeviceName = " & UPnPDeviceInfo.UPnPDeviceGivenName & " on line = " & UPnPDeviceInfo.UPnPDeviceOnLine.ToString, LogType.LOG_TYPE_INFO)
                        If UPnPDeviceInfo.UPnPDeviceOnLine Then
                            ' the rest we can safely overwrite because the name or IP address could have changed
                            WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceModelName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceModelName)
                            WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diFriendlyName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName)
                            WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diIPAddress.ToString, UPnPDeviceInfoArray(I).UPnPDeviceIPAddress)
                            WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diIPPort.ToString, UPnPDeviceInfoArray(I).UPnPDeviceIPPort)
                            'If Not ImRunningOnLinux Then ' removed v.38
                            Dim MACAddress As String = GetMACAddress(UPnPDeviceInfoArray(I).UPnPDeviceIPAddress).ToString
                            If MACAddress <> "" Then
                                If GetStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, "").ToUpper <> MACAddress.ToUpper Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DetectUPnPDevices. The MAC Address is different. IPAddress = " & UPnPDeviceInfoArray(I).UPnPDeviceIPAddress & " and stored Mac Address = " & GetStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, "") & " and on-line MACAddress = " & MACAddress, LogType.LOG_TYPE_WARNING)
                                    WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, MACAddress)
                                End If
                            End If
                            'End If
                        End If
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceServiceTypes.ToString, InterestedServices) ' store the supported services
                    Catch ex As Exception
                        Log("Error in DetectUPnPDevices while adding the UPnPDevices with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End Try
                Else
                    ' this is a new zone but it could be that it has not been added. Check the ini file
                    Try
                        Dim AnythingExist As String = GetStringIniFile("UPnP Devices UDN to Info", UPnPDeviceInfoArray(I).UPnPDeviceUDN, "")
                        If AnythingExist = "" Then
                            WriteStringIniFile("UPnP Devices UDN to Info", UPnPDeviceInfoArray(I).UPnPDeviceUDN, UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName)
                            WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diGivenName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName)
                            WriteIntegerIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, UPnPDeviceInfoArray(I).UPnPDeviceMusicAPIIndex)
                            WriteBooleanIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diAdminState.ToString, False)
                            WriteBooleanIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
                            WriteIntegerIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, 0)
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DetectUPnPDevices found new UPnPDeviceName = " & UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName & " and On-line Status= " & UPnPDeviceInfoArray(I).UPnPDeviceOnLine.ToString, LogType.LOG_TYPE_INFO)
                        End If
                        ' the rest we can safely overwrite
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceModelName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceModelName)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceType.ToString, UPnPDeviceInfoArray(I).UPnPDeviceDeviceType)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceServiceTypes.ToString, InterestedServices)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diFriendlyName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diIPAddress.ToString, UPnPDeviceInfoArray(I).UPnPDeviceIPAddress)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diIPPort.ToString, UPnPDeviceInfoArray(I).UPnPDeviceIPPort)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceModelNumber.ToString, UPnPDeviceInfoArray(I).UPnPDeviceModelNumber)
                        WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diDeviceManufacturerName.ToString, UPnPDeviceInfoArray(I).UPnPDeviceManufacturerName)
                        'If Not ImRunningOnLinux Then ' removed v.38
                        Dim MACAddress As String = GetMACAddress(UPnPDeviceInfoArray(I).UPnPDeviceIPAddress).ToString
                        If MACAddress <> "" Then
                            If GetStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, "").ToUpper <> MACAddress.ToUpper Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DetectUPnPDevices. The MAC Address is different. IPAddress = " & UPnPDeviceInfoArray(I).UPnPDeviceIPAddress & " and stored Mac Address = " & GetStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, "") & " and on-line MACAddress = " & MACAddress, LogType.LOG_TYPE_WARNING)
                                WriteStringIniFile(UPnPDeviceInfoArray(I).UPnPDeviceUDN, DeviceInfoIndex.diMACAddress.ToString, MACAddress)
                            End If
                        End If
                        'End If
                    Catch ex As Exception
                        Log("Error in DetectUPnPDevices 1 while adding the UPnPDevices with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End Try
                End If
            Else
                Log("Error in DetectUPnPDevices looking for empty UPnPDeviceName/UDN. UPnPDeviceName = " & UPnPDeviceInfoArray(I).UPnPDeviceFriendlyName & ". UPnPDeviceUDN = " & UPnPDeviceUDN & " and Index = " & I.ToString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
NextElement:
        Next
        UPnPDeviceInfoArray = Nothing
        GC.Collect()
    End Sub

    Public Sub FindUPnPDevice(ByRef LocalUPnPDevicesInfo() As MyUPnpDeviceInfo)
        ' This procedure uses SDSS / UPNP to go discover devices/services
        Dim DeviceCount As Integer = UBound(LocalUPnPDevicesInfo)
        Dim DeviceServiceType As String = ""

        Dim UPnPDevicesToGoDiscover As New System.Collections.Generic.Dictionary(Of String, String)()
        UPnPDevicesToGoDiscover = GetIniSection("UPnP Devices to discover") '  As Dictionary(Of String, String)

        If UPnPDevicesToGoDiscover Is Nothing Then
            Log("Error in FindUPnPDevice. No Devices specified in the .ini file under ""UPnP Devices to discover""", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim UPnPDeviceToDiscover As New System.Collections.Generic.KeyValuePair(Of String, String)

        Log("FindUPnPDevice: Attempting to locate all connected devices. This may take up to 9 seconds.", LogType.LOG_TYPE_INFO)
        Dim discoveryPort As Integer = GetIntegerIniFile("Options", "SSDPListenerPort", 0)
        Dim MyDevicesLinkedList As MyUPnPDevices = Nothing
        MyDevicesLinkedList = MySSDPDevice.StartSSDPDiscovery("upnp:rootdevice", discoveryPort)

        If MyDevicesLinkedList Is Nothing Then
            Log("No UPnPDevices found. Please ensure the network is functional and that UPnPDevices devices are attached.", LogType.LOG_TYPE_WARNING)
            'Exit Sub removed else the event handler is never installed!!
        End If

        Log("FindUPnPDevice - Discovery succeeded: " & MyDevicesLinkedList.Count & " UPnPDevice(s) found.", LogType.LOG_TYPE_INFO)

        Try
            Dim NeedsToBeAdded As Boolean = False
            If MyDevicesLinkedList.Count > 0 Then
                For Each MasterDevice As MyUPnPDevice In MyDevicesLinkedList
                    Dim ChildrenProcessed As Boolean = False
                    Dim Device As MyUPnPDevice = MasterDevice
                    Dim ChildDeviceIndex As Integer = 0
                    While Not ChildrenProcessed
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDevice found device = " & Device.Type & " with FriendlyName = " & Device.FriendlyName & " at Location = " & Device.Location, LogType.LOG_TYPE_INFO)
                        NeedsToBeAdded = False
                        For Each UPnPDeviceToDiscover In UPnPDevicesToGoDiscover
                            If Device.Type = UPnPDeviceToDiscover.Key Then
                                NeedsToBeAdded = True
                                DeviceServiceType = UPnPDeviceToDiscover.Value
                                Exit For
                            End If
                        Next
                        If NeedsToBeAdded Then
                            If (Mid(Device.UniqueDeviceName, 1, 12) = "uuid:RINCON_" Or Mid(Device.UniqueDeviceName, 1, 16) = "uuid:DOCKRINCON_") And Not SonosDeviceIn Then
                                ' these are the sonos devices
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDevice found Sonos device with UDN =  " & Device.UniqueDeviceName & " and Friendly Name = " & Device.FriendlyName & " at Location = " & Device.Location, LogType.LOG_TYPE_WARNING)
                            Else
                                Dim NewUDN As String = Replace(Device.UniqueDeviceName, "uuid:", "")
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindUPnPDevice is adding a " & Device.ManufacturerName & " device with UDN = " & NewUDN & " and Friendly Name = " & Device.FriendlyName & " at Location = " & Device.Location & " and adding it to the array with index = " & DeviceCount.ToString, LogType.LOG_TYPE_INFO)
                                DeviceCount = DeviceCount + 1
                                ReDim Preserve LocalUPnPDevicesInfo(DeviceCount)
                                Dim NewDevice As New MyUPnpDeviceInfo
                                LocalUPnPDevicesInfo(DeviceCount) = NewDevice
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceManufacturerName = Device.ManufacturerName
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceUDN = NewUDN
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceFriendlyName = Device.FriendlyName
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceGivenName = Device.FriendlyName ' should be overwritten if in HS database and different
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceDeviceType = DeviceServiceType
                                LocalUPnPDevicesInfo(DeviceCount).Device = Device
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceIPAddress = Device.IPAddress
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceIPPort = Device.IPPort
                                LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceOnLine = True
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            isRoot = " & Device.IsRootDevice.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceisRoot = Device.IsRootDevice
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            Description = " & Device.Description.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceDescription = Device.Description
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            ModelNumber = " & Device.ModelNumber.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceModelNumber = Device.ModelNumber
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            ModelURL = " & Device.ModelURL.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceModelURL = Device.ModelURL
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            PresentationURL = " & Device.PresentationURL.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDevicePresentationURL = Device.PresentationURL
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            hasChildren = " & Device.HasChildren.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceHasChildren = Device.HasChildren
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            ManufacturerURL = " & Device.ManufacturerURL.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceManufacturerURL = Device.ManufacturerURL
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            ModelName = " & Device.ModelName.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceModelName = Device.ModelName
                                Catch ex As Exception
                                End Try
                                Try
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            UPC = " & Device.UPC.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceUPC = Device.UPC
                                Catch ex As Exception
                                End Try
                                Try
                                    Dim ICon As String = ""
                                    ICon = Device.IconURL("image/jpeg", 200, 200, 16)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("'            IconIRL = " & ICon.ToString, LogType.LOG_TYPE_INFO)
                                    LocalUPnPDevicesInfo(DeviceCount).UPnPDeviceIconURL = ICon.ToString
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                        If MasterDevice.HasChildren Then
                            Try
                                Device = MasterDevice.Children(ChildDeviceIndex)
                                ChildDeviceIndex += 1
                            Catch ex1 As Exception
                                ChildrenProcessed = True
                            End Try
                        Else
                            ChildrenProcessed = True
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            Log("ERROR in FindUPnPDevice. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            RemoveHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
        Catch ex As Exception
        End Try
        Try
            RemoveHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
        Catch ex As Exception
        End Try
        Try
            ' RemoveHandler MySSDPDevice.MSearchEvent, AddressOf MSearchEvent
        Catch ex As Exception
        End Try
        Try
            AddHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
        Catch ex As Exception
            Log("ERROR in FindUPnPDevice trying to add a NewDeviceFound Handler with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            AddHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
        Catch ex As Exception
            Log("ERROR in FindUPnPDevice trying to add a MulticastDied Event Handler with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'AddHandler MySSDPDevice.MSearchEvent, AddressOf MSearchEvent
        Catch ex As Exception
        End Try
    End Sub

    Public Sub NewDeviceFound(inUDN As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NewDeviceFound called for device = " & inUDN, LogType.LOG_TYPE_INFO, LogColorNavy)
        SyncLock (MyNewDiscoveredDeviceQueue)
            MyNewDiscoveredDeviceQueue.Enqueue(inUDN)
        End SyncLock
        Try
            If MyAddNewDeviceTimer Is Nothing Then
                MyAddNewDeviceTimer = New Timers.Timer
            End If
            MyAddNewDeviceTimer.Stop()              ' if the timer was already running, this will stop it and restart it
            MyAddNewDeviceTimer.Interval = WaitTimeToAddNewDevice    ' wait 10 seconds for the device to come on-line properly
            MyAddNewDeviceTimer.AutoReset = False
            MyAddNewDeviceTimer.Enabled = True
            MyAddNewDeviceTimer.Start()
        Catch ex As Exception
            If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in NewDeviceFound for inUDN = " & inUDN & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub AddNewDiscoveredDevice()
        If NewDeviceHandlerReEntryFlag Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice has Re-Entry while processing Notification queue with # elements = " & MyNewDiscoveredDeviceQueue.Count.ToString, LogType.LOG_TYPE_WARNING)
            'MissedNewDeviceNotificationHandlerFlag = True
            Try
                If MyAddNewDeviceTimer Is Nothing Then
                    MyAddNewDeviceTimer = New Timers.Timer
                End If
                MyAddNewDeviceTimer.Stop()              ' if the timer was already running, this will stop it and restart it
                MyAddNewDeviceTimer.Interval = WaitTimeToAddNewDevice    ' wait 10 seconds for the device to come on-line properly
                MyAddNewDeviceTimer.AutoReset = False
                MyAddNewDeviceTimer.Enabled = True
                MyAddNewDeviceTimer.Start()
            Catch ex As Exception
                If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in AddNewDiscoveredDevice with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        NewDeviceHandlerReEntryFlag = True
        Dim UPnPDevicesToGoDiscover As New System.Collections.Generic.Dictionary(Of String, String)()
        UPnPDevicesToGoDiscover = GetIniSection("UPnP Devices to discover") '  As Dictionary(Of String, String)

        If UPnPDevicesToGoDiscover Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddNewDiscoveredDevice. No Devices specified in the .ini file under ""UPnP Devices to discover""", LogType.LOG_TYPE_ERROR)
            NewDeviceHandlerReEntryFlag = False
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice is processing Notification queue with # elements = " & MyNewDiscoveredDeviceQueue.Count.ToString, LogType.LOG_TYPE_INFO)
        Dim NewUDN As String
        Dim UPnPDeviceInfoArray As MyUPnpDeviceInfo() = Nothing
        ReDim UPnPDeviceInfoArray(0)
        Try
            While MyNewDiscoveredDeviceQueue.Count > 0
                SyncLock (MyNewDiscoveredDeviceQueue)
                    NewUDN = MyNewDiscoveredDeviceQueue.Dequeue
                End SyncLock
                If NewUDN <> "" Then
                    Dim UPnPDeviceInfo As MyUPnpDeviceInfo = Nothing
                    Dim NewUPnPDevice As MyUPnPDevice = MySSDPDevice.Item("uuid:" & NewUDN)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice dequeued UDN = " & NewUDN, LogType.LOG_TYPE_INFO)
                    If NewUPnPDevice Is Nothing Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice dequeued UDN = " & NewUDN & " but found no UPNPDevice", LogType.LOG_TYPE_WARNING)
                        GoTo NextElement
                    End If
                    Dim NeedsToBeAdded As Boolean = False
                    Dim ServiceType As String = "'"
                    If Not (Mid(NewUPnPDevice.UniqueDeviceName, 1, 12) = "uuid:RINCON_" Or Mid(NewUPnPDevice.UniqueDeviceName, 1, 16) = "uuid:DOCKRINCON_") Or SonosDeviceIn Then
                        For Each UPnPDeviceToDiscover In UPnPDevicesToGoDiscover
                            If NewUPnPDevice.Type = UPnPDeviceToDiscover.Key Then
                                NeedsToBeAdded = True
                                ServiceType = UPnPDeviceToDiscover.Value
                                Exit For
                            End If
                        Next
                    End If
                    If Not NeedsToBeAdded Then GoTo NextElement
                    Try
                        UPnPDeviceInfo = FindUPnPDeviceInfo(NewUDN)
                    Catch ex As Exception
                        Log("Error in AddNewDiscoveredDevice while finding the DeviceInfo NewUDN = " & NewUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        GoTo NextElement
                    End Try
                    Dim InterestedServices As String = ""
                    Try
                        InterestedServices = CheckForInterestedServices(NewUPnPDevice)
                        If InterestedServices = "" Then GoTo NextElement
                    Catch ex As Exception
                        Log("Error in AddNewDiscoveredDevice while checking for interested Services for NewUDN = " & NewUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        GoTo NextElement
                    End Try
                    If UPnPDeviceInfo IsNot Nothing Then
                        Try ' this was a known device that most likely just came on-line
                            UPnPDeviceInfo.NewDevice = True
                            UPnPDeviceInfo.UPnPDeviceOnLine = NewUPnPDevice.Alive
                            UPnPDeviceInfo.UPnPDeviceModelName = NewUPnPDevice.ModelName
                            UPnPDeviceInfo.UPnPDeviceModelNumber = NewUPnPDevice.ModelNumber
                            UPnPDeviceInfo.Device = NewUPnPDevice
                            UPnPDeviceInfo.UPnPDeviceIPAddress = NewUPnPDevice.IPAddress
                            UPnPDeviceInfo.UPnPDeviceIPPort = NewUPnPDevice.IPPort
                            UPnPDeviceInfo.UPnPDeviceIconURL = NewUPnPDevice.IconURL("image/jpeg", 200, 200, 16)
                            UPnPDeviceInfo.UPnPDeviceManufacturerName = NewUPnPDevice.ManufacturerName
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice found UPnPDeviceName = " & UPnPDeviceInfo.UPnPDeviceGivenName & " on line", LogType.LOG_TYPE_INFO)
                            If UPnPDeviceInfo.UPnPDeviceOnLine Then
                                ' the rest we can safely overwrite because the name or IP address could have changed
                                WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceModelName.ToString, UPnPDeviceInfo.UPnPDeviceModelName)
                                WriteStringIniFile(NewUDN, DeviceInfoIndex.diFriendlyName.ToString, NewUPnPDevice.FriendlyName)
                                WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPAddress.ToString, UPnPDeviceInfo.UPnPDeviceIPAddress)
                                WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPPort.ToString, UPnPDeviceInfo.UPnPDeviceIPPort)
                                'If Not ImRunningOnLinux Then removed v.38
                                Dim MACAddress As String = GetMACAddress(UPnPDeviceInfo.UPnPDeviceIPAddress).ToString
                                If MACAddress <> "" Then
                                    If GetStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, "").ToUpper <> MACAddress.ToUpper Then
                                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DetectUPnPDevices. The MAC Address is different. IPAddress = " & UPnPDeviceInfo.UPnPDeviceIPAddress & " and stored Mac Address = " & GetStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, "") & " and on-line MACAddress = " & MACAddress, LogType.LOG_TYPE_WARNING)
                                        WriteStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, MACAddress)
                                    End If
                                End If
                                'End If
                            End If
                        Catch ex As Exception
                            Log("Error in AddNewDiscoveredDevice while overwriting an existing UPnPDevice with NewUDN = " & NewUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Try
                        End Try
                    Else
                        ' this is a new zone but it could be that it has not been added. Check the ini file
                        Try
                            Dim AnythingExist As String = GetStringIniFile("UPnP Devices UDN to Info", NewUDN, "")
                            If AnythingExist = "" Then
                                WriteStringIniFile("UPnP Devices UDN to Info", NewUDN, NewUPnPDevice.FriendlyName)
                                WriteStringIniFile(NewUDN, DeviceInfoIndex.diGivenName.ToString, NewUPnPDevice.FriendlyName)
                                WriteBooleanIniFile(NewUDN, DeviceInfoIndex.diAdminState.ToString, False)
                                WriteBooleanIniFile(NewUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice has logged new device with UPnPDeviceName = " & NewUPnPDevice.FriendlyName, LogType.LOG_TYPE_INFO)
                            End If
                            ' the rest we can safely overwrite
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceModelName.ToString, NewUPnPDevice.ModelName)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceServiceTypes.ToString, InterestedServices)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceType.ToString, ServiceType)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diFriendlyName.ToString, NewUPnPDevice.FriendlyName)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPAddress.ToString, NewUPnPDevice.IPAddress)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPPort.ToString, NewUPnPDevice.IPPort)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceModelNumber.ToString, NewUPnPDevice.ModelNumber)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceManufacturerName.ToString, NewUPnPDevice.ManufacturerName)
                            'If Not ImRunningOnLinux Then ' removed v.38
                            Dim MACAddress As String = GetMACAddress(NewUPnPDevice.IPAddress).ToString
                            If MACAddress <> "" Then
                                If GetStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, "").ToUpper <> MACAddress.ToUpper Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DetectUPnPDevices. The MAC Address is different. IPAddress = " & NewUPnPDevice.IPAddress & " and stored Mac Address = " & GetStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, "") & " and on-line MACAddress = " & MACAddress, LogType.LOG_TYPE_WARNING)
                                    WriteStringIniFile(NewUDN, DeviceInfoIndex.diMACAddress.ToString, MACAddress)
                                End If
                            End If
                            'End If

                        Catch ex As Exception
                            Log("Error in AddNewDiscoveredDevice while adding the UPnPDevice with NewUDN = " & NewUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Try
                        End Try
                    End If
                End If
NextElement:
            End While
        Catch ex As Exception

        End Try
        UPnPDeviceInfoArray = Nothing
        GC.Collect()
        'If MissedNewDeviceNotificationHandlerFlag Then AddDeviceFlag = True ' rearm the timer to prevent events from getting lost
        'MissedNewDeviceNotificationHandlerFlag = False
        If MyNewDiscoveredDeviceQueue.Count > 0 Then
            Try
                If MyAddNewDeviceTimer Is Nothing Then
                    MyAddNewDeviceTimer = New Timers.Timer
                End If
                MyAddNewDeviceTimer.Stop()              ' if the timer was already running, this will stop it and restart it
                MyAddNewDeviceTimer.Interval = WaitTimeToAddNewDevice    ' wait 10 seconds for the device to come on-line properly
                MyAddNewDeviceTimer.AutoReset = False
                MyAddNewDeviceTimer.Enabled = True
                MyAddNewDeviceTimer.Start()
            Catch ex As Exception
                If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in AddNewDiscoveredDevice with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        NewDeviceHandlerReEntryFlag = False
    End Sub

    Public Sub MultiCastDiedEvent()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error. MultiCastDiedEvent received. Terminating the PI to try to restart it", LogType.LOG_TYPE_ERROR)
        ShutdownIO()
    End Sub

    Private Function FindUPnPDeviceInfo(ByVal UDN As String) As MyUPnpDeviceInfo
        FindUPnPDeviceInfo = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning in FindUPnPDeviceInfo for UDN = " & UDN & ". The array does not exist ", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.UPnPDeviceUDN = UDN Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDeviceInfo found UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_INFO)
                    FindUPnPDeviceInfo = HSDevice
                    Exit Function
                Else
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDeviceInfo did not find UDN = " & UDN & " but found = " & HSDevice.UPnPDeviceUDN, LogType.LOG_TYPE_INFO)
                End If
            Next
        Catch ex As Exception
            Log("Error in FindUPnPDeviceInfo Finding UPnPDevicveInfo. UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning in FindUPnPDeviceInfo did not find UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_WARNING)
    End Function

    Private Function CheckForInterestedServices(NewDevice As MyUPnPDevice) As String
        CheckForInterestedServices = ""
        If NewDevice Is Nothing Then Exit Function
        ' Get the "serviceList", iterate for "service" and check the "serviceType" for interested services
        If NewDevice.Type = "urn:samsung.com:device:RemoteControlReceiver:1" Then   ' added in V24, some Samsung TVs do not have the service urn:samsung.com:serviceId:TestRCRService

            CheckForInterestedServices = "RCR"
        End If
        Dim Services As MyUPnPServices = Nothing
        Try
            Services = NewDevice.Services
            If Services Is Nothing Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForInterestedServices for device = " & NewDevice.FriendlyName & " found no services", LogType.LOG_TYPE_INFO)
                Exit Function
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForInterestedServices for device = " & NewDevice.FriendlyName & " found " & NewDevice.Services.Count.ToString & " Services", LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Try
                    For Each objService As MyUPnPService In Services
                        If objService IsNot Nothing Then
                            Log("CheckForInterestedServices for device = " & NewDevice.FriendlyName & " found Service ID = " & objService.Id.ToString, LogType.LOG_TYPE_INFO)
                        End If
                    Next
                Catch ex As Exception
                    Log("Error in CheckForInterestedServices for device = " & NewDevice.FriendlyName & " discovering the Service IDs with error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            Try
                Dim ServicesToGoDiscover As New System.Collections.Generic.Dictionary(Of String, String)()
                ServicesToGoDiscover = GetIniSection("UPnP Services to discover") '  As Dictionary(Of String, String)
                If ServicesToGoDiscover Is Nothing Then
                    Log("Error in CheckForInterestedServices. No Services specified in the .ini file under ""UPnP Services to discover""", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                For Each objService As MyUPnPService In Services
                    Dim ObjectserviceID As String = ""
                    Dim objectserviceType As String = ""
                    Try
                        If Not objService Is Nothing Then
                            ObjectserviceID = objService.Id
                            objectserviceType = objService.ServiceType
                        End If
                    Catch ex As Exception
                        Log("Error in CheckForInterestedServices for device = " & NewDevice.FriendlyName & " extracting the Service ID with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    ' Look for [upnp Services to discover]
                    Dim NeedsToBeAdded As Boolean = False
                    Dim ServiceType As String = ""
                    For Each UPnPServiceToDiscover In ServicesToGoDiscover
                        If ObjectserviceID = UPnPServiceToDiscover.Key Then
                            NeedsToBeAdded = True
                            ServiceType = UPnPServiceToDiscover.Value
                            Exit For
                        End If
                    Next
                    If CheckForInterestedServices.IndexOf(ServiceType) = -1 Then
                        ' add it to the list of services
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForInterestedServices for device = " & NewDevice.FriendlyName & " added ServiceType = " & ServiceType, LogType.LOG_TYPE_INFO)
                        If CheckForInterestedServices <> "" Then
                            CheckForInterestedServices &= "," & ServiceType
                        Else
                            CheckForInterestedServices = ServiceType
                        End If
                    End If
                Next
            Catch ex As Exception
                Log("Error in CheckForInterestedServices for device = " & NewDevice.FriendlyName & " going through the services with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("Error in CheckForInterestedServices for device = " & NewDevice.FriendlyName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub CreateUPnPControllers(ActivateTheZone As Boolean)
        Dim UPnPDevice As HSPI

        If MyHSDeviceLinkedList.Count = 0 Then
            Log("CreateUPnPControllers called but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateUPnPControllers found " & MyHSDeviceLinkedList.Count.ToString & " devices and ActivateTheZone = " & ActivateTheZone.ToString, LogType.LOG_TYPE_INFO)
        Dim MusicAPIIndex As Integer = 0

        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.UPnPDeviceControllerRef Is Nothing Then ' no instance of HSPI exist yet. Instantiate it and add it
                Try
                    If MainInstance <> "" Then ' this allows multiple copies of the PI to exist. Normally the MainInstance is empty
                        UPnPDevice = AddInstance(MainInstance & "-" & HSDevice.UPnPDeviceUDN)
                    Else
                        UPnPDevice = AddInstance(HSDevice.UPnPDeviceUDN)
                    End If
                    HSDevice.UPnPDeviceControllerRef = UPnPDevice
                    Thread.Sleep(1000)    ' takes some time for new instance to connect and callbacks for registering pages to complete
                Catch ex As Exception
                    Log("Error in CreateUPnPControllers. Could not instantiate UPnPDeviceController with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit For
                End Try
                Try
                    ' pass all the info that was retrieved from the HS DB (BuildUPnPDeviceInfoList) and discovered (DetectUPnPDevices) in to the instance of HSPI
                    UPnPDevice.DeviceName = HSDevice.UPnPDeviceGivenName
                    UPnPDevice.DeviceUDN = HSDevice.UPnPDeviceUDN
                    UPnPDevice.DeviceServiceType = HSDevice.UPnPDeviceDeviceType
                    UPnPDevice.DeviceHSRef = HSDevice.UPnPDeviceHSRef
                    UPnPDevice.DeviceIPAddress = HSDevice.UPnPDeviceIPAddress
                    UPnPDevice.DeviceIPPort = HSDevice.UPnPDeviceIPPort
                    UPnPDevice.DeviceModelName = HSDevice.UPnPDeviceModelName
                    UPnPDevice.DeviceModelNumber = HSDevice.UPnPDeviceModelNumber
                    UPnPDevice.DeviceManufacturer = HSDevice.UPnPDeviceManufacturerName
                    UPnPDevice.PlayerIconURL = HSDevice.UPnPDeviceIconURL
                    UPnPDevice.ReferenceToController = Me
                    UPnPDevice.DeviceAPIIndex = GetNextFreeDeviceIndex()
                    UPnPDevice.DeviceAdminStateActive = GetBooleanIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diAdminState.ToString, False)
                    UPnPDevice.ReadDeviceIniSettings()
                    UPnPDevice.pDevice = HSDevice.Device
                    If HSDevice.UPnPDeviceDeviceType = "DMR" Or HSDevice.UPnPDeviceDeviceType = "HST" Then
                        UPnPDevice.LoadCurrentPlaylistTracks()
                        ' we publish a MUSIC API for this
                        MusicAPIIndex = HSDevice.UPnPDeviceMusicAPIIndex
                        If MusicAPIIndex = 0 Then
                            MusicAPIIndex = GetNextFreeMusicAPIIndex()
                        End If
                        HSDevice.UPnPDeviceMusicAPIIndex = MusicAPIIndex
                        UPnPDevice.APIInstance = MusicAPIIndex
                    End If
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateUPnPControllers has found UPnpDevice = " & HSDevice.UPnPDeviceGivenName & " with on-line status = " & HSDevice.UPnPDeviceOnLine.ToString & " and Adminstate = " & GetBooleanIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diAdminState.ToString, False).ToString, LogType.LOG_TYPE_INFO)

                    If HSDevice.UPnPDeviceOnLine And GetBooleanIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diAdminState.ToString, False) And Not (HSDevice.UPnPDeviceDeviceType = "HST") Then
                        UPnPDevice.DirectConnect(HSDevice.Device)
                    ElseIf GetBooleanIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diAdminState.ToString, False) And HSDevice.UPnPDeviceDeviceType = "HST" Then
                        Dim HSRef = GetIntegerIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diPlayerHSRef.ToString, -1)
                        If HSRef = -1 Then
                            HSRef = UPnPDevice.CreateHSPlayerDevice(HSRef, UPnPDevice.DeviceName, True)
                            If HSRef <> -1 Then
                                WriteIntegerIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diPlayerHSRef.ToString, HSRef)
                            End If
                        Else
                            UPnPDevice.CreateHSPlayerDevice(HSRef, UPnPDevice.DeviceName, False)
                        End If
                    Else
                        ' the Device was not on-line. set the last known IPAddress/Port and see if the timer routine in the instance detects it coming on-line
                        If PIDebuglevel > DebugLevel.dlEvents And Not (HSDevice.UPnPDeviceDeviceType = "HST" Or HSDevice.UPnPDeviceDeviceType = "DIAL") Then Log("Warning in CreateUPnPControllers. UPnpDevice = " & HSDevice.UPnPDeviceGivenName & " not on-line. Using last known IPAdress/Port", LogType.LOG_TYPE_WARNING)
                        UPnPDevice.DeviceIPAddress = HSDevice.UPnPDeviceIPAddress
                        UPnPDevice.DeviceIPPort = HSDevice.UPnPDeviceIPPort
                    End If
                    ' Update the info like musicAPI & DeviceIndex
                    WriteIntegerIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, HSDevice.UPnPDeviceMusicAPIIndex)
                    WriteBooleanIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, True)
                    If ActivateTheZone Or UPnPDevice.DeviceAdminStateActive Then UPnPDevice.SetAdministrativeState(True)
                    WriteIntegerIniFile(HSDevice.UPnPDeviceUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, UPnPDevice.DeviceAPIIndex)
                    UPnPDevice.SetHSMainState()
                Catch ex As Exception
                    Log("Error in CreateUPnPControllers. Could not create instance of UPnPDeviceController for UPnpDevice = " & HSDevice.UPnPDeviceGivenName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Log("CreateUPnPControllers created an instance of UPnPDeviceController for UPnPDevice = " & HSDevice.UPnPDeviceGivenName, LogType.LOG_TYPE_INFO)
            Else
                ' how do we find out which devices came on-line !!!
                ' we need to check whether something changed like IP address
                UPnPDevice = HSDevice.UPnPDeviceControllerRef
                UPnPDevice.pDevice = HSDevice.Device
                If UPnPDevice.DeviceIPAddress <> HSDevice.UPnPDeviceIPAddress Or UPnPDevice.DeviceIPPort <> HSDevice.UPnPDeviceIPPort Then
                    ' if this happens, we need to remove IP address to pinglist and add new
                    UPnPDevice.DeviceIPAddress = HSDevice.UPnPDeviceIPAddress
                    UPnPDevice.DeviceIPPort = HSDevice.UPnPDeviceIPPort
                    'UPnPDevice.DeviceDocumentURL = HSDevice.UPnPDocumentURL
                    UPnPDevice.PlayerIconURL = HSDevice.UPnPDeviceIconURL
                    UPnPDevice.DeviceModelName = HSDevice.UPnPDeviceModelName
                    UPnPDevice.DeviceModelNumber = HSDevice.UPnPDeviceModelNumber
                    UPnPDevice.DeviceManufacturer = HSDevice.UPnPDeviceManufacturerName
                End If
                If HSDevice.NewDevice Then
                    HSDevice.NewDevice = False
                    If ActivateTheZone Or UPnPDevice.DeviceAdminStateActive Then UPnPDevice.SetAdministrativeState(True)
                    UPnPDevice.SetHSMainState()
                End If
            End If
            If (HSDevice.UPnPDeviceDeviceType = "DMR" Or HSDevice.UPnPDeviceDeviceType = "HST") And Not HSDevice.UPnPDeviceWeblinkCreated Then
                UPnPDevice.CreateWebLink(HSDevice.UPnPDeviceGivenName, HSDevice.UPnPDeviceUDN)
                HSDevice.UPnPDeviceWeblinkCreated = True
            ElseIf HSDevice.UPnPDeviceDeviceType = "PMR" And Not HSDevice.UPnPDeviceWeblinkCreated Then
                UPnPDevice.CreateConfigLink(HSDevice.UPnPDeviceGivenName, HSDevice.UPnPDeviceUDN)
                HSDevice.UPnPDeviceWeblinkCreated = True
            End If
        Next
    End Sub

    Private Function ExtractIPInfo(DocumentURL As String) As IPAddressInfo
        Dim HttpIndex As Integer = DocumentURL.ToUpper.IndexOf("HTTP://")
        ExtractIPInfo.IPAddress = ""
        ExtractIPInfo.IPPort = ""
        Try
            If HttpIndex = -1 Then
                Log("ERROR in ExtractIPInfo. Not HTTP:// found. URL = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SubStr As String
            SubStr = DocumentURL.Substring(HttpIndex + 7, DocumentURL.Length - HttpIndex - 7)
            ' substring should now be primed for an IP address in the form of 192.168.1.1
            ' The next forward slash marks the end of the IP address, this could include the Port!
            Dim SlashIndex As Integer = SubStr.IndexOf("/")
            If SlashIndex <> -1 Then
                SubStr = SubStr.Substring(0, SlashIndex)
            End If
            SubStr = SubStr.Trim
            If SubStr = "" Then
                Log("ERROR in ExtractIPInfo. No IP address found = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SemiCollonIndex As Integer = SubStr.IndexOf(":")
            'log( "ExtractIPInfo has Substring = " & SubStr)
            If SemiCollonIndex <> -1 Then
                ' there is an IP address and a Port Number
                ExtractIPInfo.IPAddress = SubStr.Substring(0, SemiCollonIndex)
                ExtractIPInfo.IPPort = SubStr.Substring(SemiCollonIndex + 1, SubStr.Length - SemiCollonIndex - 1)
            Else
                ' only IP address
                ExtractIPInfo.IPAddress = SubStr
            End If
        Catch ex As Exception
            Log("ERROR in ExtractIPInfo with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub ReadIniFile()

        PIDebuglevel = GetIntegerIniFile("Options", "PIDebugLevel", DebugLevel.dlErrorsOnly)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadIniFile called", LogType.LOG_TYPE_INFO)
        UPnPDebuglevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlOff)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadIniFile called", LogType.LOG_TYPE_INFO)

        Try
            If GetStringIniFile("Options", "MaxNbrofUPNPObjects", "") <> "" Then
                MaxNbrOfUPNPObjects = GetIntegerIniFile("Options", "MaxNbrofUPNPObjects", -1)
                If MaxNbrOfUPNPObjects = -1 Then
                    MaxNbrOfUPNPObjects = cMaxNbrOfUPNP ' = maximum
                End If
                If PIDebuglevel > DebugLevel.dlEvents Then Log("INIT: MaxNbrOfUPNPObjects set to " & MaxNbrOfUPNPObjects, LogType.LOG_TYPE_INFO)
            Else
                WriteIntegerIniFile("Options", "MaxNbrofUPNPObjects", cMaxNbrOfUPNP)
                'MaxNbrOfUPNPObjects = 0
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading MaxNbrOfUPNPObjects with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            PreviousVersion = GetIntegerIniFile("Options", "PreviousVersion", 0)
        Catch ex As Exception
            Log("Error in ReadIniFile reading PreviousVersion with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "NbrOfPingRetries", "") = "" Then
                WriteIntegerIniFile("Options", "NbrOfPingRetries", 3)
                NbrOfPingRetries = 3
            Else
                NbrOfPingRetries = GetIntegerIniFile("Options", "NbrOfPingRetries", 0)
                If NbrOfPingRetries = 0 Then
                    NbrOfPingRetries = 3
                    WriteIntegerIniFile("Options", "NbrOfPingRetries", 3)
                End If
            End If
            If GetStringIniFile("Options", "ShowFailedPings", "") = "" Then
                WriteBooleanIniFile("Options", "ShowFailedPings", False)
            Else
                ShowFailedPings = GetBooleanIniFile("Options", "ShowFailedPings", True)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Ping info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        SetSpeakerProxy() ' changed in v28
        Try
            If GetStringIniFile("Options", "ArtworkHSize", "") = "" Then
                WriteIntegerIniFile("Options", "ArtworkHSize", 150)
            Else
                ArtworkHSize = GetIntegerIniFile("Options", "ArtworkHSize", 150)
            End If
            If GetStringIniFile("Options", "ArtworkVsize", "") = "" Then
                WriteIntegerIniFile("Options", "ArtworkVsize", 150)
            Else
                ArtworkVSize = GetIntegerIniFile("Options", "ArtworkVsize", 150)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Artwork Size info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MyVolumeStep = GetIntegerIniFile("Options", "VolumeStep", 0)
            If MyVolumeStep = 0 Then
                MyVolumeStep = 5
                WriteIntegerIniFile("Options", "VolumeStep", 5)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading VolumeStep with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            WriteStringIniFile("UPnP Devices to discover", "urn:schemas-upnp-org:device:MediaRenderer:1", "DMR")
            WriteStringIniFile("UPnP Devices to discover", "urn:schemas-upnp-org:device:MediaRenderer:2", "DMR")
            WriteStringIniFile("UPnP Devices to discover", "urn:schemas-upnp-org:device:MediaServer:1", "DMS")
            WriteStringIniFile("UPnP Devices to discover", "urn:schemas-upnp-org:device:MediaServer:2", "DMS")
            WriteStringIniFile("UPnP Devices to discover", "urn:samsung.com:device:RemoteControlReceiver:1", "RCR")
            WriteStringIniFile("UPnP Devices to discover", "urn:samsung.com:device:PersonalMessageReceiver:1", "PMR")
            WriteStringIniFile("UPnP Devices to discover", "urn:roku-com:device:player:1-0", "DIAL")
            WriteStringIniFile("UPnP Devices to discover", "urn:roku-com:service:ecp:1", "DIAL")
            WriteStringIniFile("UPnP Devices to discover", "urn:schemas-upnp-org:device:Basic:1", "RCR")
            WriteStringIniFile("UPnP Devices to discover", "urn:dial-multiscreen-org:device:dialreceiver:1", "DIAL") ' removed this for dcortizen
            'WriteStringIniFile("UPnP Devices to discover", "urn:dial-multiscreen-org:device:dialreceiver:1", "RCR") 'add this for dcortizen
            WriteStringIniFile("UPnP Devices to discover", "urn:dial-multiscreen-org:device:dial:1", "DIAL") ' 
            WriteStringIniFile("UPnP Devices to discover", "urn:dial-multiscreen-org:service:dial:1", "RCR") ' LG WebOS TV
            WriteStringIniFile("UPnP Devices to discover", "urn:samsung.com:device:IPControlServer:1", "RCR") ' Samsung Y2018 Q series
        Catch ex As Exception
            Log("Error in ReadIniFile finding/defining the devices to discover with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            WriteStringIniFile("UPnP Services to discover", "urn:upnp-org:serviceId:AVTransport", "DMR")
            WriteStringIniFile("UPnP Services to discover", "urn:schemas-upnp-org:service:AVTransport", "DMR")
            WriteStringIniFile("UPnP Services to discover", "urn:upnp-org:serviceId:RenderingControl", "DMR")
            WriteStringIniFile("UPnP Services to discover", "urn:schemas-upnp-org:service:RenderingControl", "DMR")
            WriteStringIniFile("UPnP Services to discover", "urn:upnp-org:serviceId:ContentDirectory", "DMS")
            WriteStringIniFile("UPnP Services to discover", "urn:samsung.com:serviceId:TestRCRService", "RCR")
            WriteStringIniFile("UPnP Services to discover", "urn:schemas-sony-com:serviceId:IRCC", "RCR")
            WriteStringIniFile("UPnP Services to discover", "urn:samsung.com:serviceId:MessageBoxService", "PMR")
            WriteStringIniFile("UPnP Services to discover", "urn:dial-multiscreen-org:service:dial", "DIAL") ' removed this for dcortizen
            'WriteStringIniFile("UPnP Services to discover", "urn:dial-multiscreen-org:service:dial", "RCR") ' add this for dcortizen
            WriteStringIniFile("UPnP Services to discover", "urn:roku-com:serviceId:ecp1-0", "RCR")
            WriteStringIniFile("UPnP Services to discover", "urn:dial-multiscreen-org:serviceId:dial1-0", "DIAL")
            WriteStringIniFile("UPnP Services to discover", "urn:dial-multiscreen-org:serviceId:dial", "DIAL") ' removed this for dcortizen
            'WriteStringIniFile("UPnP Services to discover", "urn:dial-multiscreen-org:serviceId:dial", "RCR") ' add this for dcortizen
            WriteStringIniFile("UPnP Services to discover", "urn:lge-com:serviceId:webos-second-screen-3000-3001", "RCR") ' LG WebOS TV
            WriteStringIniFile("UPnP Services to discover", "urn:samsung.com:serviceId:IPControlService", "RCR") ' Samsung Y2018 Q series
        Catch ex As Exception
            Log("Error in ReadIniFile finding/defining the Services to discover with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "NoPinging", "") <> "" Then
                MyNoPingingFlag = GetBooleanIniFile("Options", "NoPinging", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading NoPinging flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "PostAnnouncementAction", "") = "" Then
                WriteIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
            End If
            MyPostAnnouncementAction = GetIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
        Catch ex As Exception
            Log("Error in ReadIniFile reading PostAnnouncementAction with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "HSSTrackLengthSetting", "") = "" Then
                WriteIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSHoursMinutesSeconds)
            End If
            MyHSTrackLengthFormat = GetIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSHoursMinutesSeconds)
        Catch ex As Exception
            Log("Error in ReadIniFile reading HS Device Tracklengh Format Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "HSSTrackPositionSetting", "") = "" Then
                WriteIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSHoursMinutesSeconds)
            End If
            MyHSTrackPositionFormat = GetIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSHoursMinutesSeconds)
        Catch ex As Exception
            Log("Error in ReadIniFile reading HS Device Trackposition Format Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "UPnPSubscribeTimeOut", "") = "" Then
                WriteIntegerIniFile("Options", "UPnPSubscribeTimeOut", UPnPSubscribeTimeOut)
            End If
            UPnPSubscribeTimeOut = GetIntegerIniFile("Options", "UPnPSubscribeTimeOut", UPnPSubscribeTimeOut)
        Catch ex As Exception
            Log("Error in ReadIniFile reading UPnPSubscribeTimeOut Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "UPnPAddNewDeviceWaitTime", "") = "" Then
                WriteIntegerIniFile("Options", "UPnPAddNewDeviceWaitTime", WaitTimeToAddNewDevice)
            End If
            WaitTimeToAddNewDevice = GetIntegerIniFile("Options", "UPnPAddNewDeviceWaitTime", WaitTimeToAddNewDevice)
        Catch ex As Exception
            Log("Error in ReadIniFile reading WaitTimeToAddNewDevice Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub SetSpeakerProxy()
        If Not MyPIisInitialized Then Exit Sub
        Try
            If GetStringIniFile("SpeakerProxy", "Active", "") = "" Then
                WriteBooleanIniFile("SpeakerProxy", "Active", True)
                ActAsSpeakerProxy = True
            Else
                ActAsSpeakerProxy = GetBooleanIniFile("SpeakerProxy", "Active", False)
            End If
        Catch ex As Exception
            Log("Error in SetSpeakerProxy reading Speakerproxy flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If ProxySpeakerActive And Not ActAsSpeakerProxy Then
                callback.UnRegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                Log("Deactivated SpeakerProxy", LogType.LOG_TYPE_INFO)
                ProxySpeakerActive = False
            ElseIf Not ProxySpeakerActive And ActAsSpeakerProxy Then
                callback.RegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                Log("Registered SpeakerProxy", LogType.LOG_TYPE_INFO)
                ProxySpeakerActive = True
            End If
        Catch ex As Exception
            Log("Error in SetSpeakerProxy registering/unregistering Speakerproxy with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub AddDevicetoHS(DeviceUDN As String) ' this is where we add the root device for any DLNA device being added
        Log("AddDevicetoHS called with DeviceUDN = " & DeviceUDN.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim UPnPDeviceInfo As MyUPnpDeviceInfo = DMAdd() ' add it to the internal Device Info Array that the main uses
            If UPnPDeviceInfo Is Nothing Then
                ' this should realy not happen
                Log("Error in AddDevicetoHS. Unable to add an entry to the HSDeviceInfo array for UDN = " & DeviceUDN, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            UPnPDeviceInfo.Device = MySSDPDevice.Item("uuid:" & DeviceUDN)
            UPnPDeviceInfo.UPnPDeviceGivenName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diGivenName.ToString, "")
            UPnPDeviceInfo.UPnPDeviceUDN = DeviceUDN
            UPnPDeviceInfo.UPnPDeviceModelName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceModelName.ToString, "")
            UPnPDeviceInfo.UPnPDeviceDeviceType = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "")
            UPnPDeviceInfo.UPnPDeviceServiceTypes = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceServiceTypes.ToString, "")
            UPnPDeviceInfo.UPnPDeviceFriendlyName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diFriendlyName.ToString, "")
            UPnPDeviceInfo.UPnPDeviceIPAddress = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diIPAddress.ToString, "")
            UPnPDeviceInfo.UPnPDeviceIPPort = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diIPPort.ToString, "")
            UPnPDeviceInfo.UPnPDeviceMusicAPIIndex = GetIntegerIniFile(DeviceUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, 0)
            UPnPDeviceInfo.UPnPDeviceIconURL = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIConURL.ToString, "") ' Needed for Pictureshow devices where this is set in CreateSlideshowDevice
            UPnPDeviceInfo.UPnPDeviceModelNumber = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceModelNumber.ToString, "")
            UPnPDeviceInfo.UPnPDeviceManufacturerName = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceManufacturerName.ToString, "")
            If UPnPDeviceInfo.Device IsNot Nothing Then
                UPnPDeviceInfo.UPnPDeviceOnLine = UPnPDeviceInfo.Device.Alive
            Else
                UPnPDeviceInfo.UPnPDeviceOnLine = False ' CheckDeviceIsOnLine(UPnPDeviceInfo.UPnPDeviceIPAddress)
            End If
            WriteBooleanIniFile(DeviceUDN, DeviceInfoIndex.diAdminState.ToString, True) ' has to be set to true else we'll never be able to create the HS devices to activate/deactive
            ' here we need to create the HSRootDevice
            UPnPDeviceInfo.UPnPDeviceHSRef = CreateHSRootDevice(UPnPDeviceInfo.UPnPDeviceDeviceType, UPnPDeviceInfo.UPnPDeviceGivenName, UPnPDeviceInfo.UPnPDeviceUDN)

            UPnPDeviceInfo.UPnPDeviceAdminStateActive = GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diAdminState.ToString, False)
            UPnPDeviceInfo.UPnPDeviceIsAddedToHS = GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)

            CreateUPnPControllers(True) ' this will instanciate a new instance of HSPI and pass all parameters into the new instance and call DirectConnect

        Catch ex As Exception
            Log("Error in ActivateDevice with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub RemoveDevicefromHS(DeviceUDN As String)
        Log("RemoveDevicefromHS called with DeviceUDN = " & DeviceUDN.ToString, LogType.LOG_TYPE_INFO)
        If GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False) Then
            DMRemove(DeviceUDN)
        End If
        ' [UPnP Devices UDN to Info] Key is UDN
        'objIniFile.DeleteEntry("UPnP Devices UDN to Info", DeviceUDN)
        Dim KeyValues As New Dictionary(Of String, String)()
        ' [UPnP Devices DC to UDN] scan through the section and remove the entries with value = UDN and call HS to delete the device!
        KeyValues = GetIniSection("UPnP HSRef to UDN")
        For Each Entry In KeyValues
            If Entry.Value = DeviceUDN Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RemoveDevicefromHS found HSRef = " & Entry.Key.ToString, LogType.LOG_TYPE_INFO)
                If Entry.Key <> -1 Then
                    Log("RemoveDevicefromHS deleted HSRef = " & Entry.Key.ToString, LogType.LOG_TYPE_INFO)
                    hs.DeleteDevice(Entry.Key)
                End If
                DeleteEntryIniFile("UPnP HSRef to UDN", Entry.Key)
            End If
        Next

        ' [DLNADevicePage] should delete the entries for the musicAPI <> 0
        If GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "DMR" Or GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "HST" Then
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_ServerUDN")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_loaded")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_root")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level1")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level2")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level3")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level4")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level5")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level6")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level7")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level8")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level9")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level10")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_rootValue")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level1Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level2Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level3Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level4Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level5Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level6Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level7Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level8Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level9Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level10Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_playlistSelector")
            ' remove the weblink in case of DMR or HST!
            DeleteWebLink(DeviceUDN, GetStringIniFile(DeviceUDN, DeviceInfoIndex.diGivenName.ToString, ""))
        End If

        ' if there is a diRemoteType=ROKU in the [UDN] section, then we need to delete from the remote file
        Dim RemoteType As String = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        If RemoteType <> "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RemoveDevicefromHS called with DeviceUDN = " & DeviceUDN.ToString & " is removing the Remote Control information", LogType.LOG_TYPE_INFO)
            Dim objRemoteFile As String = gRemoteControlPath
            '   whole section with [UDN]
            DeleteIniSection(DeviceUDN, gRemoteControlPath)
            'Try
            'KeyValues = GetIniSection(DeviceUDN, objRemoteFile)
            'For Each KeyValue In KeyValues
            'DeleteEntryIniFile(DeviceUDN, KeyValue.Key, objRemoteFile)
            'Next
            'Catch ex As Exception
            'Log("Error in RemoveDevicefromHS for remote deleting [" & DeviceUDN & "] section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            'End Try
            '   whole section with [UDN  - Default Codes]
            DeleteIniSection(DeviceUDN & " - Default Codes", gRemoteControlPath)
            'Try
            'KeyValues = GetIniSection(DeviceUDN & " - Default Codes", objRemoteFile)
            'For Each KeyValue In KeyValues
            'DeleteEntryIniFile(DeviceUDN & " - Default Codes", KeyValue.Key, objRemoteFile)
            'Next
            'Catch ex As Exception
            'Log("Error in RemoveDevicefromHS for remote deleting [" & DeviceUDN & " - Default Codes" & "] section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            'End Try
        End If
        ' Delete parts of the section [UDN]
        WriteBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
        WriteBooleanIniFile(DeviceUDN, DeviceInfoIndex.diAdminState.ToString, False)
        WriteIntegerIniFile(DeviceUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, 0)
        WriteIntegerIniFile(DeviceUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, 0)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diHSDeviceRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diRegistered.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diAdminStateRemote.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diAdminStateMessageService.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diUseNextAV.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diNextAV.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diTrackHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSpeedHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diNextTrackHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diArtistHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diNextArtistHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diAlbumHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diNextAlbumHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diArtHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diNextArtHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diPlayStateHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diVolumeHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diMuteHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diBalanceHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diTrackLengthHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diTrackPosHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diTrackDescrHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diQueueRepeatHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diQueueShuffleHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diQueueRepeatHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diQueueShuffleHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diGenreHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diPlayerHSRef.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diLGClientKey.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSecWebSocketKey.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSamsungToken.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diWifiMacAddress.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSamsungTokenAuthSupport.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSamsungisSupportInfo.ToString)
        DeleteEntryIniFile(DeviceUDN, DeviceInfoIndex.diSamsungClientID.ToString)

        DeleteEntryIniFile(DeviceUDN, "diRemoteHSCode")
        DeleteEntryIniFile(DeviceUDN, "diDeviceControlHSCode")
        DeleteEntryIniFile(DeviceUDN, "diServerHSCode")

        Dim PartyServiceHSCode As String = ""
        PartyServiceHSCode = GetStringIniFile(DeviceUDN, "diPartyHSCode", "")
        If PartyServiceHSCode <> "" Then
            RemovePartyDevice(DeviceUDN)
        End If
        DeleteEntryIniFile(DeviceUDN, "diPartyHSCode")
        ' 
        ' 
        ' delete the playlist with UDN.txt
        Dim PlaylistName As String = CurrentAppPath & gPlaylistPath & DeviceUDN & ".txt"
        Try
            If System.IO.File.Exists(PlaylistName) = True Then System.IO.File.Delete(PlaylistName)
        Catch ex As Exception
            Log("Error in RemoveDevicefromHS for UDN = " & DeviceUDN & " and Playlist = " & PlaylistName & " deleting the Playlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            DeleteEntryIniFile("Speaker Devices", DeviceUDN)
        Catch ex As Exception
        End Try
        Try
            DeleteEntryIniFile("Message Devices", DeviceUDN)
        Catch ex As Exception
        End Try
        Try
            DeleteEntryIniFile("Message Service by UDN", DeviceUDN)
        Catch ex As Exception
        End Try
        Try
            DeleteEntryIniFile("Remote Service by UDN", DeviceUDN)
        Catch ex As Exception
        End Try
        If MainInstance <> "" Then
            RemoveInstance(MainInstance & "-" & DeviceUDN)
        Else
            RemoveInstance(DeviceUDN)
        End If
    End Sub

    Public Sub DeleteDevice(DeviceUDN As String)
        Log("DeleteDevice called with DeviceUDN = " & DeviceUDN.ToString, LogType.LOG_TYPE_INFO)

        If GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False) Then
            DMRemove(DeviceUDN)
        End If
        ' [UPnP Devices UDN to Info] Key is UDN
        DeleteEntryIniFile("UPnP Devices UDN to Info", DeviceUDN)

        Dim KeyValues As New Dictionary(Of String, String)()

        ' [UPnP Devices DC to UDN] scan through the section and remove the entries with value = UDN and call HS to delete the device!
        KeyValues = GetIniSection("UPnP HSRef to UDN")
        For Each Entry In KeyValues
            If Entry.Value = DeviceUDN Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteDevice found HSRef = " & Entry.Key.ToString, LogType.LOG_TYPE_INFO)
                If Entry.Key <> -1 Then
                    Log("DeleteDevice deleted HSRef = " & Entry.Key.ToString, LogType.LOG_TYPE_INFO)
                    hs.DeleteDevice(Entry.Key)
                End If
                DeleteEntryIniFile("UPnP HSRef to UDN", Entry.Key)
            End If
        Next

        ' [DLNADevicePage] should delete the entries for the musicAPI <> 0
        If GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "DMR" Or GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "HST" Then
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_ServerUDN")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_loaded")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_root")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level1")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level2")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level3")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level4")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level5")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level6")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level7")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level8")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level9")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level10")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_rootValue")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level1Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level2Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level3Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level4Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level5Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level6Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level7Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level8Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level9Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_level10Value")
            DeleteEntryIniFile("DevicePage", DeviceUDN & "_playlistSelector")
            ' remove the weblink in case of DMR or HST!
            DeleteWebLink(DeviceUDN, GetStringIniFile(DeviceUDN, DeviceInfoIndex.diGivenName.ToString, ""))
        End If

        ' if there is a diRemoteType=ROKU in the [UDN] section, then we need to delete from the remote file
        Dim RemoteType As String = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        If RemoteType <> "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeactivateDevice called with DeviceUDN = " & DeviceUDN.ToString & " is removing the Remote Control information", LogType.LOG_TYPE_INFO)
            Dim objRemoteFile As String = gRemoteControlPath
            '   whole section with [UDN]
            Try
                KeyValues = GetIniSection(DeviceUDN, objRemoteFile)
                For Each KeyValue In KeyValues
                    DeleteEntryIniFile(DeviceUDN, KeyValue.Key, objRemoteFile)
                Next
            Catch ex As Exception
                Log("Error in DeleteDevice for remote deleting [" & DeviceUDN & "] section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            '   whole section with [UDN  - Default Codes]
            Try
                KeyValues = GetIniSection(DeviceUDN & " - Default Codes", objRemoteFile)
                For Each KeyValue In KeyValues
                    DeleteEntryIniFile(DeviceUDN & " - Default Codes", KeyValue.Key, objRemoteFile)
                Next
            Catch ex As Exception
                Log("Error in DeleteDevice for remote deleting [" & DeviceUDN & " - Default Codes" & "] section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        ' Whole section [UDN]
        DeleteIniSection(DeviceUDN)
        ' delete the playlist with UDN.txt
        Dim PlaylistName As String = CurrentAppPath & gPlaylistPath & DeviceUDN & ".txt"
        Try
            If System.IO.File.Exists(PlaylistName) = True Then System.IO.File.Delete(PlaylistName)
        Catch ex As Exception
            Log("Error in DeleteDevice for UDN = " & DeviceUDN & " and Playlist = " & PlaylistName & " deleting the Playlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If MainInstance <> "" Then
            RemoveInstance(MainInstance & "-" & DeviceUDN)
        Else
            RemoveInstance(DeviceUDN)
        End If
    End Sub

    Public Sub DoRediscover()

        If PIDebuglevel > DebugLevel.dlEvents Then Log("DoRediscover called", LogType.LOG_TYPE_INFO)
        Dim UPnPDeviceInfo As MyUPnpDeviceInfo = Nothing
        Dim UPnPDevicesToGoDiscover As New System.Collections.Generic.Dictionary(Of String, String)()
        UPnPDevicesToGoDiscover = GetIniSection("UPnP Devices to discover") '  As Dictionary(Of String, String)

        If UPnPDevicesToGoDiscover Is Nothing Then
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoRediscover. No Devices specified in the .ini file under ""UPnP Devices to discover""", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Try
            Dim AllDevices As MyUPnPDevices = MySSDPDevice.GetAllDevices()
            If Not AllDevices Is Nothing And AllDevices.Count > 0 Then
                For Each DLNADevice As MyUPnPDevice In AllDevices
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("DoRediscover found UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_INFO) ' moved this here on 11/16/2019
                    If DLNADevice IsNot Nothing Then    ' added on 3/3/2019 to prevent errors happening here due to unknown causes
                        If (DLNADevice.UniqueDeviceName <> "") And (DLNADevice.Location <> "") And DLNADevice.Alive Then
                            ' check whether this devices was known to us and on-line
                            ' go find it in the array
                            Dim NeedsToBeAdded As Boolean = False
                            If Not (Mid(DLNADevice.UniqueDeviceName, 1, 12) = "uuid:RINCON_" Or Mid(DLNADevice.UniqueDeviceName, 1, 16) = "uuid:DOCKRINCON_") Or SonosDeviceIn Then
                                For Each UPnPDeviceToDiscover In UPnPDevicesToGoDiscover
                                    If DLNADevice.Type = UPnPDeviceToDiscover.Key Then
                                        NeedsToBeAdded = True
                                        Exit For
                                    End If
                                Next
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("DoRediscover found UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_INFO)
                            If Not NeedsToBeAdded Then GoTo NextElement
                            If Not GetBooleanIniFile(DLNADevice.UniqueDeviceName, DeviceInfoIndex.diDeviceIsAdded.ToString, False) Then GoTo NextElement ' is not added
                            Try
                                UPnPDeviceInfo = FindUPnPDeviceInfo(DLNADevice.UniqueDeviceName)
                                If UPnPDeviceInfo Is Nothing Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover found New UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                    NewDeviceFound(DLNADevice.UniqueDeviceName)
                                Else
                                    Dim Controller As HSPI = GetAPIByUDN(DLNADevice.UniqueDeviceName)
                                    If Controller Is Nothing Then
                                        ' this should really not be
                                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover shouldn''t have found UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                        'NewDeviceFound(DLNADevice.UniqueDeviceName)
                                    Else
                                        If Controller.DeviceStatus.ToUpper <> "ONLINE" Then
                                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover found Known UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                            NewDeviceFound(DLNADevice.UniqueDeviceName)
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                Log("Error in DoRediscover while finding the DeviceInfo NewUDN = " & DLNADevice.UniqueDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                GoTo NextElement
                            End Try
                        End If
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DoRediscover found missing DLNA Object", LogType.LOG_TYPE_WARNING)

                    End If
NextElement:
                Next
            End If
        Catch ex As Exception
            Log("Error in DoRediscover with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub


    Private Function CreateHSRootDevice(DevType As String, DeviceName As String, DeviceUDN As String) As Integer
        CreateHSRootDevice = -1
        Log("CreateHSRootDevice called with DevType = " & DevType & " and DeviceName = " & DeviceName & " and DeviceUDN = " & DeviceUDN, LogType.LOG_TYPE_INFO)
        Dim dv As Scheduler.Classes.DeviceClass
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        Dim HSRef As Integer = GetIntegerIniFile(DeviceUDN, DeviceInfoIndex.diHSDeviceRef.ToString, -1) ' fixed in v24
        Try
            If HSRef = -1 Then
                HSRef = hs.NewDeviceRef("Device") ' (DeviceName)
                Log("CreateHSRootDevice: Created device " & DeviceName & " with reference " & HSRef.ToString, LogType.LOG_TYPE_INFO)
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
                If HSRef <> -1 Then WriteStringIniFile("UPnP HSRef to UDN", HSRef, DeviceUDN)
            Else
                Return HSRef
            End If
            dv = hs.GetDeviceByRef(HSRef)
            dv.Interface(hs) = sIFACE_NAME
            dv.Location2(hs) = tIFACE_NAME
            dv.InterfaceInstance(hs) = MainInstance
            dv.Location(hs) = DeviceName
            dv.Device_Type_String(hs) = HSDevices.RootDevice.ToString
            dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
            dv.Address(hs) = "Root"
            Dim DT As New DeviceTypeInfo
            DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
            DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Root
            DT.Device_SubType_Description = RootHSDescription
            dv.DeviceType_Set(hs) = DT
            dv.Status_Support(hs) = True
            'hs.SetDeviceString(HSRef, "_", False)
            ' dv.Image(hs) is set in myDeviceFinderCallback_DeviceFound after the image was downloaded
            ' This device is a child device, the parent being the root device for the entire security system. 
            ' As such, this device needs to be associated with the root (Parent) device.
            If DevType = "HST" Then
                dv.Image(hs) = ImagesPath & "Slideshow.png"
                dv.ImageLarge(hs) = ImagesPath & "Slideshow.png"
            End If
            dvParent = hs.GetDeviceByRef(MasterHSDeviceRef)
            If dvParent.AssociatedDevices_Count(Nothing) < 1 Then
                ' There are none added, so it is OK to add this one.
                'dvParent.AssociatedDevice_Add(hs, HSRef)
            Else
                Dim Found As Boolean = False
                For Each ref As Integer In dvParent.AssociatedDevices(Nothing)
                    If ref = HSRef Then
                        Found = True
                        Exit For
                    End If
                Next
                If Not Found Then
                    'dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    ' This is an error condition likely as this device's reference ID should not already be associated.
                End If
            End If

            ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
            '   the child's associations.
            dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
            'dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
            dv.Relationship(hs) = Enums.eRelationship.Parent_Root
            hs.DeviceVSP_ClearAll(HSRef, True)
            hs.DeviceVGP_ClearAll(HSRef, True)

            CreateActivateDeactivateButtons(HSRef, True) ' dcor needs fixing!
            WriteIntegerIniFile(DeviceUDN, DeviceInfoIndex.diHSDeviceRef.ToString, HSRef) ' store it

            If DevType = "DIAL" Then
                WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "DIAL")
            ElseIf DevType = "PMR" Then
                WriteBooleanIniFile("Message Service by UDN", DeviceUDN, True)
            End If
            hs.SaveEventsDevices()
            Return HSRef
        Catch ex As Exception
            Log("Error in CreateHSRootDevice with DevType = " & DevType & " and DeviceName = " & DeviceName & " and DeviceUDN = " & DeviceUDN & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub CreateSlideshowDevice()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateSlideshowDevice called", LogType.LOG_TYPE_INFO)
        Dim Found As Boolean = False
        Dim LoopCount As Integer = 0
        Dim SearchNewUDN As String = "PictureshowUDN-"
        Dim SlideShowDeviceUDN As String = ""
        While LoopCount < 50
            SlideShowDeviceUDN = GetStringIniFile("UPnP Devices UDN to Info", SearchNewUDN & LoopCount.ToString, "")
            If SlideShowDeviceUDN = "" Then
                Found = True
                Exit While
            End If
            LoopCount += 1
        End While
        If Not Found Then
            Log("CreateSlideshowDevice could not find a free UDN. Max limit of 50 exceeded.", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        SearchNewUDN = SearchNewUDN & LoopCount.ToString
        Log("CreateSlideshowDevice created a Pictureshow device with UDN = " & SearchNewUDN, LogType.LOG_TYPE_INFO)
        WriteStringIniFile("UPnP Devices UDN to Info", SearchNewUDN, "Pictureshow Device")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diGivenName.ToString, "Pictureshow Device " & LoopCount.ToString)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceModelName.ToString, "")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceType.ToString, "HST")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diFriendlyName.ToString, "Pictureshow Device " & LoopCount.ToString)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diIPAddress.ToString, hs.GetIPAddress)
        Dim HTTPPort As String = "" ' found in settings.ini at gWebSvrPort=80
        HTTPPort = hs.GetINISetting("Settings", "gWebSvrPort", "")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diIPPort.ToString, HTTPPort)
        WriteIntegerIniFile(SearchNewUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, "0")
        WriteBooleanIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
        WriteIntegerIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, "0")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceIConURL.ToString, ImagesPath & "Slideshow.png")
    End Sub

    Public Sub CreateMediaDevice(DeviceType As String, IPAddress As String, IPPort As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateMediaDevice called", LogType.LOG_TYPE_INFO)
        Dim Found As Boolean = False
        Dim LoopCount As Integer = 0
        Dim SearchNewUDN As String = "MediaDeviceUDN-"
        Dim MediaDeviceUDN As String = ""
        While LoopCount < 50
            MediaDeviceUDN = GetStringIniFile("UPnP Devices UDN to Info", SearchNewUDN & LoopCount.ToString, "")
            If MediaDeviceUDN = "" Then
                Found = True
                Exit While
            End If
            LoopCount += 1
        End While
        If Not Found Then
            Log("CreateMediaDevice could not find a free UDN. Max limit of 50 exceeded.", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        SearchNewUDN = SearchNewUDN & LoopCount.ToString
        Log("CreateMediaDevice created a Media device with UDN = " & SearchNewUDN, LogType.LOG_TYPE_INFO)
        WriteStringIniFile("UPnP Devices UDN to Info", SearchNewUDN, "Media Device")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diGivenName.ToString, "Media Device " & LoopCount.ToString)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceModelName.ToString, "")
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceType.ToString, DeviceType)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diFriendlyName.ToString, "Media Device " & LoopCount.ToString)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diIPAddress.ToString, IPAddress)
        WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diIPPort.ToString, IPPort)
        If DeviceType.ToUpper = "DIAL" Then
            WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceIConURL.ToString, ImagesPath & "Roku.png")
        Else
            WriteStringIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceIConURL.ToString, ImagesPath & "Mediadevice.png")
        End If
        WriteIntegerIniFile(SearchNewUDN, DeviceInfoIndex.diMusicAPIIndex.ToString, "0")
        WriteBooleanIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
        WriteIntegerIniFile(SearchNewUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, "0")

    End Sub


    Public Sub UpdatePartyDeviceButtons()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdatePartyDeviceButtons called", LogType.LOG_TYPE_INFO)
        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Try
            ' go through each Party device. Pick up the UDN, then with UDN , pick up Device Code, then update buttonstring
            For Each PartyDevice In PartyDevices
                If PartyDevice.Key <> "" Then
                    Dim DLNADevice As HSPI = Nothing
                    DLNADevice = GetAPIByUDN(PartyDevice.Key)
                    If Not DLNADevice Is Nothing Then
                        DLNADevice.SonyUpdatePartyButtons()
                    End If
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub AddPartyDevice(PartyUDN As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddPartyDevice called with PartyUDN = " & PartyUDN.ToString, LogType.LOG_TYPE_INFO)
        If GetStringIniFile("Party Devices", PartyUDN, "") <> "" Then
            ' already exist do nothing
            Exit Sub
        End If
        ' Find next free index
        Dim FreeIndex As Integer = GetNextFreePartyIndex()
        If FreeIndex <> 0 Then
            WriteStringIniFile("Party Devices", PartyUDN, FreeIndex.ToString)
            UpdatePartyDeviceButtons()
        Else
            Log("Error in AddPartyDevice, couldn't find free PartyIndex", LogType.LOG_TYPE_ERROR)
        End If
    End Sub

    Private Function GetNextFreePartyIndex() As Integer
        GetNextFreePartyIndex = 0
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreePartyIndex called", LogType.LOG_TYPE_INFO)
        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Dim LowestFreeIndex As Integer = 0
        Dim IndexFound As Boolean = False
        While Not IndexFound
            LowestFreeIndex = LowestFreeIndex + 1
            IndexFound = True
            For Each PartyDevice In PartyDevices
                If PartyDevice.Key <> "" Then
                    Try
                        If Val(PartyDevice.Value) = LowestFreeIndex Then
                            IndexFound = False
                            Exit For
                        End If
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetNextFreePartyIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Next
            If LowestFreeIndex > 100 Then
                IndexFound = False  ' force an exit
                LowestFreeIndex = 0
            End If
        End While
        GetNextFreePartyIndex = LowestFreeIndex
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreePartyIndex found Index = " & LowestFreeIndex.ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Sub RemovePartyDevice(PartyUDN As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RemovePartyDevice called with PartyUDN = " & PartyUDN.ToString, LogType.LOG_TYPE_INFO)
        DeleteEntryIniFile("Party Devices", PartyUDN)
        UpdatePartyDeviceButtons()
    End Sub

    Public Function AuthenticateSony(DeviceUDN As String, SonyPIN As String) As Boolean
        AuthenticateSony = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AuthenticateSony called with DeviceUDN = " & DeviceUDN & " and SonyPIN = " & SonyPIN, LogType.LOG_TYPE_INFO)
        Try
            Dim UPnPDevice As HSPI = GetAPIByUDN(DeviceUDN)
            If UPnPDevice IsNot Nothing Then
                AuthenticateSony = UPnPDevice.SendJSONAuthentication(SonyPIN)
            End If
        Catch ex As Exception
            Log("Error in AuthenticateSony for DeviceUDN = " & DeviceUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Public Function AuthenticateSamsung(DeviceUDN As String, SamsungPIN As String) As Boolean
        AuthenticateSamsung = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AuthenticateSamsung called with DeviceUDN = " & DeviceUDN & " and SamsungPIN = " & SamsungPIN, LogType.LOG_TYPE_INFO)
        Try
            Dim UPnPDevice As HSPI = GetAPIByUDN(DeviceUDN)
            If UPnPDevice IsNot Nothing Then
                If UPnPDevice.SamsungAuthenticateUsePIN(SamsungPIN, False) <> "" Then
                    Return True
                End If
            End If
        Catch ex As Exception
            Log("Error in AuthenticateSamsung for DeviceUDN = " & DeviceUDN & " and Error = " & ex.ToString, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SendOpenPINtoCorrectSamsungDevice(DeviceUDN As String) As Boolean
        SendOpenPINtoCorrectSamsungDevice = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendOpenPINtoCorrectSamsungDevice called with DeviceUDN = " & DeviceUDN, LogType.LOG_TYPE_INFO)
        Try
            Dim UPnPDevice As HSPI = GetAPIByUDN(DeviceUDN)
            If UPnPDevice IsNot Nothing Then
                SendOpenPINtoCorrectSamsungDevice = UPnPDevice.SamsungOpenPinPage()
            End If
        Catch ex As Exception
            Log("Error in SendOpenPINtoCorrectSamsungDevice for DeviceUDN = " & DeviceUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub UpdateDeviceName(DeviceUDN As String, NewGivenName As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateDeviceName called for UDN = " & DeviceUDN & " and NewGivenName = " & NewGivenName, LogType.LOG_TYPE_INFO)

        Dim OldGivenName As String = GetStringIniFile(DeviceUDN, "diGivenName", "")
        Try
            WriteStringIniFile(DeviceUDN, "diGivenName", NewGivenName)
        Catch ex As Exception
            Log("Error writing UpdateDeviceName to ini file with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        'WriteStringIniFile("UPnP Devices UDN to Info", DeviceUDN, NewGivenName)

        Dim UPnPDevice As HSPI = GetAPIByUDN(DeviceUDN)
        If UPnPDevice IsNot Nothing Then
            UPnPDevice.DeviceName = NewGivenName ' update the Device Name in the instance
            If GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "DMR" Or GetStringIniFile(DeviceUDN, DeviceInfoIndex.diDeviceType.ToString, "") = "HST" Then
                UPnPDevice.ChangeWebLink(OldGivenName, NewGivenName, DeviceUDN)
            End If
        End If
        Try
            Dim SameDLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
            SameDLNADevices = GetIniSection("UPnP HSRef to UDN") '  As Dictionary(Of String, String)
            For Each SameDLNADevice In SameDLNADevices
                If SameDLNADevice.Value = DeviceUDN Then
                    Try
                        If SameDLNADevice.Key = -1 Then
                            ' should not be
                            Log("Error in UpdateDeviceName, couldn't retrieve Device Ref from HS with Ref = " & SameDLNADevice.Key.ToString, LogType.LOG_TYPE_ERROR)
                            Exit Try
                        End If
                        Dim dv As Scheduler.Classes.DeviceClass
                        dv = hs.GetDeviceByRef(SameDLNADevice.Key)
                        If dv IsNot Nothing Then
                            If dv.Location(hs) = OldGivenName Then
                                dv.Location(hs) = NewGivenName
                            End If
                            If dv.Name(hs) = OldGivenName Then
                                dv.Name(hs) = NewGivenName
                            End If
                        End If
                    Catch ex As Exception
                        Log("Error in UpdateDeviceName updating HS with HSRef = " & SameDLNADevice.Key.ToString & " and error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Next
        Catch ex As Exception
            Log("Error in UpdateDeviceName1 updating HS with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MyHSDeviceLinkedList.Count <> 0 Then
                Try
                    For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                        If HSDevice.UPnPDeviceUDN = DeviceUDN Then
                            HSDevice.UPnPDeviceGivenName = NewGivenName
                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    Log("Error in UpdateDeviceName2 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Catch ex As Exception
            Log("Error in UpdateDeviceName3 updating HS with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        hs.SaveEventsDevices()
    End Sub



    Private Sub CreateWebLink(ByVal inZoneName As String, ByVal inZoneUDN As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateWebLink called with DeviceUDN = " & inZoneUDN & " and PageName = " & PlayerControlPage & inZoneUDN.Substring(7, inZoneUDN.Length - 7), LogType.LOG_TYPE_INFO)

        Try

            If MainInstance <> "" Then
                MyPlayerControlWebPage = New PlayerControl(PlayerControlPage & ":" & MainInstance & "-" & inZoneUDN)
            Else
                MyPlayerControlWebPage = New PlayerControl(PlayerControlPage & ":" & inZoneUDN)
            End If

            MyPlayerControlWebPage.RefToPlugIn = MyReferenceToMyController 'Me
            MyPlayerControlWebPage.ZoneUDN = inZoneUDN

            ' register the page with the HS web server, HS will post back to the WebPage class
            ' "pluginpage" is the URL to access this page
            ' comment this out if you are going to use the GenPage/PutPage API istead

            If MainInstance <> "" Then
                hs.RegisterPage(PlayerControlPage, sIFACE_NAME, MainInstance & "-" & inZoneUDN)
            Else
                hs.RegisterPage(PlayerControlPage, sIFACE_NAME, inZoneUDN)
            End If

            ' register a configuration link that will appear on the interfaces page
            Dim wpd As New WebPageDesc
            ' register a normal page to appear in the HomeSeer menu
            wpd = New WebPageDesc
            wpd.link = PlayerControlPage
            wpd.page_title = "Media Controller " & inZoneName & " Config"
            wpd.plugInName = sIFACE_NAME

            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & inZoneName
                wpd.plugInInstance = MainInstance & "-" & inZoneUDN
            Else
                wpd.linktext = inZoneName
                wpd.plugInInstance = inZoneUDN
            End If

            hs.RegisterLinkEx(wpd)

        Catch ex As Exception
            Log("Error in CreateWebLink , unable to register the PlayerControl link with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        CreateConfigLink(inZoneName, inZoneUDN)

    End Sub

    Private Sub CreateConfigLink(ByVal inZoneName As String, ByVal inZoneUDN As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateConfigLink called with DeviceUDN = " & inZoneUDN & " and PageName = " & PlayerControlPage & inZoneUDN.Substring(7, inZoneUDN.Length - 7), LogType.LOG_TYPE_INFO)

        Try
            If MainInstance <> "" Then
                MyPlayerConfigWebPage = New DLNADeviceConfig(PlayerConfig & ":" & MainInstance & "-" & inZoneUDN)
            Else
                MyPlayerConfigWebPage = New DLNADeviceConfig(PlayerConfig & ":" & inZoneUDN)
            End If

            MyPlayerConfigWebPage.RefToPlugIn = MyReferenceToMyController 'Me
            MyPlayerConfigWebPage.ZoneUDN = inZoneUDN

            ' register the page with the HS web server, HS will post back to the WebPage class
            ' "pluginpage" is the URL to access this page
            ' comment this out if you are going to use the GenPage/PutPage API istead

            If MainInstance <> "" Then
                hs.RegisterPage(PlayerConfig, sIFACE_NAME, MainInstance & "-" & inZoneUDN)
            Else
                hs.RegisterPage(PlayerConfig, sIFACE_NAME, inZoneUDN)
            End If

        Catch ex As Exception
            Log("Error in CreateConfigLink , unable to register the PlayerConfig link with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Public Sub DeleteWebLink(ZoneUDN As String, inZoneName As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteWebLink called with DeviceUDN = " & ZoneUDN, LogType.LOG_TYPE_INFO)
        Exit Sub
        Try
            If MyPlayerControlWebPage IsNot Nothing Then
                ' register a configuration link that will appear on the interfaces page
                Dim wpd As New WebPageDesc
                ' register a normal page to appear in the HomeSeer menu
                wpd = New WebPageDesc
                wpd.link = PlayerControlPage
                wpd.page_title = "Media Controller " & inZoneName & " Config"
                wpd.plugInName = sIFACE_NAME

                If MainInstance <> "" Then
                    wpd.linktext = "Instance " & MainInstance & " " & inZoneName
                    wpd.plugInInstance = MainInstance & "-" & ZoneUDN
                Else
                    wpd.linktext = inZoneName
                    wpd.plugInInstance = ZoneUDN
                End If
                hs.UnRegisterLinkEx(wpd)
                MyPlayerControlWebPage.Dispose()
                MyPlayerControlWebPage = Nothing
            End If
        Catch ex As Exception
            Log("Error in DeleteWebLink for DeviceUDN = " & ZoneUDN & ". Unable to UnRegister the PlayerControl link with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MyPlayerConfigWebPage IsNot Nothing Then
                MyPlayerConfigWebPage.Dispose()
                MyPlayerConfigWebPage = Nothing
            End If
        Catch ex As Exception
            Log("Error in DeleteWebLink for DeviceUDN = " & ZoneUDN & ". Unable to UnRegister the PlayerConfig link with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ChangeWebLink(OldZoneName As String, ByVal NewZoneName As String, ByVal ZoneUDN As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeWebLink called with ZoneUDN = " & ZoneUDN & " and NewZoneName = " & NewZoneName, LogType.LOG_TYPE_INFO)
        Try

            ' register a configuration link that will appear on the interfaces page
            Dim wpd As New WebPageDesc
            ' register a normal page to appear in the HomeSeer menu
            wpd = New WebPageDesc
            wpd.link = PlayerControlPage

            wpd.page_title = "Media Controller " & NewZoneName & " Config"
            wpd.plugInName = sIFACE_NAME
            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & OldZoneName
                wpd.plugInInstance = MainInstance & "-" & ZoneUDN 'instance
            Else
                wpd.linktext = OldZoneName
                wpd.plugInInstance = ZoneUDN 'instance
            End If
            hs.UnRegisterLinkEx(wpd)
            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & NewZoneName
                wpd.plugInInstance = MainInstance & "-" & ZoneUDN 'instance
            Else
                wpd.linktext = NewZoneName
                wpd.plugInInstance = ZoneUDN 'instance
            End If
            hs.RegisterLinkEx(wpd)
        Catch ex As Exception
            Log("Error in CreateWebLink , unable to register the link(ex) with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub DestroyUPnPControllers()
        Dim UPnPController As HSPI
        If MyHSDeviceLinkedList Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroyUPnPControllers called for instance " & instance & " but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If MyHSDeviceLinkedList.Count = 0 Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroyUPnPControllers called for instance " & instance & " but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroyUPnPControllers: found " & MyHSDeviceLinkedList.Count & " Device Codes", LogType.LOG_TYPE_INFO)
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If Not HSDevice.UPnPDeviceControllerRef Is Nothing Then
                Try
                    UPnPController = HSDevice.UPnPDeviceControllerRef
                    UPnPController.Disconnect(True)
                    UPnPController.DeleteWebLink(UPnPController.DeviceUDN, UPnPController.DeviceName)

                    Dim OldUDN = UPnPController.MyUDN ' changed in .45
                    UPnPController.DestroyPlayer(True)
                    HSDevice.UPnPDeviceControllerRef = Nothing
                    RemoveInstance(OldUDN) ' changed in.45

                    'UPnPController.DestroyPlayer(True)
                    'HSDevice.UPnPDeviceControllerRef = Nothing
                Catch ex As Exception
                    Log("DestroyUPnPControllers: Could not disconnect with Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
                    'Exit For
                End Try
            End If
        Next
    End Sub



    Public Function GetAPIByUDN(ByVal inUDN As String) As HSPI
        ' Returns the name of this instance as set in the plug-in configuration.
        GetAPIByUDN = Nothing
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetInstanceByName called with value : " & Instance)
        If MyHSDeviceLinkedList.Count = 0 Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in GetAPIByUDN. There are no devices. UDN : " & inUDN, LogType.LOG_TYPE_WARNING)
            Exit Function
        End If
        inUDN = Trim(inUDN)
        If inUDN = "" Then
            Exit Function
        End If
        If Mid(inUDN, 1, 5) = "uuid:" Then
            Mid(inUDN, 1, 5) = "     "
            inUDN = Trim(inUDN)
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.UPnPDeviceUDN = inUDN Then
                GetAPIByUDN = HSDevice.UPnPDeviceControllerRef
                Exit Function
            End If
        Next
    End Function

    Public Function GetDeviceInstanceByName(ByVal Instance As String) As Integer
        ' Returns the name of this instance as set in the plug-in configuration.
        ' this is either a device name or a device UDN in starting with uuid:
        GetDeviceInstanceByName = 0
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetDeviceInstanceByName called with value : " & Instance, LogType.LOG_TYPE_INFO)
        Instance = Trim(Instance)
        If Instance = "" Then
            Exit Function
        End If
        If Mid(Instance, 1, 5) = "uuid:" Then
            Mid(Instance, 1, 5) = "     "
            Instance = Trim(Instance)
        End If
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        Dim UPnPPlayer As HSPI
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            UPnPPlayer = HSDevice.UPnPDeviceControllerRef
            If Not UPnPPlayer Is Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetDeviceInstanceByName called with value : """ & Instance & """ and found """ & UPnPPlayer.DeviceName & """ and UDN """ & UPnPPlayer.DeviceUDN & """", LogType.LOG_TYPE_INFO)
                If (UPnPPlayer.DeviceName = Instance) Or (UPnPPlayer.DeviceUDN = Instance) Then
                    GetDeviceInstanceByName = UPnPPlayer.DeviceAPIIndex
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetDeviceInstanceByName found value : " & Instance & " at Index = " & GetDeviceInstanceByName, LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetDeviceInstanceByName. Did not find value : " & Instance.ToString, LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetMusicInstanceByName(ByVal Instance As String) As Integer
        ' Returns the name of this instance as set in the plug-in configuration.
        ' this is either a device name or a device UDN in starting with uuid:
        GetMusicInstanceByName = 0
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetMusicInstanceByName called with value : " & Instance, LogType.LOG_TYPE_INFO)
        Instance = Trim(Instance)
        If Instance = "" Then
            Exit Function
        End If
        If Mid(Instance, 1, 5) = "uuid:" Then
            Mid(Instance, 1, 5) = "     "
            Instance = Trim(Instance)
        End If
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        Dim UPnPPlayer As HSPI
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            UPnPPlayer = HSDevice.UPnPDeviceControllerRef
            If Not UPnPPlayer Is Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetMusicInstanceByName called with value : """ & Instance & """ and found """ & UPnPPlayer.DeviceName & """ and UDN """ & UPnPPlayer.DeviceUDN & """", LogType.LOG_TYPE_INFO)
                If (UPnPPlayer.DeviceName = Instance) Or (UPnPPlayer.DeviceUDN = Instance) Then
                    GetMusicInstanceByName = UPnPPlayer.APIInstance
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetMusicInstanceByName found value : " & Instance & " at Index = " & GetMusicInstanceByName, LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetMusicInstanceByName. Did not find value : " & Instance.ToString, LogType.LOG_TYPE_ERROR)
    End Function


    Public Function GetDeviceGivenNameByUDN(UDN As String) As String
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetDeviceGivenNameByUDN called with UDN = " & UDN, LogType.LOG_TYPE_INFO)
        If Mid(UDN, 1, 5) = "uuid:" Then
            Mid(UDN, 1, 5) = "     "
            UDN = Trim(UDN)
        End If
        GetDeviceGivenNameByUDN = ""
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
        For Each DLNADevice In DLNADevices
            If DLNADevice.Key <> "" Then
                Try
                    If DLNADevice.Key = UDN Then
                        GetDeviceGivenNameByUDN = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diGivenName.ToString, "")
                        Exit Function
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetDeviceGivenNameByUDN with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetDeviceGivenNameByUDN. Did not find UDN = " & UDN, LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetDeviceRefByUDN(UDN As String) As Integer
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetDeviceRefByUDN called with UDN = " & UDN, LogType.LOG_TYPE_INFO)
        GetDeviceRefByUDN = -1
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
        For Each DLNADevice In DLNADevices
            If DLNADevice.Key <> "" Then
                Try
                    If DLNADevice.Key = UDN Then
                        GetDeviceRefByUDN = GetIntegerIniFile(DLNADevice.Key, DeviceInfoIndex.diHSDeviceRef.ToString, -1)
                        Exit Function
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetDeviceRefByUDN with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetDeviceRefByUDN. Did not find UDN = " & UDN, LogType.LOG_TYPE_ERROR)
    End Function

    Private Function GetNextFreeMusicAPIIndex() As Integer
        GetNextFreeMusicAPIIndex = 0
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeMusicAPIIndex called", LogType.LOG_TYPE_INFO)
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
        Dim LowestFreeIndex As Integer = 0
        Dim IndexFound As Boolean = True
        While IndexFound
            LowestFreeIndex = LowestFreeIndex + 1
            IndexFound = False
            For Each DLNADevice In DLNADevices
                If DLNADevice.Key <> "" Then
                    Try
                        If GetBooleanIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceIsAdded.ToString, False) = True Then
                            If GetIntegerIniFile(DLNADevice.Key, DeviceInfoIndex.diMusicAPIIndex.ToString, 0) = LowestFreeIndex Then
                                IndexFound = True
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetNextFreeMusicAPIIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Next
            If LowestFreeIndex > 100 Then
                IndexFound = False ' force an exit
                LowestFreeIndex = 0
            End If
        End While
        GetNextFreeMusicAPIIndex = LowestFreeIndex
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeMusicAPIIndex found Index = " & LowestFreeIndex.ToString, LogType.LOG_TYPE_INFO)
    End Function

    Private Function GetNextFreeDeviceIndex() As Integer
        GetNextFreeDeviceIndex = 0
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeDeviceIndex called", LogType.LOG_TYPE_INFO)
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
        Dim LowestFreeIndex As Integer = 0
        Dim IndexFound As Boolean = True
        While IndexFound
            LowestFreeIndex = LowestFreeIndex + 1
            IndexFound = False
            For Each DLNADevice In DLNADevices
                If DLNADevice.Key <> "" Then
                    Try
                        If GetBooleanIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceIsAdded.ToString, False) = True Then
                            If GetIntegerIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceAPIIndex.ToString, 0) = LowestFreeIndex Then
                                IndexFound = True
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetNextFreeDeviceIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Next
            If LowestFreeIndex > 100 Then
                IndexFound = False ' force an exit
                LowestFreeIndex = 0
            End If
        End While
        GetNextFreeDeviceIndex = LowestFreeIndex
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeDeviceIndex found Index = " & LowestFreeIndex.ToString, LogType.LOG_TYPE_INFO)
    End Function


    Public Function NumInstances() As Integer
        ' Returns the number of instances this plug-in supports. Plug-ins that support multiple instances probably support multiple output devices and one music library. 
        'Normally, plug-ins return 1 for this value.
        If Not gIOEnabled Then
            ' we're not intialized yet, this is a problem. 
            Log("NumInstances was called before end of initialization. Waiting for 10 seconds", LogType.LOG_TYPE_WARNING)
            wait(10)
        End If
        NumInstances = 0
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.UPnPDeviceMusicAPIIndex > NumInstances Then
                NumInstances = HSDevice.UPnPDeviceMusicAPIIndex
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NumInstances called. Instances is " & NumInstances.ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetInstanceName(ByVal Instance As Integer) As String ' this is important for HST else it won't work!!
        ' Returns the name of this instance as set in the plug-in configuration.
        GetInstanceName = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetInstanceName called with value : " & Instance.ToString, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        Dim UPnPDevice As HSPI
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            UPnPDevice = HSDevice.UPnPDeviceControllerRef
            If Not UPnPDevice Is Nothing Then
                If UPnPDevice.APIInstance = Instance Then
                    GetInstanceName = UPnPDevice.DeviceName
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetInstanceName called with value : " & Instance.ToString & " and Name = " & GetInstanceName, LogType.LOG_TYPE_INFO)
                    UPnPDevice = Nothing
                    Exit Function
                End If
            End If
        Next
    End Function

    Private Sub SendEventForAllZones()
        ' 	generate some event from all players to get ipad/iphone clients updated when they come back on-line
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("SendEventForAllZones called. We have " & MyHSDeviceLinkedList.Count.ToString & " devices", LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then Exit Sub
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If Not HSDevice.UPnPDeviceControllerRef Is Nothing And HSDevice.UPnPDeviceMusicAPIIndex <> 0 Then
                    HSDevice.UPnPDeviceControllerRef.PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, HSDevice.UPnPDeviceControllerRef.PlayerState, False)
                    HSDevice.UPnPDeviceControllerRef.PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
                End If
            Next
        Catch ex As Exception
            Log("Error in SendEventForAllZones with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetSonyPartySingers() As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("GetSonyPartySingers called. We have " & MyHSDeviceLinkedList.Count.ToString & " devices", LogType.LOG_TYPE_INFO)
        GetSonyPartySingers = ""
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If Not HSDevice.UPnPDeviceControllerRef Is Nothing And HSDevice.UPnPDeviceMusicAPIIndex <> 0 Then
                    If HSDevice.UPnPDeviceControllerRef.SonyPartyState <> "" Then
                        ' this device support Sony Party
                        If HSDevice.UPnPDeviceControllerRef.SonyPartyState = "SINGING" Then
                            GetSonyPartySingers = HSDevice.UPnPDeviceControllerRef.DeviceUDN
                            Exit Function
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in GetSonyPartySingers with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("Warning in GetSonyPartySingers. Did not find a Singer", LogType.LOG_TYPE_WARNING)
    End Function

    Public Function GetSonyPartyListeners() As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("GetSonyPartyListeners called. We have " & MyHSDeviceLinkedList.Count.ToString & " devices", LogType.LOG_TYPE_INFO)
        GetSonyPartyListeners = ""
        If MyHSDeviceLinkedList.Count = 0 Then Exit Function
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If Not HSDevice.UPnPDeviceControllerRef Is Nothing And HSDevice.UPnPDeviceMusicAPIIndex <> 0 Then
                    If HSDevice.UPnPDeviceControllerRef.SonyPartyState <> "" Then
                        ' this device support Sony Party
                        If HSDevice.UPnPDeviceControllerRef.SonyPartyState = "LISTENING" Then
                            GetSonyPartyListeners = HSDevice.UPnPDeviceControllerRef.DeviceUDN
                            Exit Function
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in GetSonyPartyListeners with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("Warning in GetSonyPartyListeners. Did not find a Singer", LogType.LOG_TYPE_WARNING)
    End Function

    Public Sub DisplayUPnPDevices()

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DisplayUPnPDevices called", LogType.LOG_TYPE_INFO)

        'Dim UPnpDeviceFinder As New MyUPnPDeviceFinder
        Dim UPnPDevices As MyUPnPDevices = MySSDPDevice.GetAllDevices

        If UPnPDevices Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("No device found in DisplayUPnPDevices.", LogType.LOG_TYPE_WARNING)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DisplayUPnPDevices - Discovery succeeded: " & UPnPDevices.Count & " UPnPDevice(s) found.", LogType.LOG_TYPE_INFO)

        Dim SearchResultTxTFilePath As String = CurrentAppPath & SearchResultFile & ".txt"

        Try
            ' Delete the file if it exists.
            If File.Exists(SearchResultTxTFilePath) Then
                File.Delete(SearchResultTxTFilePath)
            End If
        Catch ex As Exception
        End Try

        'Create the files.
        Dim fs As FileStream = File.Create(SearchResultTxTFilePath)

        If UPnPDevices.Count > 0 Then
            For Each Device In UPnPDevices
                If Device IsNot Nothing Then
                    Try
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DisplayUPnPDevices found device = " & Device.FriendlyName, LogType.LOG_TYPE_INFO)
                        AddText(fs, "DisplayUPnPDevices found device = " & Device.FriendlyName & Environment.NewLine)
                        Dim UPnPDocumentURL As String = ""
                        UPnPDocumentURL = Device.Location

                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    UDN = " & Device.UniqueDeviceName, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   UDN = " & Device.UniqueDeviceName & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    IPAddress = " & Device.IPAddress, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   IPAddress = " & Device.IPAddress & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    IPPort = " & Device.IPPort, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   IPPort = " & Device.IPPort & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    Location = " & Device.Location, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   Location = " & Device.Location & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    ModelName = " & Device.ModelName, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   ModelName = " & Device.ModelName & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    ModelNumber = " & Device.ModelNumber, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   ModelNumber = " & Device.ModelNumber & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    Alive = " & Device.Alive, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   Alive = " & Device.Alive & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    Server = " & Device.Server, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   Server = " & Device.Server & Environment.NewLine)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log(".    TimeoutValue = " & Device.CacheControl, LogType.LOG_TYPE_INFO)
                        AddText(fs, "   TimeoutValue = " & Device.CacheControl & Environment.NewLine)

                        Dim SearchResultXMLFilePath As String = CurrentAppPath & SearchResultFile & "_" & ReplaceSpecialCharacters(Device.UniqueDeviceName) & ".xml"
                        Try
                            ' Delete the file if it exists.
                            If File.Exists(SearchResultXMLFilePath) Then
                                File.Delete(SearchResultXMLFilePath)
                            End If
                            Dim xmlfs As FileStream = File.Create(SearchResultXMLFilePath)
                            AddText(xmlfs, Device.DeviceUPnPDocument.ToString)
                            xmlfs.Close()
                        Catch ex As Exception
                        End Try
                        UPnPDocumentURL = Nothing
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DisplayUPnPDevices did not find the documentURL with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        Dim ParentServices As MyUPnPServices = Device.Services
                        If ParentServices IsNot Nothing Then
                            For Each ParentService As MyUPnPService In ParentServices
                                If ParentService IsNot Nothing Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(".    DisplayUPnPDevices found Service ID = " & ParentService.Id, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "    DisplayUPnPDevices found Service ID = " & ParentService.Id & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ServiceTypeIdentifier = " & ParentService.ServiceTypeIdentifier, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ServiceTypeIdentifier = " & ParentService.ServiceTypeIdentifier & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ControlURL = " & ParentService.MycontrolURL, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ControlURL = " & ParentService.MycontrolURL & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        EventSubURL = " & ParentService.MyeventSubURL, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       EventSubURL = " & ParentService.MyeventSubURL & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ReceivedSID = " & ParentService.MyReceivedSID, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ReceivedSID = " & ParentService.MyReceivedSID & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        MySCPDURL = " & ParentService.MySCPDURL, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       MySCPDURL = " & ParentService.MySCPDURL & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ServiceActive = " & ParentService.hasActionListRetrieved, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ServiceActive = " & ParentService.hasActionListRetrieved & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        Timeout = " & ParentService.MyTimeout, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       Timeout = " & ParentService.MyTimeout & Environment.NewLine)
                                    Dim SearchResultXMLFilePath As String = CurrentAppPath & SearchResultFile & "_" & ReplaceSpecialCharacters(Device.UniqueDeviceName) & "_" & ReplaceSpecialCharacters(ParentService.Id) & ".xml"
                                    Try
                                        Dim xmldoc As XmlDocument = New XmlDocument()
                                        xmldoc.Load(ParentService.MySCPDURL)
                                        ' Delete the file if it exists.
                                        If File.Exists(SearchResultXMLFilePath) Then
                                            File.Delete(SearchResultXMLFilePath)
                                        End If
                                        If xmldoc IsNot Nothing Then
                                            Dim xmlfs As FileStream = File.Create(SearchResultXMLFilePath)
                                            AddText(xmlfs, xmldoc.OuterXml)
                                            xmlfs.Close()
                                        End If
                                    Catch ex As Exception
                                    End Try
                                End If
                            Next
                        End If
                    Catch ex As Exception
                    End Try
                    Dim Children As MyUPnPDevices
                    Dim ChildServices As MyUPnPServices
                    Dim DeviceName As String = Device.FriendlyName
                    Children = Device.Children
                    Try
                        If Children IsNot Nothing Then
                            For Each Child As MyUPnPDevice In Children
                                Try
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(".    DisplayUPnPDevices for device = " & DeviceName & " found Child = " & Child.FriendlyName, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "    DisplayUPnPDevices for device = " & DeviceName & " found Child = " & Child.FriendlyName & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        UDN = " & Child.UniqueDeviceName, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       UDN = " & Child.UniqueDeviceName & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        IPAddress = " & Child.IPAddress, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       IPAddress = " & Child.IPAddress & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        IPPort = " & Child.IPPort, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       IPPort = " & Child.IPPort & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        Location = " & Child.Location, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       Location = " & Child.Location & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ModelName = " & Child.ModelName, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ModelName = " & Child.ModelName & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        ModelNumber = " & Child.ModelNumber, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       ModelNumber = " & Child.ModelNumber & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        Alive = " & Child.Alive, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       Alive = " & Child.Alive & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        Server = " & Child.Server, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       Server = " & Child.Server & Environment.NewLine)
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        TimeoutValue = " & Child.CacheControl, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "       TimeoutValue = " & Child.CacheControl & Environment.NewLine)
                                    Dim UPnPDocumentURL As String = ""
                                    UPnPDocumentURL = Child.Location
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(".        DisplayUPnPDevices for device = " & DeviceName & " and Child = " & Child.UniqueDeviceName & " found documentURL = " & UPnPDocumentURL.ToString, LogType.LOG_TYPE_INFO)
                                    AddText(fs, "        DisplayUPnPDevices for device = " & DeviceName & " and Child = " & Child.UniqueDeviceName & " found documentURL = " & UPnPDocumentURL.ToString & Environment.NewLine)
                                    Dim SearchResultXMLFilePath As String = CurrentAppPath & SearchResultFile & "_" & ReplaceSpecialCharacters(Child.FriendlyName) & ".xml"
                                    Try
                                        ' Delete the file if it exists.
                                        If File.Exists(SearchResultXMLFilePath) Then
                                            File.Delete(SearchResultXMLFilePath)
                                        End If
                                        Dim xmlfs As FileStream = File.Create(SearchResultXMLFilePath)
                                        AddText(xmlfs, Child.DeviceUPnPDocument.ToString)
                                        xmlfs.Close()
                                    Catch ex As Exception
                                    End Try
                                    UPnPDocumentURL = Nothing
                                Catch ex As Exception
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DisplayUPnPDevices for device = " & DeviceName & " did not find the Child documentURL with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                Try
                                    ChildServices = Child.Services
                                    If ChildServices IsNot Nothing Then
                                        For Each ChildService In ChildServices
                                            If ChildService IsNot Nothing Then
                                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(".        DisplayUPnPDevices found Service ID = " & ChildService.Id, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "        DisplayUPnPDevices found Service ID = " & ChildService.Id & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            ServiceTypeIdentifier = " & ChildService.ServiceTypeIdentifier, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           ServiceTypeIdentifier = " & ChildService.ServiceTypeIdentifier & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            ControlURL = " & ChildService.MycontrolURL, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           ControlURL = " & ChildService.MycontrolURL & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            EventSubURL = " & ChildService.MyeventSubURL, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           EventSubURL = " & ChildService.MyeventSubURL & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            ReceivedSID = " & ChildService.MyReceivedSID, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           ReceivedSID = " & ChildService.MyReceivedSID & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            MySCPDURL = " & ChildService.MySCPDURL, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           MySCPDURL = " & ChildService.MySCPDURL & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            ServiceActive = " & ChildService.hasActionListRetrieved, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           ServiceActive = " & ChildService.hasActionListRetrieved & Environment.NewLine)
                                                If PIDebuglevel > DebugLevel.dlEvents Then Log(".            Timeout = " & ChildService.MyTimeout, LogType.LOG_TYPE_INFO)
                                                AddText(fs, "           Timeout = " & ChildService.MyTimeout & Environment.NewLine)
                                                Dim SearchResultXMLFilePath As String = CurrentAppPath & SearchResultFile & "_" & ReplaceSpecialCharacters(Device.UniqueDeviceName) & "_" & ReplaceSpecialCharacters(ChildService.Id) & ".xml"
                                                Try
                                                    Dim xmldoc As XmlDocument = New XmlDocument()
                                                    xmldoc.Load(ChildService.MySCPDURL)
                                                    ' Delete the file if it exists.
                                                    If File.Exists(SearchResultXMLFilePath) Then
                                                        File.Delete(SearchResultXMLFilePath)
                                                    End If
                                                    If xmldoc IsNot Nothing Then
                                                        Dim xmlfs As FileStream = File.Create(SearchResultXMLFilePath)
                                                        AddText(xmlfs, xmldoc.OuterXml)
                                                        xmlfs.Close()
                                                    End If
                                                Catch ex As Exception
                                                End Try
                                            End If
                                        Next
                                    End If
                                Catch ex As Exception
                                    Log("Error in DisplayUPnPDevices for device  = " & DeviceName & " and Child = " & Child.UniqueDeviceName & " with error=" & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Next
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Next
        End If
        Try
            fs.Close()
        Catch ex As Exception
            Log("Error in DisplayUPnPDevices closing the SearchResultFile with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DisplayUPnPDevices done!", LogType.LOG_TYPE_INFO)

    End Sub

    Private Shared Sub AddText(ByVal fs As FileStream, ByVal value As String)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
        fs.Write(info, 0, info.Length)
    End Sub

    Public Function GetPlayLists() As System.Array
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPlayLists called ", LogType.LOG_TYPE_INFO)
        GetPlayLists = Nothing
        Dim SList() As String = {""}
        Dim KeyIndex As Integer = 0
        Try
            Dim path As String = CurrentAppPath & gPlaylistPath
            If File.Exists(path) Then
                ' This path is a file.
                Log("Error in GetPlayLists. An file was passed instead of a Directory = " & CurrentAppPath & gPlaylistPath, LogType.LOG_TYPE_ERROR)
            Else
                If Directory.Exists(path) Then
                    ' This path is a directory.
                    Dim fileEntries As String() = Directory.GetFiles(path)
                    ' Process the list of files found in the directory.
                    Dim fileName As String
                    For Each fileName In fileEntries
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetPlayLists found file = " & fileName, LogType.LOG_TYPE_INFO)
                        fileName = fileName.Substring(fileName.LastIndexOf("\") + 1, fileName.Length - fileName.LastIndexOf("\") - 1) ' remove the path stuff
                        Dim FilePrefix As String = fileName.Substring(0, fileName.LastIndexOf(".")) ' pick out the file extension
                        Dim FileExtention As String = fileName.Substring(fileName.LastIndexOf(".") + 1, fileName.Length - fileName.LastIndexOf(".") - 1) ' pick out the file extension
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetPlayLists found FilePrefix = " & FilePrefix & " and Extention = " & FileExtention, LogType.LOG_TYPE_INFO)
                        Dim DeviceName As String = GetStringIniFile(FilePrefix, DeviceInfoIndex.diGivenName.ToString, "")
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetPlayLists found GivenName = " & DeviceName, LogType.LOG_TYPE_INFO)
                        Dim QEntry As String = ""
                        If DeviceName <> "" Then
                            QEntry = "Player_" & DeviceName & ":;:-:" & FilePrefix
                        Else
                            QEntry = FilePrefix & ":;:-:" & FilePrefix
                        End If
                        ReDim Preserve SList(KeyIndex)
                        SList(KeyIndex) = QEntry
                        KeyIndex = KeyIndex + 1
                    Next fileName
                Else
                    Log("Error in GetPlayLists. An invalid directory was passed. Directory = " & CurrentAppPath & gPlaylistPath, LogType.LOG_TYPE_ERROR)
                End If
            End If
        Catch ex As Exception
            Log("Error in GetPlayLists with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If KeyIndex > 0 Then
            GetPlayLists = SList
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPlayLists called and returned " & KeyIndex.ToString & " entries.", LogType.LOG_TYPE_INFO)
        End If
        SList = Nothing
    End Function

#End Region

#Region "HS Device functions"

    Private Function DMAdd() As MyUPnpDeviceInfo
        '   MyPingAddressLinkedList.AddLast(NewArrayElement)
        Dim NewArrayElement As New MyUPnpDeviceInfo
        MyHSDeviceLinkedList.AddLast(NewArrayElement)
        DMAdd = NewArrayElement
    End Function

    Private Sub DMRemove(UDN As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DMRemove called with UDN = " & UDN, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then Exit Sub
        For Each UPnPDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If UPnPDevice.UPnPDeviceUDN = UDN Then
                If Not UPnPDevice.UPnPDeviceControllerRef Is Nothing Then
                    'UPnPDevice.UPnPDeviceControllerRef.Dispose()
                    UPnPDevice.UPnPDeviceControllerRef = Nothing
                End If
                UPnPDevice.Close()
                MyHSDeviceLinkedList.Remove(UPnPDevice)
                Exit Sub
            End If
        Next
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DMRemove. Could not find UDN = " & UDN, LogType.LOG_TYPE_ERROR)
    End Sub
#End Region

#Region "    Speaker Proxy Related Procedures    "
    Public Sub SpeakIn(device As Integer, text As String, wait As Boolean, host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn
        'Public Sub SpeakIn(ByVal device As Short, ByVal text As String, ByVal wait As Boolean, ByVal host As String)
        Log("SpeakIn called for Device = " & device.ToString & ", Text = " & text & ", Wait=" & wait.ToString & ", Host = " & host, LogType.LOG_TYPE_INFO)
        If Not MyPIisInitialized Then
            If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, host)
            Exit Sub
        End If
        Dim HostDevice As String
        HostDevice = host
        HostDevice = Trim(HostDevice)
        If MyHSDeviceLinkedList.Count = 0 Then
            If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, HostDevice)
            Exit Sub
        End If
        Dim SpeakerClientList As String() = Nothing
        SpeakerClientList = Split(HostDevice, ",")
        Dim FoundOne As Boolean = False
        If HostDevice <> "" Then
            For Each SpeakerClient As String In SpeakerClientList
                If Mid(SpeakerClient, 1, 4) = "$MC$" Then
                    ' this is for the plug-in. The next "$" character ends the Linkgroup ZoneName
                    Mid(SpeakerClient, 1, 4) = "    "
                    SpeakerClient = Trim(SpeakerClient)
                    Dim PlayerName As String
                    Dim Delimiter As Integer
                    Delimiter = SpeakerClient.IndexOf("$")
                    If Delimiter = 0 Then
                        ' should not be, just pass it along
                        If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, SpeakerClient)
                        'Exit Sub
                    Else
                        PlayerName = SpeakerClient.Substring(0, Delimiter)
                        'Dim PlayerUDN As String = GetUDNByDeviceGivenName(PlayerName)
                        ' remove the PlayerName from the Host
                        SpeakerClient = SpeakerClient.Remove(0, Delimiter + 1)
                        SpeakerClient = Trim(SpeakerClient)
                        If SpeakerClient = ":*" Then SpeakerClient = "*:*"
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SpeakerProxy activated with HostName = " & SpeakerClient & " Text = " & text & " and PlayerName = " & PlayerName, LogType.LOG_TYPE_INFO)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SpeakIn is looking at " & MyHSDeviceLinkedList.Count.ToString & " UPnPDevices", LogType.LOG_TYPE_INFO)
                        Dim Index As Integer = 0
                        Dim UPnPDevice As HSPI = Nothing
                        Try
                            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                                UPnPDevice = HSDevice.UPnPDeviceControllerRef
                                If Not UPnPDevice Is Nothing And (UPnPDevice.DeviceName.ToUpper = PlayerName.ToUpper) Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SpeakIn found UPnPDevice = " & UPnPDevice.DeviceName & " in the UPnPDeviceInfoArray with DeviceServiceType= " & UPnPDevice.DeviceServiceType & " and RemoteType = " & GetStringIniFile(UPnPDevice.DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, ""), LogType.LOG_TYPE_INFO)
                                    If UPnPDevice.DeviceServiceType = "PMR" Then
                                        UPnPDevice.SamsungSendMessage("SMS", text)
                                        FoundOne = True
                                        Exit For
                                        'Exit Sub
                                    ElseIf UPnPDevice.DeviceServiceType = "RCR" Then
                                        Dim RemoteType As String = GetStringIniFile(UPnPDevice.DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                                        ' some LGs support Toasts
                                        If RemoteType = "LG" Then
                                            UPnPDevice.LGSendMessage(text)
                                            FoundOne = True
                                            Exit For
                                        End If
                                    ElseIf UPnPDevice.DeviceServiceType = "DMR" Then
                                        AddAnnouncementToQueue("uuid:  " & UPnPDevice.DeviceUDN, device, text, wait, SpeakerClient, True)
                                        DoCheckAnnouncementQueue()
                                        FoundOne = True
                                        Exit For
                                        'Exit Sub
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            Log("Error in SpeakIn trying to find the device in the array with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        'If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, SpeakerClient)
                        'Exit Sub
                    End If
                ElseIf (Mid(SpeakerClient, 1, 7) = "$SONOS$") Then
                    ' don't do anything
                    'Exit Sub
                End If
            Next
        Else
            'Dim FoundOne As Boolean = False
            Dim SpeakerDevices As New System.Collections.Generic.Dictionary(Of String, String)()
            SpeakerDevices = GetIniSection("Speaker Devices") '  As Dictionary(Of String, String)
            Try
                If SpeakerDevices IsNot Nothing Then
                    For Each SpeakerDevice In SpeakerDevices
                        If SpeakerDevice.Value <> "" Then
                            Dim Speakers As String()
                            Speakers = Split(SpeakerDevice.Value, ",")
                            If Not Speakers Is Nothing Then
                                Dim DeviceID As String
                                For Each DeviceID In Speakers
                                    'log( "SpeakIn found DeviceID = " & DeviceID & " for UDN = " & SpeakerDevice.Key & " while looking for = " & device.ToString)
                                    If Trim(DeviceID) = device.ToString Then
                                        AddAnnouncementToQueue("uuid:" & SpeakerDevice.Key, device, text, wait, HostDevice, True)
                                        DoCheckAnnouncementQueue()
                                        FoundOne = True
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                Log("Error in SpeakIn looking for Speaker devices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Dim MessageDevices As New System.Collections.Generic.Dictionary(Of String, String)()
            MessageDevices = GetIniSection("Message Devices") '  As Dictionary(Of String, String)
            Try
                For Each MessageDevice In MessageDevices
                    If MessageDevice.Value <> "" Then
                        Dim Messages As String()
                        Messages = Split(MessageDevice.Value, ",")
                        If Not Messages Is Nothing Then
                            Dim DeviceID As String
                            For Each DeviceID In Messages
                                If Trim(DeviceID) = device.ToString Then
                                    Dim UPnPDevice As HSPI = Nothing
                                    UPnPDevice = GetAPIByUDN(MessageDevice.Key)
                                    If Not UPnPDevice Is Nothing Then
                                        UPnPDevice.SamsungSendMessage("SMS", text)
                                        FoundOne = True
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Catch ex As Exception
                Log("Error in SpeakIn looking for Message devices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

        End If

        If Not FoundOne Then
            If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, HostDevice)
        Else
            If MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysForward Then hs.SpeakProxy(device, text, wait, HostDevice)
        End If

    End Sub

    Private Sub AddAnnouncementToQueue(ByVal PlayerName As String, ByVal device As Short, ByVal text As String, ByVal wait As Boolean, ByVal host As String, ByVal IsFile As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddAnnouncementToQueue called for PlayerName = " & PlayerName & " and Text = " & text, LogType.LOG_TYPE_INFO)
        Dim AnnouncementItem As New AnnouncementItems
        AnnouncementItem.device = device
        AnnouncementItem.text = text
        AnnouncementItem.wait = wait
        AnnouncementItem.host = host
        AnnouncementItem.IsFile = IsFile
        AnnouncementItem.SourceZoneMusicAPI = GetLinkgroupSourceZone(PlayerName)
        If Mid(PlayerName, 1, 5) = "uuid:" Then
            AnnouncementItem.LinkGroupName = AnnouncementItem.SourceZoneMusicAPI.DeviceName
        Else
            AnnouncementItem.LinkGroupName = PlayerName
        End If
        Dim LastAnnouncementInQueue As AnnouncementItems
        Try
            LastAnnouncementInQueue = GetTailOfAnnouncementQueue()
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue getting tail with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            AnnouncementItem.Previous_ = LastAnnouncementInQueue
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue setting previous with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            If LastAnnouncementInQueue Is Nothing Then
                ' this is the very first
                AnnouncementLink = AnnouncementItem
            Else
                LastAnnouncementInQueue.Next_ = AnnouncementItem
            End If
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue setting Next with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        ' moved this to the end of procedure to avoid race condition where TO services calls the DoCheckAnnouncementqueue before it was set up.
        ' changed in v.83
        AnnouncementsInQueue = True
        MyTimeoutActionArray(TOCheckAnnouncement) = 1
        MyAnnouncementCountdown = MyMaxAnnouncementTime
    End Sub

    Private Function GetTailOfAnnouncementQueue() As AnnouncementItems
        GetTailOfAnnouncementQueue = AnnouncementLink
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called", LogType.LOG_TYPE_INFO)
        If AnnouncementLink Is Nothing Then
            Exit Function
        End If
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink
        Dim LoopIndex As Integer = 0
        Try
            Do
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called and found Linkgroup = " & AnnouncementItem.LinkGroupName & " and text = " & AnnouncementItem.text, LogType.LOG_TYPE_INFO)
                If AnnouncementItem.Next_ Is Nothing Then
                    GetTailOfAnnouncementQueue = AnnouncementItem
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called and tail found", LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
                AnnouncementItem = AnnouncementItem.Next_
                LoopIndex = LoopIndex + 1
                If LoopIndex > 100 Then
                    ' we have a loop, force clean it
                    Log("Error in GetTailOfAnnouncementQueue, loop found, clearing all Announcement info", LogType.LOG_TYPE_ERROR)
                    AnnouncementLink = Nothing
                    Exit Function
                End If
            Loop
        Catch ex As Exception
            Log("Error in GetTailOfAnnouncementQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub DeleteHeadOfAnnouncementQueue()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteHeadOfAnnouncementQueue called", LogType.LOG_TYPE_INFO)
        If AnnouncementLink Is Nothing Then
            ' this should not be!
            Exit Sub
        End If
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink.Next_
        Try
            AnnouncementLink.Next_ = Nothing ' make sure there are no references left
            AnnouncementLink.SourceZoneMusicAPI = Nothing
            AnnouncementLink = Nothing ' return this memory
            AnnouncementLink = AnnouncementItem
            If Not AnnouncementLink Is Nothing Then
                AnnouncementLink.Previous_ = Nothing
            End If
        Catch ex As Exception
            Log("Error in DeleteHeadOfAnnouncementQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DoCheckAnnouncementQueue()
        If Not AnnouncementsInQueue Then Exit Sub
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DoCheckAnnouncementQueue called AnnouncementinQueue = " & AnnouncementsInQueue.ToString & " and AnnouncementInProgress = " & AnnouncementInProgress.ToString & " and AnnouncementCountdown = " & MyAnnouncementCountdown.ToString)
        If AnnouncementReEntry Then
            'log( "DoCheckAnnouncementQueue called and cause re-entry")
            Exit Sub ' re-entry
        End If

        AnnouncementReEntry = True
        If AnnouncementInProgress Then
            If Not AnnouncementLink Is Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue has link and AnnouncementState = " & AnnouncementLink.State_.ToString, LogType.LOG_TYPE_INFO)
                If AnnouncementLink.State_ = AnnouncementState.asLinking Then
                    ' this is re-entrance
                    AnnouncementReEntry = False
                    Exit Sub
                End If
                If AnnouncementLink.IsFile Then
                    MyAnnouncementCountdown = MyAnnouncementCountdown - 1
                    If MyAnnouncementCountdown < 0 Then
                        ' OK this really not good. A 100 seconds has gone by since the last announcement was added and we are still linked
                        AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Log("Error in DoCheckAnnouncementQueue. " & MyMaxAnnouncementTime.ToString & " seconds expired since the announcement started and no end was received.", LogType.LOG_TYPE_ERROR)
                    ElseIf Not AnnouncementLink.SourceZoneMusicAPI Is Nothing Then
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue has PlayerState = " & AnnouncementLink.SourceZoneMusicAPI.NoQueuePlayerState.ToString & " and HasAnnouncementStarted = " & AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted.ToString, LogType.LOG_TYPE_INFO)
                        If (AnnouncementLink.SourceZoneMusicAPI.NoQueuePlayerState = player_state_values.Stopped Or AnnouncementLink.SourceZoneMusicAPI.NoQueuePlayerState = player_state_values.Transitioning) And AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted Then
                            ' OK the announcement is over
                            AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Else
                            AnnouncementReEntry = False
                            Exit Sub
                        End If
                    Else
                        AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Log("Error in DoCheckAnnouncementQueue. Timer is running but there is no instance of a Player Object.", LogType.LOG_TYPE_ERROR)
                    End If
                Else
                    AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                    Log("Error in DoCheckAnnouncementQueue. Timer is running but Announcement is not marked as Speak-To-File", LogType.LOG_TYPE_ERROR)
                End If
            Else
                ' this should not be
                AnnouncementsInQueue = False
                AnnouncementInProgress = False
                MyAnnouncementIndex = 0
                AnnouncementReEntry = False
                Log("Error in DoCheckAnnouncementQueue. No Announcement info, empty link", LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
        End If

        AnnouncementInProgress = True

        If AnnouncementLink Is Nothing Then
            ' this should not be
            AnnouncementsInQueue = False
            AnnouncementInProgress = False
            MyAnnouncementIndex = 0
            AnnouncementReEntry = False
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue. No AnnouncementLink", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called for linkgroup " & AnnouncementLink.LinkGroupName & " and State = " & AnnouncementLink.State_.ToString & " and isFile " & AnnouncementLink.IsFile.ToString, LogType.LOG_TYPE_INFO)
        ' Look at first announcement in Queue
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink
        If AnnouncementItem.State_ = AnnouncementState.asIdle Then
            AnnouncementItem.State_ = AnnouncementState.asLinking
            'log( "DoCheckAnnouncementQueue called HandleLinking")
            HandleLinkingOn(AnnouncementItem.LinkGroupName, AnnouncementLink.IsFile)
            'log( "DoCheckAnnouncementQueue done calling HandleLinking")
            If AnnouncementItem.IsFile Then
                Try
                    AnnouncementItem.SourceZoneMusicAPI.ClearQueue()
                Catch ex As Exception
                End Try
                ' Also reset shuffle and repeat to avoid reordering and endless repeats
                Try
                    AnnouncementItem.SourceZoneMusicAPI.PlayModeNormal()
                Catch ex As Exception
                End Try
            End If
            AnnouncementItem.State_ = AnnouncementState.asLinked
        ElseIf AnnouncementItem.State_ = AnnouncementState.asLinking Then
            ' this is re-entrance
            AnnouncementReEntry = False
            Exit Sub
        End If
        Dim StartQueueIndex As Integer = MyAnnouncementIndex + 1
        If AnnouncementItem.State_ = AnnouncementState.asLinked Then
            Dim TextStrings As String() = Nothing
            AnnouncementItem.text.Split("|")
            TextStrings = AnnouncementItem.text.Split("|")
            Dim Index As Integer = 0
            'log( "DoCheckAnnouncementQueue activated with HostName = " & AnnouncementItem.host & " Text = " & AnnouncementItem.text & " and LinkgroupName = " & AnnouncementItem.LinkGroupName)

            For Each TextString In TextStrings
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue activated with HostName = " & AnnouncementItem.host & " Text = " & TextString & " and LinkgroupName = " & AnnouncementItem.LinkGroupName, LogType.LOG_TYPE_INFO)
                If AnnouncementItem.IsFile Then
                    Dim FileName As String
                    Dim ExtensionIndex As Integer = 0
                    Dim Extensiontype As String = ""
                    Dim Path As String
                    If HSisRunningOnLinux Then ' this will always be on the HS machine
                        Path = CurrentAppPath & "/html" & AnnouncementPath
                    Else
                        Path = CurrentAppPath & "\html" & AnnouncementPath
                    End If
                    FileName = "Ann_" & RemoveBlanks(AnnouncementItem.LinkGroupName) & "_" & MyAnnouncementIndex.ToString
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue adds file = " & Path & FileName & " to Queue", LogType.LOG_TYPE_INFO)
                    If File.Exists(TextString) Then ' this needs to be fixed for remoting tralala
                        Try
                            ' get the extension file type
                            ExtensionIndex = TextString.LastIndexOf(".")
                            If ExtensionIndex <> -1 Then
                                Extensiontype = TextString.Substring(ExtensionIndex, TextString.Length - ExtensionIndex)
                            End If
                            FileName = FileName + Extensiontype
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when searching for the file type with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            System.IO.File.Delete(Path & FileName)
                        Catch ex As Exception
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Error in DoCheckAnnouncementQueue when deleting file " & Path & FileName & " Error = " &ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            System.IO.File.Copy(TextString, Path & FileName, True)
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DoCheckAnnouncementQueue copying file " & Path & FileName)
                            AnnouncementItem.State_ = AnnouncementState.asSpeaking
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue in SpeakToFile copying file = " & TextString & " to " & Path & FileName & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                            AnnouncementInProgress = False
                            AnnouncementItem.State_ = AnnouncementState.asFilePlayed
                            AnnouncementReEntry = False
                            Exit Sub
                        End Try
                    Else
                        If GetBooleanIniFile(AnnouncementItem.SourceZoneMusicAPI.DeviceUDN, DeviceInfoIndex.diAnnouncementMP3.ToString, False) Then
                            ' convert to MP.3
                            Extensiontype = ".mp3"
                        Else
                            Extensiontype = ".wav"
                        End If
                        Try
                            System.IO.File.Delete(Path & FileName & ".wav")
                            'log( "DoCheckAnnouncementQueue deleted " & Path & FileName & " successfully")
                        Catch ex As Exception
                            'log( "DoCheckAnnouncementQueue deleted " & Path & FileName & " un-successfully with error " &ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        AnnouncementItem.State_ = AnnouncementState.asSpeaking
                        Try
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue calling SpeakToFile with Text " & TextString & " and File " & Path & FileName, LogType.LOG_TYPE_INFO)
                            Dim Voice As String = CheckForVoiceTag(TextString)
                            hs.SpeakToFile(TextString, Voice, Path & FileName & ".wav")
                            If Extensiontype = ".mp3" And Not HSisRunningOnLinux Then
                                Dim strOutput As String = ""
                                Dim strError As String = ""
                                Try
                                    Dim objProcess As New Process()
                                    ' Start the Command and redirect the output
                                    objProcess.StartInfo.UseShellExecute = False
                                    objProcess.StartInfo.RedirectStandardOutput = True
                                    objProcess.StartInfo.CreateNoWindow = True
                                    objProcess.StartInfo.RedirectStandardError = True
                                    If HSisRunningOnLinux Then
                                        objProcess.StartInfo.FileName() = CurrentAppPath & "/lame" '"/html/" & sIFACE_NAME & "/lame"
                                    Else
                                        objProcess.StartInfo.FileName() = CurrentAppPath & "\lame" '"\html\" & sIFACE_NAME & "\lame"
                                    End If
                                    objProcess.StartInfo.Arguments() = FileName & ".wav" & " " & FileName & ".mp3"
                                    objProcess.StartInfo.WorkingDirectory = Path

                                    objProcess.Start()

                                    strOutput = objProcess.StandardOutput.ReadToEnd()
                                    strError = objProcess.StandardError.ReadToEnd()
                                    objProcess.WaitForExit()
                                    objProcess.Dispose()
                                Catch ex As Exception
                                    Log("Error in DoCheckAnnouncementQueue finished conversion to mp3 with Ouptut = " & strOutput & " and Return = " & strError, LogType.LOG_TYPE_ERROR)
                                    Log("Error in DoCheckAnnouncementQueue converting .wav to .mp3 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                FileName = FileName & ".mp3"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue finished conversion to mp3 with Ouptut = " & strOutput & " and Return = " & strError, LogType.LOG_TYPE_INFO)
                            Else
                                FileName = FileName & ".wav"
                            End If

                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue finished SpeakToFile", LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            Log("Error in DoCheckAnnouncementQueue called SpeakToFile unsuccessfully with Text " & TextString & " and File " & Path & FileName & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                    ' Meaning of DLNA.ORG_FLAGS
                    ' Bits [31,21] (inclusive) of the primary-flags token are valid for use (i.e. those bits are valid for use)
                    ' Bits [19,0] of the primary-flags token have undefined values.
                    ' Bit-31: sp-flag (Sender Paced Flag)                                              : should be zero in our case, content source cannot source (player) clocking
                    ' Bit-30: lop-npt (Limited Operations Flags: Time-Based Seek)
                    '       lop-npt: indicates support of the TimeSeekRange.dlna.org HTTP header for the context of the protocolInfo under the "Limited Random Access Data Availability" model
                    ' Bit-29: lop-bytes (Limited Operations Flags: Byte-Based Seek)
                    '       lop-bytes: indicates support of the Range HTTP header for the context of the protocolInfo under the "Limited Random Access Data Availability" model
                    ' Bit-28: playcontainer-param (DLNA PlayContainer Flag)
                    ' Bit 27: s0-increasing (UCDAM s0 Increasing Flag)
                    ' Bit 26: sN-increasing (UCDAM sN Increasing Flag)
                    ' Bit-25: rtsp-pause (Pause media operation support for RTP Serving Endpoints)
                    ' Bit 24: tm-s (Streaming Mode Flag)
                    ' Bit 23: tm-i (Interactive Mode Flag)
                    ' Bit 22: tm-b (Background Mode Flag)
                    ' Bit 21: http-stalling (HTTP Connection Stalling Flag)
                    ' Bit 20: dlna-v1.5-flag (DLNA v1.5 versioning flag)

                    ' ORG_CI = Conversion Indicator = boolean
                    Dim HSIpAddress As String
                    HSIpAddress = hs.GetIPAddress()
                    Dim HTTPPort As String = "" ' found in settings.ini at gWebSvrPort=80
                    HTTPPort = hs.GetINISetting("Settings", "gWebSvrPort", "")
                    If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                    Dim MetaData As String = ""
                    ' if MP3 protocolInfo="http-get:*:audio/mpeg:DLNA.ORG_PN=MP3;DLNA.ORG_OP=11;DLNA.ORG_FLAGS=01700000000000000000000000000000"
                    ' if .Wav = protocolInfo="http-get:*:audio/wav:DLNA.ORG_OP=01;DLNA.ORG_FLAGS=01500000000000000000000000000000"
                    MetaData = "<DIDL-Lite xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:upnp=""urn:schemas-upnp-org:metadata-1-0/upnp/"" xmlns:dlna=""urn:schemas-dlna-org:metadata-1-0/"" xmlns:pv=""http://www.pv.com/pvns/"" xmlns=""urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/""><item id=""-1"" parentID=""-1"" restricted=""true"">"
                    If Extensiontype.ToUpper = ".MP3" Then
                        MetaData = MetaData & "<upnp:class>object.item.audioItem.musicTrack</upnp:class>"
                        'MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/mpeg:DLNA.ORG_PN=MP3;DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000"">http://"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/mpeg:*"">http://"
                        'MetaData = MetaData & "<res protocolInfo=""http-get:*:*:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".JPG" Or Extensiontype.ToUpper = ".JPEG" Then
                        MetaData = MetaData & "<upnp:class>object.item.imageItem.photo</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:image/jpeg:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".GIF" Then
                        MetaData = MetaData & "<upnp:class>object.item.imageItem.photo</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:image/gif:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".PNG" Then
                        MetaData = MetaData & "<upnp:class>object.item.imageItem.photo</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:image/png:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".MPG" Then
                        MetaData = MetaData & "<upnp:class>object.item.videoItem.movie</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:video/mpeg:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".MP4" Then
                        MetaData = MetaData & "<upnp:class>object.item.videoItem.movie</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:video/mp4:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".WMV" Then
                        MetaData = MetaData & "<upnp:class>object.item.videoItem.movie</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:video/x-ms-wmv:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".AVI" Then
                        MetaData = MetaData & "<upnp:class>object.item.videoItem.movie</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:video/avi:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".WMA" Then
                        MetaData = MetaData & "<upnp:class>object.item.audioItem.musicTrack</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/x-ms-wma:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".OGG" Then
                        MetaData = MetaData & "<upnp:class>object.item.audioItem.musicTrack</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/ogg:*"">http://"
                    ElseIf Extensiontype.ToUpper = ".WAV" Then
                        MetaData = MetaData & "<upnp:class>object.item.audioItem.musicTrack</upnp:class>"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:*:*"">http://"
                    Else ' should be .wav
                        MetaData = MetaData & "<upnp:class>object.item.audioItem.musicTrack</upnp:class>"
                        'MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/wav:DLNA.ORG_OP=01;DLNA.ORG_FLAGS=01500000000000000000000000000000"">http://"
                        MetaData = MetaData & "<res protocolInfo=""http-get:*:*:*"">http://"
                        'MetaData = MetaData & "<res protocolInfo=""http-get:*:audio/wav:*"">http://"
                        ' http-get:*:audio/x-ms-wma:DLNA.ORG_PN=WMABASE;DLNA.ORG_FLAGS=9d700000000000000000000000000000
                        ' http-get:*:audio/wav:DLNA.ORG_OP=01;DLNA.ORG_FLAGS=01500000000000000000000000000000"
                    End If
                    MetaData = MetaData & HSIpAddress & HTTPPort & AnnouncementPath & FileName & "</res><upnp:albumArtURI>http://"
                    MetaData = MetaData & HSIpAddress & HTTPPort & ImagesPath & "Announcement.jpg</upnp:albumArtURI><dc:title>"
                    'MetaData = MetaData & AnnouncementTitle & "</dc:title><upnp:class>object.item.audioItem.musicTrack</upnp:class><dc:creator>"
                    MetaData = MetaData & AnnouncementTitle & "</dc:title><dc:creator>"
                    MetaData = MetaData & AnnouncementAuthor & "</dc:creator><upnp:album>"
                    MetaData = MetaData & AnnouncementAlbum & "</upnp:album></item></DIDL-Lite>"
                    If UBound(TextStrings, 1) > 0 Then
                        ' Multiple Announcements, queue them up
                        Try
                            If AnnouncementItem.SourceZoneMusicAPI.AVTSetAVTransportURI("http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, MetaData) <> "OK" Then
                                Log("Error in DoCheckAnnouncementQueue adding a track to the Queue = http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, LogType.LOG_TYPE_ERROR)
                                AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = True
                                AnnouncementReEntry = False
                                Exit Sub
                            End If
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is adding a track to the Queue = http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            Log("Error in DoCheckAnnouncementQueue when adding Announcement to Sonos Queue with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        If Index >= UBound(TextStrings, 1) Then
                            ' Last announcement, start playing
                            AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = False
                            Try
                                If AnnouncementItem.SourceZoneMusicAPI.NoQueuePlayerState <> player_state_values.Playing Then
                                    If StartQueueIndex = 1 Then
                                        If AnnouncementItem.SourceZoneMusicAPI.AVTPlay() <> "OK" Then
                                            Log("Error in DoCheckAnnouncementQueue when calling Play", LogType.LOG_TYPE_ERROR)
                                            AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted = True
                                        End If
                                    Else
                                        ' it either never started or already stopped
                                        'AnnouncementItem.SourceZoneMusicAPI.SeekTrack(StartQueueIndex)
                                        If AnnouncementItem.SourceZoneMusicAPI.AVTPlay() <> "OK" Then
                                            Log("Error in DoCheckAnnouncementQueue when calling Play", LogType.LOG_TYPE_ERROR)
                                            AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted = True
                                        End If
                                    End If

                                End If
                            Catch ex As Exception
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when calling PlayURI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                            MyAnnouncementIndex = MyAnnouncementIndex + 1
                            'hs.waitsecs(1) ' this is to make sure the playerstate has moved to playing before the timeout procedure begins checking for the "end of file" which is player stopped
                            AnnouncementReEntry = False
                            Exit Sub
                        End If
                    Else
                        ' Single announcement
                        AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = False
                        Try

                            'If AnnouncementItem.SourceZoneMusicAPI.AVTSetAVTransportURI("http://" & HSIpAddress & HTTPPort & FileName, MetaData) <> "OK" Then
                            If AnnouncementItem.SourceZoneMusicAPI.AVTSetAVTransportURI("http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, MetaData) <> "OK" Then
                                Log("Error in DoCheckAnnouncementQueue adding a track to the Queue = http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, LogType.LOG_TYPE_ERROR)
                                AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = True
                                AnnouncementReEntry = False
                                Exit Sub
                            End If
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is calling PlayURI with http://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when adding Announcement to Sonos Queue with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            If AnnouncementItem.SourceZoneMusicAPI.AVTPlay() <> "OK" Then
                                Log("Error in DoCheckAnnouncementQueue when calling Play", LogType.LOG_TYPE_ERROR)
                                AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted = True
                            End If
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when calling PlayURI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        'hs.waitsecs(1) ' this is to make sure the playerstate has moved to playing before the timeout procedure begins checking for the "end of file" which is player stopped
                        MyAnnouncementIndex = MyAnnouncementIndex + 1
                        AnnouncementReEntry = False
                        Exit Sub
                    End If

                Else
                    ' this is the old TTS way
                    AnnouncementItem.State_ = AnnouncementState.asSpeaking
                    hs.SpeakProxy(0, TextString, True, AnnouncementItem.host)
                    AnnouncementItem.State_ = AnnouncementState.asLinked
                End If
                Index = Index + 1
                MyAnnouncementIndex = MyAnnouncementIndex + 1
            Next
        End If
        If Not AnnouncementItem.Next_ Is Nothing Then
            ' there is more
            Try
                Dim NextAnnouncementItem As AnnouncementItems
                NextAnnouncementItem = AnnouncementItem.Next_
                If NextAnnouncementItem.LinkGroupName = AnnouncementItem.LinkGroupName Then
                    ' this is the same, so do not unlink
                    NextAnnouncementItem.State_ = AnnouncementState.asLinked ' indicate we are already linked
                    DeleteHeadOfAnnouncementQueue()
                    AnnouncementInProgress = False
                    AnnouncementReEntry = False
                    Exit Sub
                End If
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue looking at next announcement in queue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        AnnouncementItem.State_ = AnnouncementState.asUnlinking
        HandleLinkingOff(AnnouncementItem.LinkGroupName)
        AnnouncementItem.State_ = AnnouncementState.asIdle
        DeleteHeadOfAnnouncementQueue()
        If AnnouncementLink Is Nothing Then
            AnnouncementsInQueue = False
            MyAnnouncementIndex = 0
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called and all announcements were processed", LogType.LOG_TYPE_INFO)
        End If
        AnnouncementInProgress = False
        AnnouncementReEntry = False
    End Sub

    Public Function GetLinkgroupSourceZone(ByVal LinkgroupName As String) As Object
        GetLinkgroupSourceZone = Nothing
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLinkgroupSourceZone called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim LinkgroupZoneSource As String
        LinkgroupZoneSource = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
        If LinkgroupZoneSource = "" Then
            ' This could be a zone Name
            Try
                GetLinkgroupSourceZone = GetAPIByUDN(LinkgroupName)
                Exit Function
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLinkgroupSourceZone didn't find " & LinkgroupName & " under [LinkgroupZoneSource] in the .ini file", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            Dim LinkgroupZoneSourceDetails() As String
            LinkgroupZoneSourceDetails = Split(LinkgroupZoneSource, ";")
            LinkgroupZoneSource = LinkgroupZoneSourceDetails(0)
            Try
                GetLinkgroupSourceZone = GetAPIByUDN(LinkgroupZoneSource)
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLinkgroupSourceZone didn't find MusicAPI for " & LinkgroupZoneSource, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Function

    Public Function RemoveBlanks(ByVal InString) As String
        RemoveBlanks = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.length
                If InString(InIndex) = " " Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "!" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "#" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "$" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "&" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "(" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ")" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "*" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "+" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "," Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "-" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "." Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "/" Then
                    Outstring = Outstring + "_"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
            Loop
        Catch ex As Exception
            Log("Error in RemoveBlanks. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        RemoveBlanks = Outstring
    End Function

    Public Sub HandleLinkingOn(ByVal LinkgroupName As String, Optional ByVal IsFile As Boolean = False)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim MusicAPI As HSPI = GetLinkgroupSourceZone(LinkgroupName)
        Try
            If MusicAPI IsNot Nothing Then MusicAPI.SaveQueue()
        Catch ex As Exception
            Log("Error in HandleLinkingOn with LinkgroupName = " & LinkgroupName & " and IsFile = " & IsFile.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub HandleLinkingOff(ByVal LinkgroupName As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOff called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim MusicAPI As HSPI = GetLinkgroupSourceZone(LinkgroupName)
        Try
            If MusicAPI IsNot Nothing Then MusicAPI.RestoreQueue()
        Catch ex As Exception
            Log("Error in HandleLinkingOff with LinkgroupName = " & LinkgroupName & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function CheckForVoiceTag(ByRef inText As String) As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForVoiceTag called with inText = " & inText, LogType.LOG_TYPE_INFO)
        CheckForVoiceTag = ""
        ' Structure = <voice required='Name=Microsoft Anna'>Hello World how are things around here said Anna</voice>
        If inText.IndexOf("<voice ") = -1 Then
            Exit Function
        End If
        Try
            ' OK there appear to be something that looks like a tag
            Dim StartIndexVoiceTag As Integer = inText.IndexOf("<voice ")
            Dim EndIndexVoiceTag As Integer = inText.IndexOf(">", StartIndexVoiceTag)
            If EndIndexVoiceTag = -1 Then Exit Function ' shouldn't be!
            Dim StartIndexCloseVoiceTag As Integer = inText.IndexOf("</voice>", EndIndexVoiceTag)
            If StartIndexCloseVoiceTag = -1 Then Exit Function ' shouldn't be!
            Dim VoiceTagInfo As String = Trim(inText.Substring(StartIndexVoiceTag + 7, EndIndexVoiceTag - StartIndexVoiceTag - 7))
            inText = inText.Remove(StartIndexCloseVoiceTag, 8) ' remove the </voice> tag first
            inText = inText.Remove(StartIndexVoiceTag, EndIndexVoiceTag - StartIndexVoiceTag + 1)
            ' now the VoiceTagInfo should look something like this required='Name=Microsoft Anna'
            If VoiceTagInfo.IndexOf("optional") = 0 Then
                ' not sure what to do with this, won't make any difference, but could decide to simply ignore
                'Exit Function '??
                VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 8))
            ElseIf VoiceTagInfo.IndexOf("required") = 0 Then
                VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 8))
            End If
            ' now the VoiceTagInfo should look something like this ='Name=Microsoft Anna'
            If VoiceTagInfo.IndexOf("=") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))
            If VoiceTagInfo.IndexOf("'") <> 0 Then Exit Function ' should not be
            If VoiceTagInfo.LastIndexOf("'") <> VoiceTagInfo.Length - 1 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(VoiceTagInfo.Length - 1, 1)) ' remove ending ' char
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))  ' remove starting ' char
            ' now the VoiceTagInfo should look something like this Name=Microsoft Anna
            If VoiceTagInfo.IndexOf("Name") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 4))  ' remove starting name char
            If VoiceTagInfo.IndexOf("=") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))  ' remove the = char, now what is left is the voice required
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForVoiceTag returns with inText = " & inText & " and Voice = " & VoiceTagInfo, LogType.LOG_TYPE_INFO)
            Return VoiceTagInfo
        Catch ex As Exception
            Log("Error in CheckForVoiceTag called with inText = " & inText & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


#End Region


End Class

<Serializable()>
Public Structure DBRecord
    Public Id As String
    Public Title As String
    Public ParentID As String
    Public ItemOrContainer As String
    Public ClassType As String
    Public IconURL As String
    Public AlbumName As String
    Public ArtistName As String
    Public Genre As String
End Structure


<Serializable()>
Public Class UPnPController_Config

    Public ghspi As HSPI
    Public page_title As String = ""            ' The custom web page title
    Public linktext As String = ""              ' The custom web page link text
    Public link As String = ""                  ' The custom web page link

    Sub New()
        MyBase.New()
    End Sub


    Public Function GenPage(ByVal lnk As String) As String

        ' GenPage is called on an HTTP Get request for your web page, the one registered with RegisterLinkEx or RegisterConfigLink.
        ' If parameters are provided with the link, such as when they are added to the link, they are provided in the lnk variable.
        ' Example: Your page is registered under the link my_page, the user types: 
        '       http://HomeSeerMachine/my_page?item1=value1&item2=value2
        ' When this procedure is called, it will pass into lnk the data:  /my_page?item1=value1&item2=value2
        ' You can process these pairs yourself manually, or you can call:         GetFormData(data, Me.lPairs, Me.tPair) and it will
        '   process the data and put it into name/value pairs.  See the other web page (WebLink1) for an example.

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenPage called with Link = " & lnk, LogType.LOG_TYPE_INFO)
        Dim sb As New StringBuilder()
        If lnk = "/" & sIFACE_NAME & "_config" Then
            ' this is to configure the plugin itself
            Try
                '
                ' Generate the HTML re-direct
                '
                sb.Append("HTTP/1.1 303 See Other" & vbCrLf)
                sb.Append("Location: " & sIFACE_NAME & "/" & sIFACE_NAME & "Config.aspx?plugin=" & sIFACE_NAME & "&State=Init" & vbCrLf)

                Return sb.ToString
            Catch pEx As Exception
                Return ""
            End Try
            Exit Function
        End If
        Dim LinkParts()
        LinkParts = Split(lnk, ";")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenPage called with Link Part = " & LinkParts(0) & " " & LinkParts(1), LogType.LOG_TYPE_INFO)
        'GenPage = BuildPage()
        Try
            '
            ' Generate the HTML re-direct
            '
            sb.Append("HTTP/1.1 303 See Other" & vbCrLf)
            sb.Append("Location: " & sIFACE_NAME & "/" & sIFACE_NAME & "DevicePage.aspx" & "?plugin=" & sIFACE_NAME & "&instance=" & LinkParts(1) & vbCrLf)

            Return sb.ToString
        Catch pEx As Exception
            Return ""
        End Try
    End Function

End Class


<Serializable()> _
Public Class MyUPnpDeviceInfo
    Public UPnPDeviceFriendlyName As String
    Public UPnPDeviceGivenName As String
    Public UPnPDeviceModelNumber As String
    Public UPnPDeviceUDN As String
    Public UPnPDeviceHSRef As Integer
    Public UPnPDeviceDeviceType As String
    Public UPnPDeviceOnLine As Boolean
    Public UPnPDeviceMusicAPIIndex As Integer
    Public UPnPDeviceControllerRef As HSPI
    Public UPnPDeviceWeblinkCreated As Boolean
    Public Device As MyUPnPDevice
    Public UPnPDeviceisRoot As Boolean = False
    Public UPnPDeviceDescription As String = False
    Public UPnPDeviceModelURL As String = False
    Public UPnPDevicePresentationURL As String = False
    Public UPnPDeviceHasChildren As Boolean = False
    Public UPnPDeviceManufacturerURL As String = ""
    Public UPnPDeviceModelName As String = ""
    Public UPnPDeviceUPC As String = ""
    Public UPnPDeviceIconURL As String = ""
    Public UPnPDeviceManufacturerName As String = ""
    Public UPnPDeviceIPAddress As String = ""
    Public UPnPDeviceIPPort As String = ""
    Public UPnPDeviceAdminStateActive As Boolean = False
    Public UPnPDeviceIsAddedToHS As Boolean = False
    Public UPnPDeviceServiceTypes As String = ""
    Public NewDevice As Boolean = False


    Public Sub New()
        MyBase.New()
        UPnPDeviceFriendlyName = ""
        UPnPDeviceGivenName = ""
        UPnPDeviceModelNumber = ""
        UPnPDeviceUDN = ""
        UPnPDeviceHSRef = -1
        UPnPDeviceDeviceType = ""
        UPnPDeviceOnLine = False
        UPnPDeviceMusicAPIIndex = 0
        UPnPDeviceControllerRef = Nothing
        UPnPDeviceWeblinkCreated = False
        Device = Nothing
        UPnPDeviceisRoot = False
        UPnPDeviceDescription = False
        UPnPDeviceModelURL = False
        UPnPDevicePresentationURL = False
        UPnPDeviceHasChildren = False
        UPnPDeviceManufacturerURL = ""
        UPnPDeviceModelName = ""
        UPnPDeviceUPC = ""
        UPnPDeviceIconURL = ""
        UPnPDeviceManufacturerName = ""
        UPnPDeviceIPAddress = ""
        UPnPDeviceIPPort = ""
        UPnPDeviceAdminStateActive = False
        UPnPDeviceIsAddedToHS = False
        NewDevice = False
        UPnPDeviceServiceTypes = ""
    End Sub

    Public Sub Close()
        UPnPDeviceFriendlyName = ""
        UPnPDeviceGivenName = ""
        UPnPDeviceModelNumber = ""
        UPnPDeviceUDN = ""
        UPnPDeviceHSRef = -1
        UPnPDeviceDeviceType = ""
        UPnPDeviceOnLine = False
        UPnPDeviceMusicAPIIndex = 0
        UPnPDeviceControllerRef = Nothing
        UPnPDeviceWeblinkCreated = False
        Device = Nothing
        UPnPDeviceisRoot = False
        UPnPDeviceDescription = False
        UPnPDeviceModelURL = False
        UPnPDevicePresentationURL = False
        UPnPDeviceHasChildren = False
        UPnPDeviceManufacturerURL = ""
        UPnPDeviceModelName = ""
        UPnPDeviceUPC = ""
        UPnPDeviceIconURL = ""
        UPnPDeviceManufacturerName = ""
        UPnPDeviceIPAddress = ""
        UPnPDeviceIPPort = ""
        NewDevice = False
        UPnPDeviceServiceTypes = ""
    End Sub
End Class


Module HS_GLOBAL_VARIABLES

    ' interface status
    ' for InterfaceStatus function call
    Public Const ERR_NONE = 0
    Public Const ERR_SEND = 1
    Public Const ERR_INIT = 2

    ' Master State Pairs
    Public Const msDisconnected = 100
    Public Const msConnected = 101
    Public Const msInitializing = 102
    Public Const msDisconnect = 1
    Public Const msConnect = 2
    Public Const msDoRediscovery = 3

    ' Device State Pairs
    Public Const dsDeactivate = 1
    Public Const dsActivate = 2

    Public Const dsWOL = 5
    Public Const dsBuildDB = 6

    Public Const dsDeactivated = 200
    Public Const dsActivatedOffLine = 201
    Public Const dsActivateOnLine = 202
    Public Const dsActivatedOnLineUnregistered = 203
    Public Const dsBuildingDB = 204

    ' Party device states
    Public Const dsIdle = 200
    Public Const dsSinger = 201
    Public Const dsListener = 202

    ' Player State Pairs
    Public Const psStopped = 2
    Public Const psPaused = 3
    Public Const psPlaying = 1
    Public Const psTransitioning = 6
    Public Const psRewind = 5
    Public Const psFastForward = 4
    Public Const psPlay = 1
    Public Const psStop = 2
    Public Const psPause = 3
    Public Const psNext = 4
    Public Const psPrevious = 5
    Public Const psBuildiPodDB = 6
    Public Const psShuffle = 7
    Public Const psRepeat = 8
    Public Const psVolUp = 9
    Public Const psVolDown = 10
    Public Const psMute = 11
    Public Const psBalanceLeft = 12
    Public Const psBalanceRight = 13
    Public Const psLoudness = 14
    Public Const psVolSlider = 15
    Public Const psBalanceSlider = 16
    Public Const psClearQueue = 17
    Public Const psFF = 18
    Public Const psREW = 19
    Public Const psQueueShuffle = 20
    Public Const psQueueRepeat = 21

    ' remote 
    Public Const psRemoteOn = 0
    Public Const psRemoteOff = 1
    Public Const psPartyAll = 2
    Public Const psExitParty = 3
    Public Const psRegister = 4
    Public Const psUnregister = 5
    Public Const psWOL = 6
    Public Const psCreateRemoteButtons = 7
    Public Const psCreateRemoteAppButtons = 8

    Public Const psStartPartyDevices = 100


    ' MuteState Pairs
    Public Const msMuted = 1001
    Public Const msUnmuted = 1000

    ' ShuffleState Pairs
    Public Const ssShuffled = 1001
    Public Const ssNoShuffle = 1000

    ' RepeatState Pairs
    Public Const rsRepeat = 1001
    Public Const rsnoRepeat = 1000

    ' Down - Slider - Up pairs
    Public Const vpDown = 1000
    Public Const vpSlider = 1
    Public Const vpUp = 1001
    Public Const vpMidPoint = 30
    Public Const vpLGDown = 1010
    Public Const vpLGUp = 1011

    ' Toggle pairs
    Public Const tpOff = 1000
    Public Const tpOn = 1001
    Public Const tpToggle = 1002
    Public Const tpLGOff = 1010
    Public Const tpLGOn = 1011
    Public Const tpLGToggle = 1012

    Public Const lsLoudnessOff = 1000
    Public Const lsLoudnessOn = 1001

    ' Generic Pairs
    Public Const gsDefault = 1000


    Public hs As HomeSeerAPI.IHSApplication
    Public callback As HomeSeerAPI.IAppCallbackAPI
    Public InterfaceVersion As Integer
    Public bShutDown As Boolean = False
    Public MyShutDownRequest As Boolean = False

    'Public instance As String = ""                             ' set when SupportMultipleInstances is TRUE
    Public PIDebuglevel As DebugLevel = DebugLevel.dlErrorsOnly
    Public gLogToDisk As Boolean = False
    Public gHSInitialized As Boolean = False
    Public ImRunningOnLinux As Boolean = False
    Public HSisRunningOnLinux As Boolean = False
    Public ImRunningLocal As Boolean = True
    Public UPnPDebuglevel As DebugLevel = DebugLevel.dlErrorsOnly
    Public PlugInIPAddress As String = ""
    Public PluginIPPort As String = ""

    Public bHSInitialized As Boolean = False                ' Indicates HS object was successfully initialized
    Public gInterfaceStatus As Integer = ERR_INIT           ' Interface status
    Public MasterHSDeviceRef As Integer = -1

    Public gIOEnabled As Boolean = False                    ' IO interface enabled
    Public MyPIisInitialized As Boolean = False

    Public Const IFACE_NAME As String = "Media Controller"   ' This is the plugin's name
    Public Const sIFACE_NAME As String = "MediaController"   ' This is the system wide plugin's name
    Public tIFACE_NAME As String = "MediaController"
    Public ShortIfaceName As String = "MC"
    Public MyINIFile As String = tIFACE_NAME & ".ini"    ' Configuration File' 

    ' File Path Name definition. They may get overwritten in InitIO, based on whether we run on Linux or not
    Public CurrentAppPath As String = ""
    Public DBPath As String = "\html\" & sIFACE_NAME & "\Databases\"
    Public SearchResultFile As String = "\html\" & sIFACE_NAME & "\SearchResults\UPnPSearchResult"
    Public gPlaylistPath As String = "\html\" & sIFACE_NAME & "\Playlists\"
    Public gRemoteControlPath As String = sIFACE_NAME & "_RemoteControl.ini"
    'Public gRemoteImportPath = "\html\" & sIFACE_NAME & "\RemoteControl\Imports\"
    Public DebugLogFileName As String = "\" & tIFACE_NAME & "\Logs\MediaControllerDebug.txt"
    Public BinPath As String = "\html\" & sIFACE_NAME & "\bin\"

    ' Used as URL and FileName
    Public AnnouncementPath As String = "/" & tIFACE_NAME & "/Announcements/"

    ' URLs
    Public ImagesPath As String = "/images/" & sIFACE_NAME & "/"
    Public URLImagesPath As String = "/images/" & sIFACE_NAME & "/"
    Public NoArtPath As String = "/images/" & sIFACE_NAME & "/NoArt.png"
    Public ArtWorkPath As String = "/images/" & tIFACE_NAME & "/Artwork/"
    Public URLArtWorkPath As String = "/images/" & tIFACE_NAME & "/Artwork/"
    Public FileArtWorkPath As String = tIFACE_NAME & "\Artwork\"

    Public Const ConfigPage As String = "MediaControl"
    Public Const PlayerControlPage As String = "MediaConfig"
    Public Const PlayerConfig As String = "MediaPlayerConfig"
    Public Const UPnPViewPage As String = "UPnPViewer"

    ' Part of the UPNP stuff
    Public TCPListenerPort As Integer = 0
    Public UPnPSubscribeTimeOut As Integer = 1800
    Public MySSDPDevice As MySSDP
    Public Const LoopBackIPv4Address As String = "127.0.0.1"
    Public Const AnyIPv4Address As String = "0.0.0.0"


    Public TransportInfo, RenderingInfo, ContentInfo As String
    Public Const cMaxNbrOfUPNP = 999
    Public MaxNbrOfUPNPObjects As Integer = cMaxNbrOfUPNP

    Public NbrOfPingRetries As Integer = 3
    Public ShowFailedPings As Boolean = False
    Public ArtworkHSize As Integer = 0
    Public ArtworkVSize As Integer = 0
    Public MaxWaitTimeRetrievingContainer As Integer = 120


    Public AnnouncementLink As AnnouncementItems = Nothing
    Public AnnouncementsInQueue As Boolean = False
    Public AnnouncementInProgress As Boolean = False
    Public AnnouncementReEntry As Boolean = False
    Public MyAnnouncementCountdown As Integer = 100
    Public MyAnnouncementIndex As Integer = 0
    Public LastStoredQueueID As String = ""
    Public AnnouncementTitle As String = "HomeSeer Announcement"
    Public AnnouncementAuthor As String = "Dirk Corsus"
    Public AnnouncementAlbum As String = "Media Controller"
    Public MyMaxAnnouncementTime As Integer = 100
    Public MyVolumeStep As Integer = 5
    Public MyNoPingingFlag As Boolean = False

    Public MyHSTrackLengthFormat As HSSTrackPositionSettings = HSSTrackPositionSettings.TPSSeconds
    Public MyHSTrackPositionFormat As HSSTrackPositionSettings = HSSTrackPositionSettings.TPSSeconds

    Public PreviousVersion As Integer = 0
    Public Const CurrentVersion As Integer = 0

    Public Const OpcodeText = 1
    Public Const OpcodeBinary = 2
    Public Const OpcodeClose = 8
    Public Const OpcodePing = 9
    Public Const OpcodePong = 10

    Public Enum player_state_values
        playing = 1
        stopped = 2
        paused = 3
        forwarding = 4
        rewinding = 5
        Transitioning = 6
        UpdateHSServerOnly = 17
    End Enum

    Public Enum DebugLevel
        dlOff = 0
        dlErrorsOnly = 1
        dlEvents = 2
        dlVerbose = 3
    End Enum

    Public Enum repeat_modes
        repeat_off = 0
        repeat_one = 1
        repeat_all = 2
    End Enum

    Public Enum player_selections
        Playlist_Track = 0
        Artist_Album_Track = 1
        Artist_Track = 2
        Album_Track = 3
        Playlist = 4
        Artist_Album = 5
        Album = 6
        Artist = 7
        Genre = 8
        audiobook = 9
        podcast = 10
        LineInput = 11
    End Enum

    Public Enum AnnouncementState
        asIdle = 0
        asLinking = 1
        asLinked = 2
        asSpeaking = 3
        asFilePlayed = 4
        asUnlinking = 5
    End Enum

    Public Enum ManagePingActions
        mpAdd
        mpRemove
        mpTimeOut
        mpStatus
    End Enum

    Public Enum HSDevices
        Player = 0
        Status = 1
        Control = 2
        Repeat = 3
        Shuffle = 4
        Volume = 5
        Balance = 6
        Mute = 7
        Loudness = 8
        Tittle = 9
        NexTittle = 10
        Artist = 11
        NextArtist = 12
        Album = 13
        NextAlbum = 14
        Art = 15
        NextArt = 16
        TrackLength = 17
        TrackPosition = 18
        RadioStationName = 19
        TrackDescriptor = 20
        Genre = 21
        TrackDate = 22
        Master = 23
        RootDevice = 24
        Message = 25
        Remote = 26
        Party = 27
        Speed = 28
        Server = 29
    End Enum

    Public Enum DeviceInfoIndex
        diGivenName = 0
        diUDN = 1
        diDeviceModelName = 2
        diDeviceType = 3
        diFriendlyName = 4
        diIPAddress = 5
        diIPPort = 6
        diMusicAPIIndex = 7
        diAdminState = 8
        diDeviceIsAdded = 9
        diHSDeviceRef = 10
        diDeviceAPIIndex = 11
        diTimeBetweenPictures = 12
        diArtistObjectID = 13
        diAlbumObjectID = 14
        diTrackObjectID = 15
        diGenreObjectID = 16
        diPlayListObjectID = 17
        diPollTransportChanges = 18
        diQueueRepeat = 19
        diQueueShuffle = 20
        diServerUDN = 21
        diSystemUpdateID = 22
        diPictureSize = 23
        diPollVolumeChanges = 24
        diRemoteControl = 25
        diRegistered = 26
        diRemoteType = 27
        diNextAV = 28
        diUseNextAV = 29
        diAnnouncementMP3 = 30
        diMusicObjectID = 31
        diPhotosObjectID = 32
        diVideosObjectID = 33
        diSystemUpdateIDAtDBCreation = 34
        diMACAddress = 35
        ' these are for HS3
        diTrackHSRef = 36
        diNextTrackHSRef = 37
        diArtistHSRef = 38
        diNextArtistHSRef = 39
        diAlbumHSRef = 40
        diNextAlbumHSRef = 41
        diArtHSRef = 42
        diNextArtHSRef = 43
        diPlayStateHSRef = 44
        diVolumeHSRef = 45
        diMuteHSRef = 46
        diLoudnessHSRef = 47
        diBalanceHSRef = 48
        diTrackLengthHSRef = 49
        diTrackPosHSRef = 50
        diRadiostationNameHSRef = 51
        diTrackDescrHSRef = 52
        diRepeatHSRef = 53
        diShuffleHSRef = 54
        diPlayerHSRef = 55
        diGenreHSRef = 56
        ' These specific for Sonos
        diSonosPlayerType = 57
        diSonosReplicationInfo = 58
        diDeviceIConURL = 59
        diRoomIcon = 60
        diSpeedHSRef = 61
        diSonyPartyHSRef = 62
        diAdminStateRemote = 63
        diAdminStateMessageService = 65
        diDeviceModelNumber = 66
        diDeviceManufacturerName = 67
        diQueueRepeatHSRef = 68
        diQueueShuffleHSRef = 69
        diArtFileIndex = 70
        diNextArtFileIndex = 71
        diSonyAuthenticationPIN = 72
        diSonyRemoteRegisterType = 73
        diSonyCookieExpiryDate = 74
        diDeviceServiceTypes = 75
        diOnkyoPortNbr = 76
        diSamsungWebSocketPort = 77
        diSamsungWebSocketName = 78
        diSamsungWebSocketLocation = 79
        diSamsungWebSocketProductCap = 80
        diSamsungDeviceID = 81
        diSamsungProductCap = 82
        diSamsungSessionID = 83
        diSamsungAesKey = 84
        diSamsungDeviceCapabilities = 85
        diSamsungRemotePIN = 86
        diSecWebSocketKey = 87
        diSamsungPairingPort = 88
        diSamsungAppId = 89
        diSamsungServerAckMsg = 90
        diLGClientKey = 91
        diSamsungToken = 92
        diWifiMacAddress = 93
        diSamsungTokenAuthSupport = 94
        diSamsungisSupportInfo = 95
        diSamsungClientID = 96
    End Enum


    Public Enum PictureSize
        psDefault = 0
        psTiny = 1
        psSmall = 2
        psMedium = 3
        psLarge = 4
    End Enum

    <Serializable()> Public Structure IPAddressInfo
        Public IPAddress As String
        Public IPPort As String
    End Structure

    <Serializable()> Public Class PingArrayElement
        Public IPAddress As String
        Public Status As String
        Public FailedPingCount As Integer
        Public ClientNameList As String()
    End Class


    <Serializable()> _
    Public Class AnnouncementItems
        Public Sub New()
            MyBase.New()
        End Sub
        Public LinkGroupName As String = ""
        Public device As Short = 0
        Public text As String = ""
        Public wait As Boolean = True
        Public host As String = ""
        Public IsFile As Boolean = False
        Public State_ As AnnouncementState = AnnouncementState.asIdle
        Public Next_ As Object = Nothing
        Public Previous_ As Object = Nothing
        Public SourceZoneMusicAPI As Object = Nothing
    End Class

    Public Enum PostAnnouncementAction
        paaAlwaysForward = 0
        paaForwardNoMatch = 1
        paaAlwaysDrop = 2
    End Enum

    Public Enum HSSTrackLengthSettings
        TLSSeconds = 0
        TLSHoursMinutesSeconds = 1
    End Enum

    Public Enum HSSTrackPositionSettings
        TPSSeconds = 0
        TPSHoursMinutesSeconds = 1
        TPSPercentage = 2
    End Enum

    <Serializable>
    Class NetworkInfo
        Public description As String
        Public mac As String
        Public operationalstate As String
        Public id As String
        Public name As String
        Public interfacetype As String
        Public addressinfo As List(Of NetworkInfoAddrMask)
    End Class

    <Serializable>
    Class NetworkInfoAddrMask
        Public address As String
        Public mask As String
        Public addrtype As String
    End Class
End Module