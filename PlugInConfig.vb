Imports Scheduler
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Text
Imports System.Web
Imports System.Web.UI.HtmlControls

Public Class PlugInConfig


    Inherits clsPageBuilder

    Private PIReference As HSPI = Nothing
    Private MyPageName As String = ""


    Private stb As New StringBuilder

    Private PIDebugLevelDropList As clsJQuery.jqDropList
    Private UPnPDebugLevelDropList As clsJQuery.jqDropList
    Private LogErrorOnlyChkBox As clsJQuery.jqCheckBox
    Private LogToDiskChkBox As clsJQuery.jqCheckBox

    Private VolumeStepBox As clsJQuery.jqTextBox
    Private SpeakerProxyCheckBox As clsJQuery.jqCheckBox
    Private UPNPNbrOfElementsBox As clsJQuery.jqTextBox
    Private HSizeBox As clsJQuery.jqTextBox
    Private VSizeBox As clsJQuery.jqTextBox
    Private AddEntryLinkTableBtn As clsJQuery.jqButton
    Private AddSlideShowDeviceBtn As clsJQuery.jqButton
    Private AddMediaDeviceBtn As clsJQuery.jqButton
    Private LogNetworkBtn As clsJQuery.jqButton
    Private PostAnnouncementActionBox As clsJQuery.jqDropList
    Private HSSTrackLengthBox As clsJQuery.jqDropList
    Private HSSTrackPositionBox As clsJQuery.jqDropList
    Private PINBox As clsJQuery.jqTextBox


    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
        MyPageName = pagename
        PIDebugLevelDropList = New clsJQuery.jqDropList("PIDebugLvlBox", MyPageName, False)
        PIDebugLevelDropList.toolTip = "Set the level of debug logging for the plugin functions"
        UPnPDebugLevelDropList = New clsJQuery.jqDropList("UPnPDebugLvlBox", MyPageName, False)
        UPnPDebugLevelDropList.toolTip = "Set the level of debug logging for the UPnP functions"
        LogErrorOnlyChkBox = New clsJQuery.jqCheckBox("LogErrorOnlyChkBox", " Log Error Only Flag", MyPageName, True, False)
        LogErrorOnlyChkBox.toolTip = "Not implemented"
        LogToDiskChkBox = New clsJQuery.jqCheckBox("LogToDiskChkBox", " Log to Disk Flag", MyPageName, True, False)
        LogToDiskChkBox.toolTip = "Log the plug-in errors to a standard txt file. Note this slows down performance substantial. Suggested use for remote ran PIs or capture issues when terminating the PI"
        PostAnnouncementActionBox = New clsJQuery.jqDropList("PostAnnouncementActionBox", MyPageName, False)
        PostAnnouncementActionBox.toolTip = "Specify what to do with the announcement that was intercepted by the plugin's proxy client"
        VolumeStepBox = New clsJQuery.jqTextBox("VolumeStepBox", "text", "", MyPageName, 5, False)
        VolumeStepBox.toolTip = "Specify the amount of volume you want the player to increase/decrease when you click on the Volume up/down buttons"
        SpeakerProxyCheckBox = New clsJQuery.jqCheckBox("SpeakerProxyCheckBox", " Proxy Flag", MyPageName, True, False)
        SpeakerProxyCheckBox.toolTip = "Set this if you want the plug-in to participate in announcements"
        UPNPNbrOfElementsBox = New clsJQuery.jqTextBox("UPNPNbrOfElementsBox", "text", "", MyPageName, 5, False)
        UPNPNbrOfElementsBox.toolTip = "Set the max nbr of UPnP objects that can be retrieved in one read. For XP set to 50 for other versions of Windows less then 999"
        HSizeBox = New clsJQuery.jqTextBox("HSizeBox", "text", "", MyPageName, 5, False)
        HSizeBox.toolTip = "Set the size of the album artwork (in pixels) on the HS Device Management page"
        VSizeBox = New clsJQuery.jqTextBox("VSizeBox", "text", "", MyPageName, 5, False)
        VSizeBox.toolTip = "Set the size of the album artwork (in pixels) on the HS Device Management page"
        AddEntryLinkTableBtn = New clsJQuery.jqButton("AddEntryLinkTableBtn", "Add Entry", MyPageName, False)
        AddEntryLinkTableBtn.toolTip = "Create a new entry in the Linktable"
        AddSlideShowDeviceBtn = New clsJQuery.jqButton("AddSlideShowDeviceBtn", "Add Pictureshow Device", MyPageName, False)
        AddSlideShowDeviceBtn.toolTip = "Add a new Pictureshow device"
        AddMediaDeviceBtn = New clsJQuery.jqButton("AddMediaDeviceBtn", "Add Media Device", MyPageName, False)
        AddMediaDeviceBtn.toolTip = "Add a new Media device"
        LogNetworkBtn = New clsJQuery.jqButton("LogNetworkBtn", "Log Whole Network", MyPageName, False)
        LogNetworkBtn.toolTip = "Find all active devices and log all UPnP info in the SearchResults sub-directory"
        LogNetworkBtn.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Scanning & Logging UPnP Devices ...</h2>' });"
        HSSTrackLengthBox = New clsJQuery.jqDropList("HSSTrackLengthBox", MyPageName, False)
        HSSTrackLengthBox.toolTip = "Set the format on how you want the info to be displayed as part of the HS Device"
        HSSTrackPositionBox = New clsJQuery.jqDropList("HSSTrackPositionBox", MyPageName, False)
        HSSTrackPositionBox.toolTip = "Set the format on how you want the info to be displayed as part of the HS Device"


    End Sub

    Public WriteOnly Property RefToPlugIn As HSPI
        Set(value As HSPI)
            PIReference = value
        End Set
    End Property


    ' build and return the actual page
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for PlugInConfig called with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)

        Dim stb As New StringBuilder
        Dim stbPlayerTable As New StringBuilder

        Try
            Me.reset()

            ' handle any queries like mode=something
            Dim parts As Collections.Specialized.NameValueCollection = Nothing
            If (queryString <> "") Then
                parts = HttpUtility.ParseQueryString(queryString)
            End If

            Me.AddHeader(hs.GetPageHeader(pageName, "Media Controller Configuration", "", "", False, True))
            stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd)

            ' specific page starts here

            Dim PIdeblevel As DebugLevel = GetIntegerIniFile("Options", "PIDebugLevel", DebugLevel.dlOff)
            PIDebugLevelDropList.ClearItems()
            PIDebugLevelDropList.AddItem("Off", DebugLevel.dlOff, PIdeblevel = DebugLevel.dlOff)
            PIDebugLevelDropList.AddItem("Errors Only", DebugLevel.dlErrorsOnly, PIdeblevel = DebugLevel.dlErrorsOnly)
            PIDebugLevelDropList.AddItem("Events and Errors", DebugLevel.dlEvents, PIdeblevel = DebugLevel.dlEvents)
            PIDebugLevelDropList.AddItem("Verbose", DebugLevel.dlVerbose, PIdeblevel = DebugLevel.dlVerbose)
            LogToDiskChkBox.checked = GetBooleanIniFile("Options", "LogToDisk", False)
            Dim UPNPdeblevel As DebugLevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlOff)
            UPnPDebugLevelDropList.ClearItems()
            UPnPDebugLevelDropList.AddItem("Off", DebugLevel.dlOff, UPNPdeblevel = DebugLevel.dlOff)
            UPnPDebugLevelDropList.AddItem("Errors Only", DebugLevel.dlErrorsOnly, UPNPdeblevel = DebugLevel.dlErrorsOnly)
            UPnPDebugLevelDropList.AddItem("Events and Errors", DebugLevel.dlEvents, UPNPdeblevel = DebugLevel.dlEvents)
            UPnPDebugLevelDropList.AddItem("Verbose", DebugLevel.dlVerbose, UPNPdeblevel = DebugLevel.dlVerbose)

            VolumeStepBox.defaultText = GetStringIniFile("Options", "VolumeStep", "")
            SpeakerProxyCheckBox.checked = GetBooleanIniFile("SpeakerProxy", "Active", False)
            UPNPNbrOfElementsBox.defaultText = GetStringIniFile("Options", "MaxNbrofUPNPObjects", "")
            HSizeBox.defaultText = GetStringIniFile("Options", "ArtworkHSize", "")
            VSizeBox.defaultText = GetStringIniFile("Options", "ArtworkVsize", "")

            stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
            stb.Append("<h1>Media Manager Plugin Configuration</h1>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)

            stb.Append("<table style='width: 50%' >")

            stb.Append("<tr><td colspan='2'>")
            stb.Append("<hr /> ")
            stb.Append("</tr></td><tr><td>")
            stb.Append(clsPageBuilder.DivStart("DebugPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Debug Information</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)


            stb.Append(PIDebugLevelDropList.Build & " Plugin Functions Debug Level")
            stb.Append("</br>")
            stb.Append(UPnPDebugLevelDropList.Build & " UPnP Functions Debug Level")
            stb.Append("</br>")
            stb.Append(LogToDiskChkBox.Build)
            stb.Append("<hr /> ")

            stb.Append(clsPageBuilder.DivStart("UPnPSettingPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>UPnP Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(UPNPNbrOfElementsBox.Build)
            stb.Append("# of UPnP Elements<hr />")

            stb.Append(clsPageBuilder.DivStart("VolumeSettingPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Volume Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(VolumeStepBox.Build)
            stb.Append("Volume Step")

            stb.Append(clsPageBuilder.DivStart("SpeakerProxyPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Speaker Proxy Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(SpeakerProxyCheckBox.Build)
            stb.Append("</br>")
            Dim paaSetting = GetIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
            PostAnnouncementActionBox.ClearItems()
            PostAnnouncementActionBox.AddItem("Always Forward", PostAnnouncementAction.paaAlwaysForward, paaSetting = PostAnnouncementAction.paaAlwaysForward)
            PostAnnouncementActionBox.AddItem("Forward When No Match", PostAnnouncementAction.paaForwardNoMatch, paaSetting = PostAnnouncementAction.paaForwardNoMatch)
            PostAnnouncementActionBox.AddItem("Never Forward", PostAnnouncementAction.paaAlwaysDrop, paaSetting = PostAnnouncementAction.paaAlwaysDrop)
            stb.Append(PostAnnouncementActionBox.Build)
            stb.Append(" Post Announcement Action")
            stb.Append("</br></br>")
            stb.Append("</td><td>")

            stb.Append(clsPageBuilder.DivStart("ArtWorkPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Artwork Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(HSizeBox.Build)
            stb.Append("Artwort Width</br>")
            stb.Append(VSizeBox.Build)
            stb.Append("Artwork Height<hr />")

            stb.Append(clsPageBuilder.DivStart("HSDeviceSettingsPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>HS3 Devices Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            Dim HSSTrackLengthBoxSetting = GetIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSSeconds)
            HSSTrackLengthBox.ClearItems()
            HSSTrackLengthBox.AddItem("Seconds", HSSTrackLengthSettings.TLSSeconds, HSSTrackLengthBoxSetting = HSSTrackLengthSettings.TLSSeconds)
            HSSTrackLengthBox.AddItem("HH:MM:SS", HSSTrackLengthSettings.TLSHoursMinutesSeconds, HSSTrackLengthBoxSetting = HSSTrackLengthSettings.TLSHoursMinutesSeconds)
            stb.Append(HSSTrackLengthBox.Build & " Tracklength Format")
            stb.Append("</br>")
            Dim HSSTrackPositionBoxSetting = GetIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSSeconds)
            HSSTrackPositionBox.ClearItems()
            HSSTrackPositionBox.AddItem("Seconds", HSSTrackPositionSettings.TPSSeconds, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSSeconds)
            HSSTrackPositionBox.AddItem("HH:MM:SS", HSSTrackPositionSettings.TPSHoursMinutesSeconds, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSHoursMinutesSeconds)
            HSSTrackPositionBox.AddItem("Percentage", HSSTrackPositionSettings.TPSPercentage, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSPercentage)
            stb.Append(HSSTrackPositionBox.Build & " Track Position Format")

            stb.Append("</br></br></br>") ' this is to line them up on the page
            stb.Append("</td></tr></table>")



            stb.Append("<table><tr><td>")

            ' create the player table

            stbPlayerTable.Append("<table ID='PlayerListTable' border='1'  style='background-color:DarkGray;color:black'>")
            stbPlayerTable.Append("<tr ID='HeaderRow'  style='background-color:DarkGray;color:black'>")
            stbPlayerTable.Append("<td><h3> Status </h3></td><td><h3> Device Given Name </h3></td><td><h3> Device Own Name </h3></td><td><h3> IP Address </h3></td><td><h3> Unique Device Name </h3></td><td><h3> Service Type </h3></td><td><h3> Player Alive </h3></td><td><h3> Configure/Add </h3></td><td><h3> Remove/Delete </h3></td></tr>")

            Dim PlayerListTableRow As Integer = 0

            Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
            DLNADevices = GetIniSection("UPnP Devices UDN to Info")

            Try
                If Not DLNADevices Is Nothing And DLNADevices.Count > 0 Then
                    Dim DLNADevice As System.Collections.Generic.KeyValuePair(Of String, String)
                    Dim ConfigBtnIndex As Integer = 1
                    For Each DLNADevice In DLNADevices
                        If DLNADevice.Key <> "" Then
                            Dim DeviceGivenName As String = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diGivenName.ToString, "")
                            Dim DeviceUDN As String = DLNADevice.Key
                            Dim DeviceType As String = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceType.ToString, "")
                            Select Case DeviceType
                                Case "DMR"
                                    DeviceType = "Player"
                                Case "DMS"
                                    DeviceType = "Server"
                                Case "RCR"
                                    DeviceType = "Remote"
                                Case "HST"
                                    DeviceType = "Pictureshow"
                                Case "PMR"
                                    DeviceType = "Message"
                                Case "DIAL"
                                    DeviceType = "MultiScreen"
                            End Select
                            Dim DeviceServiceTypes As String = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceServiceTypes.ToString, "")
                            Dim Services As String() = Split(DeviceServiceTypes, ",")
                            Dim outServiceTypes As String = ""
                            For Each ServiceType As String In Services
                                Select Case ServiceType
                                    Case "DMR"
                                        If outServiceTypes = "" Then outServiceTypes &= "Player" Else outServiceTypes &= ", " & "Player"
                                    Case "DMS"
                                        If outServiceTypes = "" Then outServiceTypes &= "Server" Else outServiceTypes &= ", " & "Server"
                                    Case "RCR"
                                        If outServiceTypes = "" Then outServiceTypes &= "Remote" Else outServiceTypes &= ", " & "Remote"
                                    Case "HST"
                                        If outServiceTypes = "" Then outServiceTypes &= "Pictureshow" Else outServiceTypes &= ", " & "Pictureshow"
                                    Case "PMR"
                                        If outServiceTypes = "" Then outServiceTypes &= "Message" Else outServiceTypes &= ", " & "Message"
                                    Case "DIAL"
                                        If outServiceTypes = "" Then outServiceTypes &= "MultiScreen" Else outServiceTypes &= ", " & "MultiScreen"
                                End Select
                            Next
                            Dim DeviceOwnName As String = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diFriendlyName.ToString, "")
                            Dim DeviceIpAddress As String = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diIPAddress.ToString, "")
                            Dim DeviceActive As Boolean = GetBooleanIniFile(DLNADevice.Key, DeviceInfoIndex.diAdminState.ToString, False)
                            Dim DeviceIsAddedToHS As Boolean = GetBooleanIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceIsAdded.ToString, False)
                            Dim DeviceAPI As HSPI = Nothing
                            Dim StatusImage As String = "" 'New Image
                            If DeviceActive Then
                                DeviceAPI = PIReference.GetAPIByUDN(DeviceUDN)
                                If Not DeviceAPI Is Nothing Then
                                    Dim DeviceState As Integer = DeviceAPI.DeviceState
                                    If DeviceState = 2 Then
                                        StatusImage = ImagesPath & "OKBtn.png" '"OKBtn-Small.png"
                                    Else
                                        StatusImage = ImagesPath & "PartialOKBtn.png" ' "PartialOKBtn-Small.png"
                                    End If
                                Else
                                    StatusImage = ImagesPath & "NOKBtn.png" ' "NOKBtn-Small.png"
                                End If
                            Else
                                StatusImage = ImagesPath & "NOKBtn.png" ' "NOKBtn-Small.png"
                            End If
                            stbPlayerTable.Append("<tr ID='EntryRow'  style='background-color:LightGray;color:black'>")
                            stbPlayerTable.Append("<td>" & "<img src=" & StatusImage & " height='40px' width='40px' >" & "</td>")
                            Dim DeviceGivenNameEditBox As New clsJQuery.jqTextBox("DeviceGNBox" & "_" & PlayerListTableRow.ToString, "text", "", MyPageName, 20, False)
                            DeviceGivenNameEditBox.defaultText = DeviceGivenName
                            DeviceGivenNameEditBox.toolTip = "Enter here your own preferred name for this device"
                            stbPlayerTable.Append("<td>" & DeviceGivenNameEditBox.Build & "</td>")
                            stbPlayerTable.Append("<td>" & DeviceOwnName & "</td>")
                            stbPlayerTable.Append("<td>" & DeviceIpAddress & "</td>")
                            stbPlayerTable.Append("<td>" & DeviceUDN & "</td>")
                            stbPlayerTable.Append("<td>" & outServiceTypes & "</td>")
                            Dim Alive As String = "?"
                            If MySSDPDevice IsNot Nothing Then
                                Dim DevicesList As MyUPnPDevices = MySSDPDevice.GetAllDevices()
                                If DevicesList IsNot Nothing Then
                                    Dim Device As MyUPnPDevice = DevicesList.Item("uuid:" & DeviceUDN, True)
                                    If Device IsNot Nothing Then
                                        If Device.Alive Then
                                            Alive = "True"
                                        Else
                                            Alive = "False"
                                        End If
                                    End If
                                End If
                            End If
                            stbPlayerTable.Append("<td>" & Alive & "</td>")
                            stbPlayerTable.Append("<td>")

                            If DeviceIsAddedToHS Then
                                If (DeviceType = "Pictureshow" Or DeviceType = "Player" Or DeviceType = "Message") Then 'Or GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diRemoteType.ToString, "") = "SonyIRRC") Then
                                    Dim ConfigPlayerBtn As New clsJQuery.jqButton("ConfigBtn" & "_" & PlayerListTableRow.ToString, "Config", MyPageName, False)
                                    ConfigPlayerBtn.toolTip = "Open up the configuration page for this device"
                                    ConfigPlayerBtn.urlNewWindow = True
                                    Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                                    If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                                    If MainInstance <> "" Then
                                        ConfigPlayerBtn.url = "/" & PlayerConfig & ":" & MainInstance & "-" & DeviceUDN
                                    Else
                                        ConfigPlayerBtn.url = "/" & PlayerConfig & ":" & DeviceUDN
                                    End If
                                    stbPlayerTable.Append(ConfigPlayerBtn.Build)
                                End If
                                If GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diRemoteType.ToString, "") = "SonyIRRC" And GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diSonyRemoteRegisterType.ToString, "") = "JSON" Then
                                    stbPlayerTable.Append(BuildSonyAuthenticationOverlay("Authenticate", PlayerListTableRow.ToString, DLNADevice.Key))
                                ElseIf GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocketPIN" Then
                                    Dim AuthenticateBtn As New clsJQuery.jqButton("AuthenticateBtn" & "_" & PlayerListTableRow.ToString, "Get PIN", MyPageName, False)
                                    AuthenticateBtn.toolTip = "Click to open the PIN page on your TV"
                                    stbPlayerTable.Append(AuthenticateBtn.Build)
                                    PINBox = New clsJQuery.jqTextBox("PINBox_" & PlayerListTableRow.ToString, "text", "", MyPageName, 4, False)
                                    PINBox.toolTip = "Click on the autenticate button and enter the PIN here as displayed on your TV"
                                    PINBox.defaultText = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diSamsungRemotePIN.ToString, "")
                                    stbPlayerTable.Append(PINBox.Build)
                                End If
                                stbPlayerTable.Append("<td Align='middle'>" & BuildPlayerRemoveOverlay("Are you sure?", PlayerListTableRow.ToString) & "</td>")
                            Else
                                Dim AddPlayerBtn As New clsJQuery.jqButton("AddPlayerBtn" & "_" & PlayerListTableRow.ToString, "Add", MyPageName, False)
                                AddPlayerBtn.toolTip = "Click to add this device to the HS DB and create HS devices"
                                stbPlayerTable.Append(AddPlayerBtn.Build)
                                stbPlayerTable.Append("<td Align='middle'>" & BuildPlayerDeleteOverlay("Are you sure?", PlayerListTableRow.ToString) & "</td>")
                            End If
                            stbPlayerTable.Append("</td></tr>")
                            PlayerListTableRow += 1
                        End If
                    Next
                End If
            Catch ex As Exception
                Log("Error in Page load building the player list with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            'stb.Append("<hr /> ")
            stbPlayerTable.Append("</table>")
            BuildAllPlayersDeleteOverlay("Are you sure?")
            'stbPlayerTable.Append("</br>" & AddSlideShowDeviceBtn.Build & BuildMediaDeviceOverlay() & ResetPingCountersBtn.Build & BuildAllPlayersDeleteOverlay("Are you sure?") & RediscoverPlayersBtn.Build & LogNetworkBtn.Build)
            Dim UPnPViewerBtn As New clsJQuery.jqButton("UPnPViewerBtn", "View UPNP Devices", MyPageName, False)
            UPnPViewerBtn.toolTip = "Open up the configuration page for this device"
            UPnPViewerBtn.urlNewWindow = True
            Dim HTTPPort_ As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
            If HTTPPort_ <> "" Then HTTPPort_ = ":" & HTTPPort_
            If MainInstance <> "" Then
                'UPnPViewerBtn.url = "http://" & hs.GetIPAddress & HTTPPort_ & "/" & UPnPViewPage & ":" & MainInstance
                UPnPViewerBtn.url = "/" & UPnPViewPage & ":" & MainInstance
            Else
                ' UPnPViewerBtn.url = "http://" & hs.GetIPAddress & HTTPPort_ & "/" & UPnPViewPage
                UPnPViewerBtn.url = "/" & UPnPViewPage
            End If
            stbPlayerTable.Append("</br>" & AddSlideShowDeviceBtn.Build & BuildAllPlayersDeleteOverlay("Are you sure?") & LogNetworkBtn.Build & UPnPViewerBtn.Build)
            stb.Append(clsPageBuilder.FormStart("PlayerListSlideform", MyPageName, "post"))
            Dim stpl As New clsJQuery.jqSlidingTab("myPlayerListSlide", MyPageName, False)
            stpl.initiallyOpen = GetBooleanIniFile("Options", "PlayerListSliderOpen", False)
            stpl.toolTip = "Player List"
            stpl.tab.name = "myPlayerListSlide_name"
            stpl.tab.tabName.Unselected = clsPageBuilder.DivStart("PlayerListPanel", "style=""color:#0000FF"" ") & "<h3>Player Table</h3>" & clsPageBuilder.DivEnd
            stpl.tab.tabName.Selected = clsPageBuilder.DivStart("PlayerListPanel", "style=""color:#0000FF"" ") & "<h3>Player Table</h3>" & clsPageBuilder.DivEnd & "</br>" & stbPlayerTable.ToString
            stb.Append(stpl.Build)
            stb.Append("<hr /> ")

            stb.Append(clsPageBuilder.FormEnd)

        Catch ex As Exception
            Log("Error in Page load with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' add the body html to the page
        Me.AddBody(stb.ToString)

        Me.AddFooter(hs.GetPageFooter)
        Me.suppressDefaultFooter = True

        ' return the full page
        Return Me.BuildPage()

    End Function

    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostBackProc for PluginControl called with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(data)

        If parts IsNot Nothing Then
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(Part, "_")
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PluginControl found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PluginControl found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                        Dim ObjectValue As String = parts(Part)
                        Select Case ObjectNameParts(0).ToString
                            Case "PIDebugLvlBox"
                                Try
                                    WriteIntegerIniFile("Options", "PIDebugLevel", Val(ObjectValue))
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving PIDebugLevel with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                UPnPDebuglevel = GetIntegerIniFile("Options", "PIDebugLevel", DebugLevel.dlErrorsOnly)
                            Case "UPnPDebugLvlBox"
                                Try
                                    WriteIntegerIniFile("Options", "UPnPDebugLevel", Val(ObjectValue))
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving UPnPDebugLevel with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                UPnPDebuglevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly)
                            Case "LogToDiskChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "LogToDisk", ObjectValue.ToUpper = "CHECKED")
                                    gLogToDisk = GetBooleanIniFile("Options", "LogToDisk", False)
                                    If gLogToDisk Then OpenLogFile(DebugLogFileName) Else CloseLogFile()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving LogToDisk flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "VolumeStepBox"
                                Try
                                    WriteStringIniFile("Options", "VolumeStep", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving VolumeStep. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "SpeakerProxyCheckBox"
                                Try
                                    WriteBooleanIniFile("SpeakerProxy", "Active", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving Speaker Proxy flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "UPNPNbrOfElementsBox"
                                Try
                                    WriteStringIniFile("Options", "MaxNbrofUPNPObjects", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving Max Nbr of UPNP Objects. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "HSizeBox"
                                Try
                                    WriteStringIniFile("Options", "ArtworkHSize", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving Artwork Size Info. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "VSizeBox"
                                Try
                                    WriteStringIniFile("Options", "ArtworkVsize", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving Artwork Size Info. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "DoRediscoveryChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "NoRediscovery", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving NoRediscovery flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "FailPingChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "ShowFailedPings", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving ShowFailedPings with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "FailPingCountBox"
                                Try
                                    WriteStringIniFile("Options", "NbrOfPingRetries", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving NbrOfPingRetries with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "MaxAnnTimeBox"
                                Try
                                    WriteStringIniFile("Options", "MaxAnnouncementTime", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving Maximum Announcement. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "myPlayerListSlide"
                                If ObjectValue = "myPlayerListSlide_name_open" Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc has open slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "PlayerListSliderOpen", True)
                                ElseIf ObjectValue = "myPlayerListSlide_name_close" Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc has closed slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "PlayerListSliderOpen", False)
                                End If
                            Case "AddPlayerBtn"
                                ItemChange(DeviceTableItems.ptiAddBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ConfigBtn"
                                ItemChange(DeviceTableItems.ptiConfigBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "DeviceGNBox"
                                ItemChange(DeviceTableItems.ptiDeviceGivenName, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovPlayerRemSubmit"
                                ItemChange(DeviceTableItems.ptiRemoveBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovPlayerDelSubmit"
                                ItemChange(DeviceTableItems.ptiDeleteBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovDelAllSubmit"
                                DeleteAllPlayersClick()
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovConfirmCancel", "ovPlayerRemCancel", "ovPlayerDelCancel", "ovDelAllCancel"
                                Me.pageCommands.Add("refresh", "true")
                            Case "AddSlideShowDeviceBtn"
                                PIReference.CreateSlideshowDevice()
                                Me.pageCommands.Add("refresh", "true")
                            Case "AddMediaDeviceBtn"
                                PIReference.CreateMediaDevice("ROKU", "", "8060")
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovAddMediaDeviceSubmit"
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovAddMediaDeviceCancel"
                                Me.pageCommands.Add("refresh", "true")
                                ' ovIPAddressBox_
                            Case "ovIPAddressBox"
                                ItemChange(DeviceTableItems.ptiEnterSonyPINBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovEnterSonyPINOverlaySubmit"
                                ItemChange(DeviceTableItems.ptiSonyAuthenticate, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovEnterSonyPINOverlayCancel"
                                Me.pageCommands.Add("refresh", "true")
                            Case "PINBox"
                                ItemChange(DeviceTableItems.ptiEnterSamsungPINBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "AuthenticateBtn"
                                ItemChange(DeviceTableItems.ptiSamsungAuthenticate, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "LogNetworkBtn"
                                PIReference.DisplayUPnPDevices()
                                Me.pageCommands.Add("refresh", "true")
                            Case "HSSTrackLengthBox"
                                Try
                                    WriteIntegerIniFile("Options", "HSSTrackLengthSetting", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving HS TrackLength Settings. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "HSSTrackPositionBox"
                                Try
                                    WriteIntegerIniFile("Options", "HSSTrackPositionSetting", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for Plugin saving HS TrackPosition Settings. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in postBackProc for PluginControl processing with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            ' call the plug-in to have it read and apply the new settings
            Try
                PIReference.ReadIniFile()
            Catch ex As Exception
                Log("Error in postBackProc for PluginControl calling Plugin to update values with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PluginControl found parts to be empty", LogType.LOG_TYPE_INFO)
        End If

        Return MyBase.postBackProc(page, data, user, userRights)
    End Function

    Public Enum DeviceTableItems
        ptiDeviceGivenName = 0
        ptiAddBtn = 1
        ptiRemoveBtn = 2
        ptiConfigBtn = 3
        ptiDeleteBtn = 4
        ptiEnterSonyPINBtn = 5
        ptiSonyAuthenticate = 6
        ptiEnterSamsungPINBtn = 7
        ptiSamsungAuthenticate = 8
    End Enum

    Public Sub ItemChange(DeviceTableItem As DeviceTableItems, Value As String, RowIndex As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ItemChange called with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString, LogType.LOG_TYPE_INFO)
        Value = Trim(Value)
        Dim KeyValue As New System.Collections.Generic.KeyValuePair(Of String, String)

        Try
            Select Case DeviceTableItem
                Case DeviceTableItems.ptiDeviceGivenName
                    ' value holds the index into the Sonos Zone Names
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        PIReference.UpdateDeviceName(DeviceUDN, Value)
                    Catch ex As Exception
                    End Try
                Case DeviceTableItems.ptiConfigBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Dim DeviceAPIIndex As Integer = 0
                    Dim DeviceServiceType As String = ""
                    If DeviceUDN <> "" Then
                        Try
                            DeviceAPIIndex = GetIntegerIniFile(DeviceUDN, "diDeviceAPIIndex", 0)
                            DeviceServiceType = GetStringIniFile(DeviceUDN, "diDeviceType", "")
                        Catch ex As Exception
                            Log("Error in ItemChange geting the DeviceAPIIndex for ConfigBtn with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Sub
                        End Try
                    End If
                    Log("ItemChange for ConfigBtn retrieved from ini file DeviceAPIINdex = " & DeviceAPIIndex.ToString, LogType.LOG_TYPE_INFO)
                    If DeviceAPIIndex <> 0 Then
                        If DeviceServiceType = "DMS" Then
                            'Me.pageCommands.Add.Transfer("DLNA_ManagerDMSPage.aspx")
                        ElseIf DeviceServiceType = "DMR" Or DeviceServiceType = "HST" Or DeviceServiceType = "DIAL" Or DeviceServiceType = "RCR" Or DeviceServiceType = "PMR" Then
                            'Server.Transfer("DLNA_ManagerDMRPage.aspx")
                        End If
                    End If
                Case DeviceTableItems.ptiAddBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        PIReference.AddDevicetoHS(DeviceUDN)
                    Catch ex As Exception
                        Log("Error in ItemChange for AddBtn with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiRemoveBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        PIReference.RemoveDevicefromHS(DeviceUDN)
                    Catch ex As Exception
                        Log("Error in ItemChange for RemoveBtn with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiDeleteBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        PIReference.DeleteDevice(DeviceUDN)
                    Catch ex As Exception
                        Log("Error in ItemChange for DeleteBtn with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiEnterSonyPINBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, Value)
                    Catch ex As Exception
                        Log("Error in ItemChange for EnterSonyPINBTN with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiSonyAuthenticate
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        PIReference.AuthenticateSony(DeviceUDN, GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, ""))
                    Catch ex As Exception
                        Log("Error in ItemChange for ptiSonyAuthenticate with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiEnterSamsungPINBtn
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    Try
                        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, Value)
                        PIReference.AuthenticateSamsung(DeviceUDN, GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, ""))
                    Catch ex As Exception
                        Log("Error in ItemChange for EnterSamsungPINBTN with DeviceTableItems = " & DeviceTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case DeviceTableItems.ptiSamsungAuthenticate
                    Dim DeviceUDN As String = GetZoneUDNByIndex(RowIndex)
                    PIReference.SendOpenPINtoCorrectSamsungDevice(DeviceUDN)
            End Select
        Catch ex As Exception
            Log("Error in ItemChange case selector with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeletePlayerClick(PlayerTableItem As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeletePlayerClick called with tableItem = " & PlayerTableItem.ToString, LogType.LOG_TYPE_INFO)
        PlayerTableItem = Trim(PlayerTableItem)
        ' go find the UDN
        Dim IndexCount As Integer = 0
        If PIReference.MyHSDeviceLinkedList.Count > 0 Then
            For Each HSDevice As MyUPnpDeviceInfo In PIReference.MyHSDeviceLinkedList
                If Not HSDevice Is Nothing And (IndexCount = PlayerTableItem) Then
                    Dim Player As HSPI = HSDevice.UPnPDeviceControllerRef
                    Dim UDN As String = Player.DeviceUDN
                    If MainInstance <> "" Then
                        RemoveInstance(MainInstance & "-" & UDN)
                    Else
                        RemoveInstance(UDN)
                    End If
                    PIReference.DeleteDevice(UDN)
                    Exit Sub
                End If
                IndexCount += 1
            Next
        End If
    End Sub

    Public Sub DeleteAllPlayersClick()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteAllPlayersClick called", LogType.LOG_TYPE_INFO)
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
        For Each DLNADevice In DLNADevices
            If DLNADevice.Key <> "" Then
                Try
                    Dim Player As HSPI = PIReference.GetAPIByUDN(DLNADevice.Key)
                    If Player IsNot Nothing Then
                        Player.DeleteWebLink(DLNADevice.Key, GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diGivenName.ToString, ""))
                        If MainInstance <> "" Then
                            RemoveInstance(MainInstance & "-" & DLNADevice.Key)
                        Else
                            RemoveInstance(DLNADevice.Key)
                        End If
                    End If
                    DeleteIniSection(DLNADevice.Key)
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DeleteAllPlayersClick with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Next
        PIReference.MyHSDeviceLinkedList.Clear()
        DeleteIniSection("UPnP Devices UDN to Info")
        DeleteIniSection("UPnP HSRef to UDN")
        DeleteIniSection("UPnP UDN to HSRef")
        DeleteIniSection("Message Devices")
        DeleteIniSection("Message Service by UDN")
        DeleteIniSection("Remote Service by UDN")
        DeleteIniSection("Party Devices")
        DeleteIniSection("DevicePage")
        DeleteIniSection("Speaker Devices")
        WriteIntegerIniFile("Settings", "MasterHSDeviceRef", -1) ' reset the master code, next time Sonos initializes everything will be cleaned out.
        hs.DeleteIODevices(sIFACE_NAME, MainInstance)
        ' Force HomeSeer to save changes to devices and events so we can find our new device
        hs.SaveEventsDevices()
        MyShutDownRequest = True
    End Sub

    Private Function GetZoneNameByIndex(index As Integer) As String
        GetZoneNameByIndex = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetZoneNameByIndex called with Index = " & index.ToString, LogType.LOG_TYPE_INFO)
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info")
        Try
            If Not DLNADevices Is Nothing And DLNADevices.Count > index Then
                Dim IndexCount As Integer = 0
                Dim DLNADevice As System.Collections.Generic.KeyValuePair(Of String, String) = Nothing
                For Each DLNADevice In DLNADevices
                    If IndexCount = index Then Exit For
                    IndexCount += 1
                Next
                If IndexCount = index Then
                    Return DLNADevice.Value
                End If
            End If
        Catch ex As Exception
            Log("Error in GetZoneNameByIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function GetZoneUDNByIndex(index As Integer) As String
        GetZoneUDNByIndex = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetZoneUDNByIndex called with Index = " & index.ToString, LogType.LOG_TYPE_INFO)
        Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
        DLNADevices = GetIniSection("UPnP Devices UDN to Info")
        Try
            If Not DLNADevices Is Nothing And DLNADevices.Count > index Then
                Dim IndexCount As Integer = 0
                Dim DLNADevice As System.Collections.Generic.KeyValuePair(Of String, String) = Nothing
                For Each DLNADevice In DLNADevices
                    If IndexCount = index Then Exit For
                    IndexCount += 1
                Next
                If IndexCount = index Then
                    Return DLNADevice.Key
                End If
            End If
        Catch ex As Exception
            Log("Error in GetZoneUDNByIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Private Function GetHeadContent() As String
        Try
            Return hs.GetPageHeader(sIFACE_NAME, "", "", False, False, True, False, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function GetFooterContent() As String
        Try
            Return hs.GetPageFooter(False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function GetBodyContent() As String
        Try
            Return hs.GetPageHeader(StrConv(sIFACE_NAME, VbStrConv.ProperCase), "", "", False, False, False, True, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function BuildPlayerDeleteOverlay(HeaderText As String, ButtonSuffix As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("PlayerDelOverlay" & ButtonSuffix.ToString, MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "Delete this Player from the plug-in DB. Used to remove stale entries in plug-in DB. If player is still on the network it will reappear upon restart/rediscovery"
        ConfirmOverlay.label = "Delete"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("PlayerDeleteConfirmform" & ButtonSuffix.ToString, MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovPlayerDelSubmit_" & ButtonSuffix, "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovPlayerDelCancel_" & ButtonSuffix, "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildPlayerRemoveOverlay(HeaderText As String, ButtonSuffix As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("PlayerRemOverlay" & ButtonSuffix.ToString, MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "Remove this device from the HS DB and delete all associated HS Devices. Warning this will invalidate all events and HS Touch screens that use this device"
        ConfirmOverlay.label = "Remove"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("PlayerRemoveConfirmform" & ButtonSuffix.ToString, MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovPlayerRemSubmit_" & ButtonSuffix, "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovPlayerRemCancel_" & ButtonSuffix, "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildAllPlayersDeleteOverlay(HeaderText As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("DellAllOverlay", MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "This will delete all devices, including the root device. Upon completion the plugin will terminate and will have to be restarted"
        ConfirmOverlay.label = "Delete All Devices"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("DelAllPlayersConfirmform", MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovDelAllSubmit", "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovDelAllCancel", "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildMediaDeviceOverlay() As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("AddMediaDeviceOverlay", MyPageName, False, "events_overlay")

        Dim IPAddressBox As clsJQuery.jqTextBox
        IPAddressBox = New clsJQuery.jqTextBox("IPAddressBox", "text", "", MyPageName, 16, False)
        IPAddressBox.toolTip = "Specify the IP Address of the Media device"
        Dim IPPortBox As clsJQuery.jqTextBox
        IPPortBox = New clsJQuery.jqTextBox("IPPortBox", "text", "", MyPageName, 4, False)
        IPPortBox.toolTip = "Specify the IP Port of the Media device"

        Dim MediaDeviceTypeBox As clsJQuery.jqDropList
        MediaDeviceTypeBox = New clsJQuery.jqDropList("MediaDeviceTypeBox", MyPageName, True)
        MediaDeviceTypeBox.toolTip = "Select the Media Device Type"
        MediaDeviceTypeBox.ClearItems()
        MediaDeviceTypeBox.AddItem("No Selection", "", True)
        MediaDeviceTypeBox.AddItem("Roku", "Roku", False)

        ConfirmOverlay.toolTip = "Add this Media device. Make sure you selected the type and set IP Address and IP Port"
        ConfirmOverlay.label = "Add Media Device"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("AddMediaDeviceConfirmform", MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & "Enter Additional info to create the Media Device" & "<br><br>"
        ConfirmOverlay.overlayHTML &= MediaDeviceTypeBox.Build & "Select the Media Device Type <br></br>"
        ConfirmOverlay.overlayHTML &= IPAddressBox.Build & "Set the IP Address in the form of xxx.xxx.xxx.xxx <br></br>"
        ConfirmOverlay.overlayHTML &= IPPortBox.Build & "Set the IP Port <br></br>"

        Dim tbut1 As New clsJQuery.jqButton("ovAddMediaDeviceSubmit", "Submit", MyPageName, True)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovAddMediaDeviceCancel", "Cancel", MyPageName, True)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildSonyAuthenticationOverlay(HeaderText As String, ButtonSuffix As String, DeviceUDN As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("EnterSonyPINOverlay" & ButtonSuffix.ToString, MyPageName, False, "events_overlay")
        Dim PINBox As clsJQuery.jqTextBox
        PINBox = New clsJQuery.jqTextBox("ovIPAddressBox_" & ButtonSuffix.ToString, "text", GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, ""), MyPageName, 4, False)
        PINBox.editable = True
        PINBox.toolTip = "Enter the 4 digit PIN number that is shown on your TV"
        PINBox.enabled = True
        ConfirmOverlay.toolTip = "Click here to Authenticate this Plug-In with your Sony TV. Enter the PIN as displayed on your screen"
        ConfirmOverlay.label = "Authenticate"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("EnterSonyPINOverlayform" & ButtonSuffix.ToString, MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        ConfirmOverlay.overlayHTML &= PINBox.Build & "Enter the 4 digit PIN <br></br>"
        Dim tbut1 As New clsJQuery.jqButton("ovEnterSonyPINOverlaySubmit_" & ButtonSuffix, "Submit", MyPageName, True)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovEnterSonyPINOverlayCancel_" & ButtonSuffix, "Cancel", MyPageName, True)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

End Class
