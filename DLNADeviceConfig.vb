Imports Scheduler
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Text
Imports System.Web
Imports System.Web.UI.HtmlControls

Public Class DLNADeviceConfig
    Inherits clsPageBuilder

    Private PIReference As HSPI = Nothing
    Private MyZoneUDN As String = ""
    Private MusicAPI As HSPI = Nothing
    Private ZoneName As String = ""
    Private MyPageName As String = ""

    Private ServerDeviceNameBox As clsJQuery.jqDropList
    Private PictureSizeNameBox As clsJQuery.jqDropList
    Private TimeBetweenPicturesBox As clsJQuery.jqTextBox
    Private PollTransportChkBox As clsJQuery.jqCheckBox
    Private PollVolumeChkBox As clsJQuery.jqCheckBox
    Private NextAVBtn As clsJQuery.jqCheckBox
    Private AnnouncementDeviceBox As clsJQuery.jqTextBox
    Private MessageDeviceBox As clsJQuery.jqTextBox
    Private UseMP3ChkBox As clsJQuery.jqCheckBox

    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
        MyPageName = pagename
        TimeBetweenPicturesBox = New clsJQuery.jqTextBox("TimeBetweenPicturesBox", "text", "", MyPageName, 5, False)
        TimeBetweenPicturesBox.toolTip = "Specify the amount in seconds for the picture show to step through your pictures"
        PollTransportChkBox = New clsJQuery.jqCheckBox("PollTransportChkBox", " Poll Transport Changes", MyPageName, True, False)
        PollVolumeChkBox = New clsJQuery.jqCheckBox("PollVolumeChkBox", " Poll Volume Changes", MyPageName, True, False)
        ServerDeviceNameBox = New clsJQuery.jqDropList("ContentDeviceNameBox", MyPageName, False)
        PictureSizeNameBox = New clsJQuery.jqDropList("PictureSizeNameBox", MyPageName, False)
        NextAVBtn = New clsJQuery.jqCheckBox("NextAVBtn", " Use Next Method", MyPageName, True, False)
        AnnouncementDeviceBox = New clsJQuery.jqTextBox("AnnouncementDeviceBox", "text", "", MyPageName, 30, False)
        MessageDeviceBox = New clsJQuery.jqTextBox("MessageDeviceBox", "text", "", MyPageName, 30, False)
        UseMP3ChkBox = New clsJQuery.jqCheckBox("UseMP3ChkBox", " Use MP3 Format", MyPageName, True, False)
    End Sub

    Public WriteOnly Property RefToPlugIn
        Set(value As Object)
            PIReference = value
        End Set
    End Property

    Public WriteOnly Property ZoneUDN As String
        Set(value As String)
            MyZoneUDN = value
            Try
                MusicAPI = PIReference.GetAPIByUDN(MyZoneUDN)
            Catch ex As Exception
                Log("Error in GetPagePlugin getting the MusicAPI for ZoneUDN = " & MyZoneUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If MusicAPI Is Nothing Then
                Log("Error in GetPagePlugin, MusicAPI not found for ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_ERROR)
            End If
            ZoneName = MusicAPI.DeviceName
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for Zoneplayer = " & ZoneName & " set ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    ' build and return the actual page
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, GenerateHeaderFooter As Boolean) As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for DLNADeviceControl called for Zoneplayer = " & ZoneName & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)
        Dim stb As New StringBuilder
        Me.reset()
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for DLNADeviceControl called for ZoneUDN = " & MyZoneUDN & " and ZoneName = " & ZoneName & " and PageName = " & MyPageName, LogType.LOG_TYPE_INFO)
        Dim DeviceType As String = GetStringIniFile(MyZoneUDN, DeviceInfoIndex.diDeviceType.ToString, "")
        Try
            ' handle any queries like mode=something
            Dim parts As Collections.Specialized.NameValueCollection = Nothing
            If (queryString <> "") Then
                parts = HttpUtility.ParseQueryString(queryString)
            End If

            Try
                ZoneName = MusicAPI.DeviceName
            Catch ex As Exception
                Log("Error in GetPagePlugin getting ZoneName. Error  = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Return ""
            End Try

            If DeviceType <> "PMR" Then
                Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
                Try
                    DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
                    If DLNADevices Is Nothing Then
                        Log("Error in GetPagePlugin for Player = " & ZoneName & ". No Devices specified in the .ini file under ""UPnP Devices UDN to Info""", LogType.LOG_TYPE_ERROR)
                        DLNADevices = Nothing
                        ' set all the object to invisible
                    Else
                        Dim DLNADevice 'As System.Collections.Generic.Dictionary(Of String, String)
                        Dim DLNAServerName As String = ""
                        Dim ServerUDN As String = GetStringIniFile(MyZoneUDN, DeviceInfoIndex.diServerUDN.ToString, "")
                        ServerDeviceNameBox.ClearItems()
                        ServerDeviceNameBox.AddItem("No Selection", "No Selection", ServerUDN = "")
                        For Each DLNADevice In DLNADevices
                            If DLNADevice.Key <> "" Then
                                If GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceType.ToString, "") = "DMS" And GetBooleanIniFile(DLNADevice.Key, DeviceInfoIndex.diDeviceIsAdded.ToString, False) Then ' is it a DMR and is it added to HS
                                    ' this is a Media server
                                    DLNAServerName = GetStringIniFile(DLNADevice.Key, DeviceInfoIndex.diGivenName.ToString, "")
                                    ServerDeviceNameBox.AddItem(DLNAServerName, DLNADevice.Key, DLNADevice.Key = ServerUDN)
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Log("Error in GetPagePlugin for Player = " & ZoneName & " retrieving the ServerAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Dim PictureSizeIndex As String = GetIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psDefault)
                Try
                    PictureSizeNameBox.ClearItems()
                    PictureSizeNameBox.AddItem("Default", "Default", PictureSizeIndex = PictureSize.psDefault)
                    PictureSizeNameBox.AddItem("Tiny", "Tiny", PictureSizeIndex = PictureSize.psTiny)
                    PictureSizeNameBox.AddItem("Small", "Small", PictureSizeIndex = PictureSize.psSmall)
                    PictureSizeNameBox.AddItem("Medium", "Medium", PictureSizeIndex = PictureSize.psMedium)
                    PictureSizeNameBox.AddItem("Large", "Large", PictureSizeIndex = PictureSize.psLarge)
                Catch ex As Exception
                End Try
            End If

            If GenerateHeaderFooter Then
                Me.AddHeader(hs.GetPageHeader(MyPageName, "Device Config", "", "", False, True))
            End If

            stb.Append(clsPageBuilder.FormStart("DeviceConfigform", MyPageName, "post"))

            stb.Append(clsPageBuilder.DivStart(MyPageName, ""))
            'Me.RefreshIntervalMilliSeconds = 2000
            'stb.Append(Me.AddAjaxHandlerPost("action=updatetime", MyPageName))
            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd) ' ErrorMessage

            ' specific page starts here

            TimeBetweenPicturesBox.defaultText = GetStringIniFile(MyZoneUDN, DeviceInfoIndex.diTimeBetweenPictures.ToString, "")
            PollTransportChkBox.checked = GetBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diPollTransportChanges.ToString, False)
            PollVolumeChkBox.checked = GetBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diPollVolumeChanges.ToString, False)
            NextAVBtn.enabled = GetBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diNextAV.ToString, False)
            NextAVBtn.checked = GetBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diUseNextAV.ToString, False)

            AnnouncementDeviceBox.defaultText = GetStringIniFile("Speaker Devices", MyZoneUDN, "")
            UseMP3ChkBox.checked = GetBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diAnnouncementMP3.ToString, False)
            If DeviceType <> "DMR" Then
                PollTransportChkBox.enabled = False
                PollVolumeChkBox.enabled = False
                AnnouncementDeviceBox.enabled = False
                UseMP3ChkBox.enabled = False
                MessageDeviceBox.enabled = False
            End If
            ' MessageDeviceBox
            MessageDeviceBox.defaultText = GetStringIniFile("Message Devices", MyZoneUDN, "")
            If DeviceType = "PMR" Then
                MessageDeviceBox.enabled = True
            End If

            stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
            stb.Append("<h1> Media Controller Plugin Configuration</h1>" & vbCrLf)
            stb.Append("<h1><img src=" & MusicAPI.PlayerIconURL & " style='height:50px; width:50px'>    ")
            stb.Append(ZoneName & "</h1>")
            stb.Append(clsPageBuilder.DivEnd)

            stb.Append("<table style='width: 50%' >")

            stb.Append("<tr><td colspan='2'>")
            stb.Append("<hr /> ")

            If DeviceType <> "PMR" Then
                stb.Append(clsPageBuilder.DivStart("QBehaviorPanel", "style=""color:#0000FF"" "))
                stb.Append("<h3>Picture Queue settings</h3>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)
                stb.Append(TimeBetweenPicturesBox.Build)
                stb.Append(" Time between pictures in seconds")
                stb.Append("</br>")
                stb.Append("<hr /> ")
                stb.Append(clsPageBuilder.DivStart("PlayerSettingPanel", "style=""color:#0000FF"" "))
                stb.Append("<h3>Player Settings</h3>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)
                stb.Append(PollTransportChkBox.Build & "</br>")
                'stb.Append("Poll Transport State Flag<hr />")
                stb.Append(PollVolumeChkBox.Build & "</br>")
                'stb.Append("Poll Volume State Flag<hr />")
                stb.Append(NextAVBtn.Build & "</br>")
                'stb.Append("Use Next Method<hr />")
                stb.Append(ServerDeviceNameBox.Build)
                stb.Append("Select Default Content Server</br>")
                stb.Append(PictureSizeNameBox.Build)
                stb.Append("Select Picture Size<hr />")
                stb.Append(clsPageBuilder.DivStart("AnnouncementPanel", "style=""color:#0000FF"" "))
                stb.Append("<h3>Announcement Settings</h3>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)
                stb.Append(AnnouncementDeviceBox.Build)
                stb.Append("Speaker Device IDs</br>")
                If Not HSisRunningOnLinux Then stb.Append(UseMP3ChkBox.Build & "</br>")
            Else
                stb.Append(clsPageBuilder.DivStart("AnnouncementPanel", "style=""color:#0000FF"" "))
                stb.Append("<h3>Announcement Settings</h3>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)
                stb.Append(MessageDeviceBox.Build)
                stb.Append("Message Device IDs</br>")
            End If

            stb.Append("</br></br>")
            stb.Append("</td></tr></table>")
            stb.Append(clsPageBuilder.FormEnd)

            If GenerateHeaderFooter Then
                ' add the body html to the page
                Me.AddBody(stb.ToString)
                Me.AddFooter(hs.GetPageFooter)
                Me.suppressDefaultFooter = True
                ' return the full page
                Return Me.BuildPage()
            End If


        Catch ex As Exception
            Log("Error in GetPagePlugin for DLNADeviceControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return stb.ToString

    End Function

    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostBackProc for DLNADeviceControl called  for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(System.Web.HttpUtility.HtmlDecode(data))
        'parts = HttpUtility.ParseQueryString(data)

        If parts IsNot Nothing Then
            'Log("PostBackProc for DLNADeviceControl called  for Player = " & ZoneName & " with part = '" & parts.GetKey(0).ToUpper.ToString & "'", LogType.LOG_TYPE_INFO)
            If (data.ToString.ToUpper <> "ACTION=UPDATETIME") And (PIDebuglevel > DebugLevel.dlErrorsOnly = True) Then Log("PostBackProc for DLNADeviceControl called for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(HttpUtility.UrlDecode(Part), "_")
                        If (data.ToString.ToUpper <> "ACTION=UPDATETIME") Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for DLNADeviceControl for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for DLNADeviceControl for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                        End If
                        Dim ObjectValue As String = HttpUtility.UrlDecode(parts(Part))
                        'Dim ObjectValue As String = parts(Part)

                        Select Case ObjectNameParts(0).ToString.ToUpper
                            Case "INSTANCE"
                            Case "REF"
                            Case "ID"
                            Case "ACTION"
                                If ObjectValue = "updatetime" Then
                                    'CheckForChanges()
                                End If
                            Case "TIMEBETWEENPICTURESBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued TimeBetweenPicturesBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diTimeBetweenPictures.ToString, Val(ObjectValue.ToString))
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving TimeBetweenPicturesBox with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "POLLTRANSPORTCHKBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued PollTransportChkBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diPollTransportChanges.ToString, ObjectValue.ToUpper = "CHECKED")
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving PollTransportChkBox flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "POLLVOLUMECHKBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued PollVolumeChkBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diPollVolumeChanges.ToString, ObjectValue.ToUpper = "CHECKED")
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving PollVolumeChkBox flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "CONTENTDEVICENAMEBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued ContentDeviceNameBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteStringIniFile(MyZoneUDN, DeviceInfoIndex.diServerUDN.ToString, ObjectValue.ToString)
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving ContentDeviceNameBox with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "PICTURESIZENAMEBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued PictureSizeNameBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    Select Case ObjectValue.ToString.ToUpper
                                        Case "DEFAULT"
                                            WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psDefault)
                                        Case "TINY"
                                            WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psTiny)
                                        Case "SMALL"
                                            WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psSmall)
                                        Case "MEDIUM"
                                            WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psMedium)
                                        Case "LARGE"
                                            WriteIntegerIniFile(MyZoneUDN, DeviceInfoIndex.diPictureSize.ToString, PictureSize.psLarge)
                                    End Select
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving PictureSizeNameBox with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "NEXTAVBTN"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued NextAVBtn command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diUseNextAV.ToString, ObjectValue.ToUpper = "CHECKED")
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving NextAVBtn flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "ANNOUNCEMENTDEVICEBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued AnnouncementDeviceBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteStringIniFile("Speaker Devices", MyZoneUDN, ObjectValue.ToString)
                                    'MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving AnnouncementDeviceBox. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "MESSAGEDEVICEBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued MessageDeviceBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteStringIniFile("Message Devices", MyZoneUDN, ObjectValue.ToString)
                                    'MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving MessageDeviceBox. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "USEMP3CHKBOX"
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued UseMP3ChkBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                Try
                                    WriteBooleanIniFile(MyZoneUDN, DeviceInfoIndex.diAnnouncementMP3.ToString, ObjectValue.ToUpper = "CHECKED")
                                    MusicAPI.ReadDeviceIniSettings()
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving UseMP3ChkBox flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try

                            Case Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_WARNING)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_WARNING)
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in postBackProc for Player = " & ZoneName & " processing page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for Player = " & ZoneName & " found parts to be empty", LogType.LOG_TYPE_INFO)
        End If
        Return MyBase.postBackProc(page, data, user, userRights)
    End Function


End Class
