Imports HomeSeerAPI
Imports Scheduler
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Text
Imports System.Web
Imports System.Net
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Class PlayerControl
    Inherits clsPageBuilder

    Private PIReference As HSPI = Nothing
    Private MyZoneUDN As String = ""
    Private MusicAPI As HSPI = Nothing
    Private ServerAPI As HSPI = Nothing
    Dim ar As DBRecord() = Nothing
    Private arplaylist() As String
    Private ZoneName As String = ""
    Private MyPageName As String = ""

    Private LblTrackName As String = ""
    Private LblArtistName As String = ""
    Private LblAlbumName As String = ""
    Private LblNextTrackName As String = ""
    Private LblNextArtistName As String = ""
    Private LblNextAlbumName As String = ""
    Private LblDuration As Integer = 0
    Private LblDescr As String = ""
    Private LblNextDescr As String = ""
    Private LblWhatsLoaded As String = ""
    Private ButRepeatImagePath As String = ""
    Private ButShuffleImagePath As String = ""
    Private ButMuteImagePath As String = ""
    'Private ArtImagePath As String = ""

    Private MyPosition As Integer = 0
    Private MyVolume As Integer = 0
    Private MyTrack As String = ""
    Private MyAlbum As String = ""
    Private MyArtist As String = ""
    Private MyRadiostation As String = ""
    Private MyNextTrack As String = ""
    Private MyNextAlbum As String = ""
    Private MyNextArtist As String = ""
    Private MyArt As String = ""
    Private MyMute As Boolean = False
    Private MyRepeat As Boolean = False
    Private MyShuffle As Boolean = False
    Private MyPlayerState As HSPI.player_state_values = HSPI.player_state_values.Stopped
    Private MyTrackDesc As String = ""
    Private MyNextTrackDesc As String = ""

    Private MyLastSelecteNavBoxItems As String = ""
    Private MyLastSelectedNavBoxClass As String = ""
    Private MyLastSelectedPlaylistName As String = ""

    Private DescriptionCell As String = ""

    Private ButPrev As clsJQuery.jqButton
    Private ButStop As clsJQuery.jqButton
    Private ButPlay As clsJQuery.jqButton
    Private ButNext As clsJQuery.jqButton
    Private ButRew As clsJQuery.jqButton
    Private ButFF As clsJQuery.jqButton
    Private ButMute As clsJQuery.jqButton
    Private ButLoudness As clsJQuery.jqButton
    Private ButRepeat As clsJQuery.jqButton
    Private ButShuffle As clsJQuery.jqButton
    Private ButAddToPlaylist As clsJQuery.jqButton
    Private ButClearList As clsJQuery.jqButton
    Private ButSavePlaylist As clsJQuery.jqButton
    Private ButEditor As clsJQuery.jqButton
    Private NavigationBox As clsJQuery.jqListBoxEx
    Private PlaylistBox As clsJQuery.jqListBoxEx
    Private ArtImageBox As clsJQuery.jqButton
    Private LblRoot As clsJQuery.jqButton
    Private LblLevel1 As clsJQuery.jqButton
    Private LblLevel2 As clsJQuery.jqButton
    Private LblLevel3 As clsJQuery.jqButton
    Private LblLevel4 As clsJQuery.jqButton
    Private LblLevel5 As clsJQuery.jqButton
    Private LblLevel6 As clsJQuery.jqButton
    Private LblLevel7 As clsJQuery.jqButton
    Private LblLevel8 As clsJQuery.jqButton
    Private LblLevel9 As clsJQuery.jqButton
    Private LblLevel10 As clsJQuery.jqButton
    Private ObjectIDRoot As String = ""
    Private ObjectIDLevel1 As String = ""
    Private ObjectIDLevel2 As String = ""
    Private ObjectIDLevel3 As String = ""
    Private ObjectIDLevel4 As String = ""
    Private ObjectIDLevel5 As String = ""
    Private ObjectIDLevel6 As String = ""
    Private ObjectIDLevel7 As String = ""
    Private ObjectIDLevel8 As String = ""
    Private ObjectIDLevel9 As String = ""
    Private ObjectIDLevel10 As String = ""
    Private ContentDeviceNameBox As clsJQuery.jqDropList
    Private PlayListSelector As clsJQuery.jqDropList
    Private InitDone As Boolean = False
    Private SavePlayListAsNameBox As clsJQuery.jqTextBox


    Private Enum ITEM_TYPE
        genres = 1
        artists = 2
        albums = 3
        tracks = 4
        playlists = 5
        radioLists = 6
        Audiobooks = 7
        Podcasts = 8
        Root = 9
        Level1 = 10
        Level2 = 11
        Level3 = 12
        Level4 = 13
        Level5 = 14
        Level6 = 15
        Level7 = 16
        Level8 = 17
        Level9 = 18
        Level10 = 19
    End Enum

    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
        MyPageName = pagename
        NavigationBox = New clsJQuery.jqListBoxEx("NavigationBox", pagename)
        PlaylistBox = New clsJQuery.jqListBoxEx("PlaylistBox", pagename)
        ArtImageBox = New clsJQuery.jqButton("ArtImageBox", "", pagename, False)
        ButPrev = New clsJQuery.jqButton("ButPrev", "Previous", MyPageName, False)
        ButStop = New clsJQuery.jqButton("ButStop", "Stop", MyPageName, False)
        ButPlay = New clsJQuery.jqButton("ButPlay", "Play", MyPageName, False)
        ButNext = New clsJQuery.jqButton("ButNext", "Next", MyPageName, False)
        ButRew = New clsJQuery.jqButton("ButRew", "Rewind", MyPageName, False)
        ButFF = New clsJQuery.jqButton("ButFF", "Fast Forward", MyPageName, False)
        ButMute = New clsJQuery.jqButton("ButMute", "Mute", MyPageName, False)
        ButLoudness = New clsJQuery.jqButton("ButLoudness", "Loudness", MyPageName, False)
        ButRepeat = New clsJQuery.jqButton("ButRepeat", "Repeat", MyPageName, False)
        ButShuffle = New clsJQuery.jqButton("ButShuffle", "Shuffle", MyPageName, False)
        ButAddToPlaylist = New clsJQuery.jqButton("ButAddToPlaylist", "Add To Playlist", MyPageName, False)
        ButClearList = New clsJQuery.jqButton("ButClearList", "Clear List", MyPageName, False)
        ButSavePlaylist = New clsJQuery.jqButton("ButSavePlaylist", "Save Playlist As", MyPageName, False)
        ButEditor = New clsJQuery.jqButton("ButEditor", "ButEditor", MyPageName, False)
        LblRoot = New clsJQuery.jqButton("LblRoot", "", MyPageName, False)
        LblRoot.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel1 = New clsJQuery.jqButton("LblLevel1", "", MyPageName, False)
        LblLevel1.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel2 = New clsJQuery.jqButton("LblLevel2", "", MyPageName, False)
        LblLevel2.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel3 = New clsJQuery.jqButton("LblLevel3", "", MyPageName, False)
        LblLevel3.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel4 = New clsJQuery.jqButton("LblLevel4", "", MyPageName, False)
        LblLevel4.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel5 = New clsJQuery.jqButton("LblLevel5", "", MyPageName, False)
        LblLevel5.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel6 = New clsJQuery.jqButton("LblLevel6", "", MyPageName, False)
        LblLevel6.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel7 = New clsJQuery.jqButton("LblLevel7", "", MyPageName, False)
        LblLevel7.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel8 = New clsJQuery.jqButton("LblLevel8", "", MyPageName, False)
        LblLevel8.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel9 = New clsJQuery.jqButton("LblLevel9", "", MyPageName, False)
        LblLevel9.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        LblLevel10 = New clsJQuery.jqButton("LblLevel10", "", MyPageName, False)
        LblLevel10.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        ContentDeviceNameBox = New clsJQuery.jqDropList("ContentDeviceNameBox", MyPageName, False)
        PlayListSelector = New clsJQuery.jqDropList("PlayListSelector", MyPageName, False)
        SavePlayListAsNameBox = New clsJQuery.jqTextBox("SavePlayListAsNameBox", "text", "", MyPageName, 25, False)
        SavePlayListAsNameBox.toolTip = "Enter here a PlaylistName to SaveAs or leave empty to Save under current select PlayListName"
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
                Exit Property
            End Try
            If MusicAPI Is Nothing Then
                Log("Error in GetPagePlugin, MusicAPI not found for ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_ERROR)
                Exit Property
            End If
            ZoneName = MusicAPI.DeviceName
            If g_bDebug Then Log("GetPagePlugin for Zoneplayer = " & ZoneName & " set ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    ' build and return the actual page
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, GenerateHeaderFooter As Boolean) As String
        'If g_bDebug Then Log("GetPagePlugin for PlayerControl called for Zoneplayer = " & ZoneName & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)
        Dim stb As New StringBuilder
        Me.reset()
        'If g_bDebug Then Log("GetPagePlugin for PlayerControl called for ZoneUDN = " & MyZoneUDN & " and ZoneName = " & ZoneName & " and PageName = " & MyPageName, LogType.LOG_TYPE_INFO)

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

            MyLastSelectedPlaylistName = GetStringIniFile("DevicePage", "playlistSelector_" & MyZoneUDN, "")
            LoadPlayListSelector(MyLastSelectedPlaylistName)

            Dim DLNADevices As New System.Collections.Generic.Dictionary(Of String, String)()
            ServerAPI = Nothing
            Try
                DLNADevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
                If DLNADevices Is Nothing Then
                    Log("Error in GetPagePlugin for Player = " & ZoneName & ". No Devices specified in the .ini file under ""UPnP Devices UDN to Info""", LogType.LOG_TYPE_ERROR)
                    DLNADevices = Nothing
                    ' set all the object to invisible
                Else
                    Dim DLNADevice 'As System.Collections.Generic.Dictionary(Of String, String)
                    Dim DLNAServerName As String = ""
                    Dim ServerUDN As String = GetStringIniFile("DevicePage", MyZoneUDN & "_ServerUDN", "")
                    ContentDeviceNameBox.ClearItems()
                    ContentDeviceNameBox.AddItem("No Selection", "No Selection", ServerUDN = "")
                    For Each DLNADevice In DLNADevices
                        If DLNADevice.Key <> "" Then
                            If GetStringIniFile(DLNADevice.Key, "diDeviceType", "") = "DMS" And GetBooleanIniFile(DLNADevice.Key, "diDeviceIsAdded", False) And GetBooleanIniFile(DLNADevice.Key, "diAdminState", False) Then ' is it a DMR and is it added to HS
                                ' this is a Media server
                                DLNAServerName = GetStringIniFile(DLNADevice.Key, "diGivenName", "")
                                ContentDeviceNameBox.AddItem(DLNAServerName, DLNADevice.Key, DLNADevice.Key = ServerUDN)
                            End If
                        End If
                    Next
                    If ServerUDN <> "" Then
                        ' go get the ServerAPI
                        Try
                            ServerAPI = PIReference.GetAPIByUDN(ServerUDN)
                        Catch ex As Exception
                            Log("Error in GetPagePlugin getting Server API for Player = " & ZoneName & " with UDN = " & ServerUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                End If
            Catch ex As Exception
                Log("Error in GetPagePlugin for Player = " & ZoneName & " retrieving the ServerAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            LblWhatsLoaded = GetStringIniFile("DevicePage", MyZoneUDN & "_loaded", ITEM_TYPE.Root.ToString)



            LblRoot.label = GetStringIniFile("DevicePage", MyZoneUDN & "_root", "")
            LblLevel1.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level1", "")
            LblLevel2.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level2", "")
            LblLevel3.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level3", "")
            LblLevel4.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level4", "")
            LblLevel5.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level5", "")
            LblLevel6.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level6", "")
            LblLevel7.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level7", "")
            LblLevel8.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level8", "")
            LblLevel9.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level9", "")
            LblLevel10.label = GetStringIniFile("DevicePage", MyZoneUDN & "_level10", "")
            ObjectIDRoot = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDroot", "")
            ObjectIDLevel1 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel1", "")
            ObjectIDLevel2 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel2", "")
            ObjectIDLevel3 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel3", "")
            ObjectIDLevel4 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel4", "")
            ObjectIDLevel5 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel5", "")
            ObjectIDLevel6 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel6", "")
            ObjectIDLevel7 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel7", "")
            ObjectIDLevel8 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel8", "")
            ObjectIDLevel9 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel9", "")
            ObjectIDLevel10 = GetStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel10", "")

            ar = Nothing
            arplaylist = Nothing
            If Not ServerAPI Is Nothing Then
                ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 0)
                LoadNavigationBox(False)
            End If
            arplaylist = MusicAPI.GetCurrentPlaylistTracks
            LoadPlayListBox(False)


            UpdateStatus()


            If GenerateHeaderFooter Then
                Me.AddHeader(hs.GetPageHeader(MyPageName, "Player Control", "", "", False, True))
            End If

            stb.Append(clsPageBuilder.FormStart("PlayerConfigform", MyPageName, "post"))

            stb.Append(clsPageBuilder.DivStart(MyPageName, ""))
            Me.RefreshIntervalMilliSeconds = 2000
            stb.Append(Me.AddAjaxHandlerPost("action=updatetime", MyPageName))
            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd) ' ErrorMessage

            stb.Append(clsPageBuilder.DivStart("PlayerPanel", "style='color:#0000FF'"))
            stb.Append("<table border='0' cellpadding='0' cellspacing='0' height='389px'><tr>")
            stb.Append("<td style='height: 389px; width: 860px; background-repeat:no-repeat;background-position:left;' background='" & ImagesPath & "Player-Green.png'>")

            stb.Append("<div style='position: relative'>") '; top: -150px'>")


            stb.Append("<div style='position: absolute; left: 300px; top: -185px'>")
            stb.Append("<img src=" & MusicAPI.PlayerIconURL & " style='height:50px; width:50px'>")
            stb.Append("</div>")
            stb.Append("<div style='position: absolute; left: 360px; top: -190px'>")
            stb.Append("<h1>" & ZoneName & "</h1>")
            stb.Append("</div>")

            stb.Append("<div id='ArtImageDiv' style='position: absolute; left: 55px; top: -130px;'>")
            ArtImageBox.style = "height:180px;width:180px"
            ArtImageBox.imagePathNormal = MyArt
            stb.Append(ArtImageBox.Build)
            stb.Append("</div>")

            stb.Append("<div id='VolumeDiv' style='position: absolute; left: 25px; top: -130px'>")
            If MusicAPI.VolumeIsConfigurable Then
                MyVolume = MusicAPI.Volume
                Dim VolumeSlider As New clsJQuery.jqSlider("VolumeSlider", MusicAPI.MyMinimumVolume, MusicAPI.MyMaximumVolume, MyVolume, clsJQuery.jqSlider.jqSliderOrientation.vertical, 180, MyPageName, False)
                VolumeSlider.toolTip = "Shows volume. Drag to set player volume"
                stb.Append(VolumeSlider.build)
            End If
            stb.Append("</div>")

            stb.Append("<div id='MuteDiv' style='position: absolute; left: 15px; top: 55px'>")
            If MusicAPI.MuteIsConfigurable Then
                ButMute.style = "height:35px;width:35px"
                ButMute.imagePathNormal = ButMuteImagePath
                ButMute.toolTip = "Toggle between mute and unmute"
                stb.Append(ButMute.Build)
            End If
            stb.Append("</div>")

            Dim PlayerPosition As Integer = MusicAPI.PlayerPosition
            Dim AlreadyPlayedSpan As TimeSpan = TimeSpan.FromSeconds(PlayerPosition)
            Dim AlreadyPlayedString As String = FormatMyTimeString(AlreadyPlayedSpan, LblDuration)
            Dim ToPlaySpan As TimeSpan = TimeSpan.FromSeconds(LblDuration - PlayerPosition)
            Dim ToPlayString As String = FormatMyTimeString(ToPlaySpan, LblDuration)
            If LblDuration = 0 Then
                ToPlayString = "0:00"
            End If
            stb.Append("<div id='PositionDiv' style='position: absolute; left: 25px; top: -165px; float: center'>")
            MyPosition = MusicAPI.PlayerPosition
            Dim PositionSlider As New clsJQuery.jqSlider("PositionSlider", 0, LblDuration, MyPosition, clsJQuery.jqSlider.jqSliderOrientation.horizontal, 155, MyPageName, False)
            PositionSlider.toolTip = "Shows track position. Drag to change position"
            stb.Append(AlreadyPlayedString & "&nbsp;&nbsp;" & PositionSlider.build & "&nbsp;&nbsp;-" & ToPlayString)
            stb.Append("</div>")

            stb.Append("<div id='RepeatDiv' style='position: absolute; left: 235px; top: -90px'>")
            ButRepeat.imagePathNormal = ButRepeatImagePath
            ButRepeat.style = "height:40px;width:40px"
            ButRepeat.toolTip = "Toggle between Repeat and no-Repeat"
            stb.Append(ButRepeat.Build)
            stb.Append("</div>")

            stb.Append("<div id='ShuffleDiv' style='position: absolute; left: 235px; top: -40px'>")
            ButShuffle.imagePathNormal = ButShuffleImagePath
            ButShuffle.style = "height:40px;width:40px"
            ButShuffle.toolTip = "Toggle between Shuffled and Ordered (no-Shuffle)"
            stb.Append(ButShuffle.Build)
            stb.Append("</div>")


            stb.Append("<div id='TrackDiv' style='position: absolute; left: 295px; top: -120px'>")

            stb.Append(GenerateTrackInfo)
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 175px; top: 80px'>")
            ButPrev.imagePathNormal = ImagesPath & "player-prev.png"
            ButPrev.style = "height:65px;width:65px"
            ButPrev.toolTip = "Play Previous Track"
            stb.Append(ButPrev.Build)
            stb.Append("</div>")

            stb.Append("<div id='PlayBtnDiv' style='position: absolute; left: 260px; top: 80px'>")
            MyPlayerState = MusicAPI.PlayerState()
            Select Case MyPlayerState
                Case HSPI.player_state_values.Forwarding
                    ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                Case HSPI.player_state_values.Playing
                    ButPlay.imagePathNormal = ImagesPath & "player-pause.png"
                Case HSPI.player_state_values.Rewinding
                    ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                Case Else
                    ButPlay.imagePathNormal = ImagesPath & "player-play.png"
            End Select
            ButPlay.style = "height:95px;width:95px"
            ButPlay.toolTip = "Toggle between Play and Pause state"
            stb.Append(ButPlay.Build)
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 370px; top: 110px'>")
            ButStop.imagePathNormal = ImagesPath & "player-stop.png"
            ButStop.style = "height:65px;width:65px"
            ButStop.toolTip = "Stop Player"
            stb.Append(ButStop.Build)
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 450px; top: 80px'>")
            ButNext.imagePathNormal = ImagesPath & "player-next.png"
            ButNext.style = "height:65px;width:65px"
            ButNext.toolTip = "Play Next Track"
            stb.Append(ButNext.Build)
            stb.Append("</div>")

            If MusicAPI.SpeedIsConfigurable Then
                stb.Append("<div style='position: absolute; left: 105px; top: 80px'>")
                ButRew.imagePathNormal = ImagesPath & "player-rew.png"
                ButRew.style = "height:65px;width:65px"
                ButRew.toolTip = "Rewind. Multiple clicks to rewind faster"
                stb.Append(ButRew.Build)
                stb.Append("</div>")
                stb.Append("<div style='position: absolute; left: 520px; top: 80px'>")
                ButFF.imagePathNormal = ImagesPath & "player-fwd.png"
                ButFF.style = "height:65px;width:65px"
                ButFF.toolTip = "Fast Forward. Multiple clicks to forward faster"
                stb.Append(ButFF.Build)
                stb.Append("</div>")
            End If

            stb.Append(clsPageBuilder.DivEnd) ' Position Relative


            stb.Append("</td></tr></table><br />")
            stb.Append(clsPageBuilder.DivEnd)   ' Player Panel

            stb.Append(clsPageBuilder.DivStart("DescriptionCellDiv", ""))
            stb.Append("<table><tr>")
            stb.Append("<td nowrap='nowrap'></td><td nowrap='nowrap'></td>")
            stb.Append("</tr></table>")
            stb.Append(clsPageBuilder.DivEnd)

            stb.Append(clsPageBuilder.DivStart("NavigationInfoDiv", ""))
            stb.Append("<table><tr><td>")
            stb.Append("Select your Server Device &nbsp;" & ContentDeviceNameBox.Build & "</td><td>")
            stb.Append(clsPageBuilder.DivStart("PlaylistSelectorDiv", ""))
            stb.Append("Select your Playlist &nbsp;" & PlayListSelector.Build)
            stb.Append(clsPageBuilder.DivEnd) ' PlaylistSelectorDiv
            stb.Append("</td></tr><tr><td></td><td>")
            stb.Append(ButSavePlaylist.Build & SavePlayListAsNameBox.Build)
            stb.Append("</td></tr><tr><td colspan='2'>")
            stb.Append(clsPageBuilder.DivStart("LabelDiv", "style='color:#0000FF; text-align: middle'"))
            stb.Append(LblRoot.Build & ">" & LblLevel1.Build & ">" & LblLevel2.Build & ">" & LblLevel3.Build & ">" & LblLevel4.Build & ">" & LblLevel5.Build & ">" & LblLevel6.Build & ">" & LblLevel7.Build & ">" & LblLevel8.Build & ">" & LblLevel9.Build & ">" & LblLevel10.Build)
            stb.Append(clsPageBuilder.DivEnd) ' LabelDiv
            stb.Append("</td></tr>")

            stb.Append("<tr><td nowrap='nowrap'>")
            NavigationBox.submitForm = False
            'NavigationBox.style = "height:380px;width:370px;overflow:auto"
            NavigationBox.style = "height:380px;width:370px"
            NavigationBox.height = 380
            NavigationBox.width = 370
            NavigationBox.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
            'NavigationBox.UseBothClickEvents = True
            'NavigationBox.WaitTime = 200


            'NavigationBox.functionToCallOnDoubleClick 
            PlaylistBox.submitForm = False
            'PlaylistBox.style = "height:380px;width:370px;overflow:auto"
            PlaylistBox.style = "height:380px;width:370px"
            PlaylistBox.height = 380
            PlaylistBox.width = 370
            stb.Append(clsPageBuilder.DivStart("NavigationDiv", ""))
            NavigationBox.toolTip = "Navigationbox: Click to navigate into an item. Once at the bottom, use the 'Add to Playlist' button"
            stb.Append(NavigationBox.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td><td nowrap='nowrap'>")
            stb.Append(clsPageBuilder.DivStart("PlaylistDiv", ""))
            PlaylistBox.toolTip = "Playlistbox: Click to play item"
            stb.Append(PlaylistBox.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td></tr><tr><td align='left'>")
            stb.Append(clsPageBuilder.DivStart("ButAddToPlaylistDiv", ""))
            ButAddToPlaylist.toolTip = "The item selected in the Navigationbox will be added to the playlist when clicked"
            stb.Append(ButAddToPlaylist.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td><td align='left'>")
            stb.Append(clsPageBuilder.DivStart("ButClearListDiv", ""))
            ButClearList.toolTip = "Clears the entire Playlist when clicked"
            stb.Append(ButClearList.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td></tr></table>")
            stb.Append(clsPageBuilder.DivEnd) ' NavigationPanel



            stb.Append(clsPageBuilder.DivEnd) ' page end
            stb.Append(clsPageBuilder.FormEnd)

            InitDone = True

            If GenerateHeaderFooter Then
                ' add the body html to the page
                Me.AddBody(stb.ToString)
                Me.AddFooter(hs.GetPageFooter)
                Me.suppressDefaultFooter = True
                ' return the full page
                Return Me.BuildPage()
            End If


        Catch ex As Exception
            Log("Error in GetPagePlugin for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return stb.ToString

    End Function

    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        'If g_bDebug Then Log("PostBackProc for PlayerControl called for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(System.Web.HttpUtility.HtmlDecode(data))
        'parts = HttpUtility.ParseQueryString(data)

        If parts IsNot Nothing Then
            'Log("PostBackProc for PlayerControl called  for Player = " & ZoneName & " with part = '" & parts.GetKey(0).ToUpper.ToString & "'", LogType.LOG_TYPE_INFO)
            If (data.ToString.ToUpper <> "ACTION=UPDATETIME") And (g_bDebug = True) Then Log("PostBackProc for PlayerControl called for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
            Dim DoubleClickFlag As Boolean = False
            Try
                If parts.Item("click") = "double" Then
                    If g_bDebug Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found double click key", LogType.LOG_TYPE_WARNING)
                    DoubleClickFlag = True
                End If
            Catch ex As Exception
                If g_bDebug Then Log("Error in postBackProc for PlayerControl for Player = " & ZoneName & " searching for double click key with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(HttpUtility.UrlDecode(Part), "_")
                        If (data.ToString.ToUpper <> "ACTION=UPDATETIME") Then
                            If g_bDebug Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                            If g_bDebug Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                        End If
                        Dim ObjectValue As String = HttpUtility.UrlDecode(parts(Part))
                        'Dim ObjectValue As String = parts(Part)

                        Select Case ObjectNameParts(0).ToString.ToUpper
                            Case "INSTANCE"
                            Case "REF"
                            Case "PLUGIN"
                                If ObjectValue <> sIFACE_NAME Then
                                    Log("Error in postBackProc for Zoneplayer = " & ZoneName & " and page = " & page.ToString & ", data = " & data & ", user = " & user & ", userRights = " & userRights.ToString, LogType.LOG_TYPE_ERROR)
                                    Exit For
                                End If
                            Case "ACTION"
                                If ObjectValue = "updatetime" Then
                                    CheckForChanges()
                                End If
                            Case "ID"
                            Case "BUTPLAY"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TogglePlay()
                                    If g_bDebug Then Log("postBackProc issued Toggle Play command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTPREV"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackPrev()
                                    If g_bDebug Then Log("postBackProc issued Previous command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTSTOP"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.StopPlay()
                                    If g_bDebug Then Log("postBackProc issued Stop command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTPAUSE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.Pause()
                                    If g_bDebug Then Log("postBackProc issued Pause command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTNEXT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackNext()
                                    If g_bDebug Then Log("postBackProc issued Next command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTREW"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.DoRewind()
                                    If g_bDebug Then Log("postBackProc issued Rew command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTFF"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.DoFastForward()
                                    If g_bDebug Then Log("postBackProc issued FF command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTMUTE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleMuteState("Master")
                                    If g_bDebug Then Log("postBackProc issued Mute command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTREPEAT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleQueueRepeat()
                                    If g_bDebug Then Log("postBackProc issued Repeat command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTSHUFFLE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleQueueShuffle()
                                    If g_bDebug Then Log("postBackProc issued Shuffle command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "VOLUMESLIDER"
                                If g_bDebug Then Log("postBackProc issued Volume command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.Volume = Val(ObjectValue)
                            Case "POSITIONSLIDER"
                                If g_bDebug Then Log("postBackProc issued Position command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.PlayerPosition = Val(ObjectValue)
                            Case "BUTLOUDNESS"
                                If g_bDebug Then Log("postBackProc issued Loudness command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.ToggleLoudnessState("Master")
                            Case "BUTADDTOPLAYLIST"
                                If g_bDebug Then Log("postBackProc issued AddToPlaylist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                AddTrack_Click()
                            Case "BUTCLEARLIST"
                                If g_bDebug Then Log("postBackProc issued ClearList command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButClearList_Click()
                            Case "BUTSAVEPLAYLIST"
                                If g_bDebug Then Log("postBackProc issued ButSavePlaylist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                SavePlayListAs_Click()
                            Case "BUTEDITOR"
                                If g_bDebug Then Log("postBackProc issued Editor command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case "PLAYLISTBOX"
                                If g_bDebug Then Log("postBackProc issued Playlistbox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                PlaylistBox_SelectedIndexChanged(ObjectValue.ToString)
                            Case "NAVIGATIONBOX"
                                If g_bDebug Then Log("postBackProc issued NavigationBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                NavigationBox_SelectedIndexChanged(ObjectValue.ToString, DoubleClickFlag)
                            Case "CONTENTDEVICENAMEBOX"
                                If g_bDebug Then Log("postBackProc issued ContentDeviceNameBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue.ToString = "" Then Exit Select
                                If ObjectValue.ToString.ToUpper = "NO SELECTION" Then
                                    Try
                                        ClearSelections()
                                        LblWhatsLoaded = ITEM_TYPE.Root.ToString
                                        UpdateIniFile()
                                        NavigationBox.items.Clear()
                                        ar = Nothing
                                        ServerAPI = Nothing
                                        LoadNavigationBox(True)
                                        WriteStringIniFile("DevicePage", MyZoneUDN & "_ServerUDN", "")
                                    Catch ex As Exception
                                    End Try
                                    Exit Select
                                End If
                                Try
                                    ServerAPI = PIReference.GetAPIByUDN(ObjectValue.ToString)
                                    If ServerAPI Is Nothing Then
                                        Log("Error getting Server API for UDN = " & ObjectValue.ToString, LogType.LOG_TYPE_ERROR)
                                        Exit Select
                                    End If
                                    If g_bDebug Then Log("postBackProc for Zoneplayer = " & ZoneName & " is writing = " & ServerAPI.DeviceUDN & " ServerUDN to ini file", LogType.LOG_TYPE_INFO)
                                    If ServerAPI.DeviceUDN <> GetStringIniFile("DevicePage", MyZoneUDN & "_ServerUDN", "") Then ' only update when value has changed
                                        WriteStringIniFile("DevicePage", MyZoneUDN & "_ServerUDN", ServerAPI.DeviceUDN)
                                        Try
                                            ClearSelections()
                                            LblWhatsLoaded = ITEM_TYPE.Root.ToString
                                            UpdateIniFile()
                                            NavigationBox.items.Clear()
                                            ar = ServerAPI.GetContainerFromServer("0", False)
                                            If Not ar Is Nothing Then
                                                LoadNavigationBox(True)
                                            End If
                                        Catch ex As Exception
                                            Log("Error in postBackProc for Zoneplayer = " & ZoneName & " loading the NavigationBox with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                            Exit Select
                                        End Try
                                    End If
                                Catch ex As Exception
                                    Log("Error in postBackProc for Zoneplayer = " & ZoneName & " for ContentDeviceNameBox change with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLROOT"
                                If g_bDebug Then Log("postBackProc issued LblRoot command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    ar = ServerAPI.GetContainerFromServer("0", False)
                                    LblWhatsLoaded = ITEM_TYPE.Root.ToString
                                    LblLevel1.label = ""
                                    LblLevel2.label = ""
                                    LblLevel3.label = ""
                                    LblLevel4.label = ""
                                    LblLevel5.label = ""
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDRoot = "0"
                                    ObjectIDLevel1 = ""
                                    ObjectIDLevel2 = ""
                                    ObjectIDLevel3 = ""
                                    ObjectIDLevel4 = ""
                                    ObjectIDLevel5 = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LoadNavigationBox()
                                    UpdateIniFile()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " LblRoot_Click with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL1"
                                If g_bDebug Then Log("postBackProc issued LblLevel1 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblWhatsLoaded = ITEM_TYPE.Level1.ToString
                                    LblLevel2.label = ""
                                    LblLevel3.label = ""
                                    LblLevel4.label = ""
                                    LblLevel5.label = ""
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel2 = ""
                                    ObjectIDLevel3 = ""
                                    ObjectIDLevel4 = ""
                                    ObjectIDLevel5 = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 1)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel1 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL2"
                                If g_bDebug Then Log("postBackProc issued LblLevel2 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel3.label = ""
                                    LblLevel4.label = ""
                                    LblLevel5.label = ""
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel3 = ""
                                    ObjectIDLevel4 = ""
                                    ObjectIDLevel5 = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level2.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 2)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel2 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL3"
                                If g_bDebug Then Log("postBackProc issued LblLevel3 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel4.label = ""
                                    LblLevel5.label = ""
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel4 = ""
                                    ObjectIDLevel5 = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level3.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 3)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel3 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL4"
                                If g_bDebug Then Log("postBackProc issued LblLevel4 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel5.label = ""
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel5 = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level4.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 4)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel4 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL5"
                                If g_bDebug Then Log("postBackProc issued LblLevel5 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel6.label = ""
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel6 = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level5.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 5)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel5 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL6"
                                If g_bDebug Then Log("postBackProc issued LblLevel6 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel7.label = ""
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel7 = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level6.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 6)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel6 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL7"
                                If g_bDebug Then Log("postBackProc issued LblLevel7 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel8.label = ""
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel8 = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level7.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 7)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel7 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL8"
                                If g_bDebug Then Log("postBackProc issued LblLevel8 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel9.label = ""
                                    LblLevel10.label = ""
                                    ObjectIDLevel9 = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level8.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 8)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel8 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL9"
                                If g_bDebug Then Log("postBackProc issued LblLevel9 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue = "" Or ServerAPI Is Nothing Then Exit Select
                                Try
                                    LblLevel10.label = ""
                                    ObjectIDLevel10 = ""
                                    LblWhatsLoaded = ITEM_TYPE.Level9.ToString
                                    UpdateIniFile()
                                    ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, 9)
                                    LoadNavigationBox()
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " at LblLevel9 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LBLLEVEL10"
                                If g_bDebug Then Log("postBackProc issued LblLevel10 command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case "CLICK"
                                If g_bDebug Then Log("postBackProc issued click command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case "PLAYLISTSELECTOR"
                                If g_bDebug Then Log("postBackProc issued PlaylistSelector command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue.ToString = "" Then Exit Select
                                WriteStringIniFile("DevicePage", "playlistSelector_" & MyZoneUDN, ObjectValue.ToString)
                                MyLastSelectedPlaylistName = ObjectValue.ToString
                                Try
                                    MusicAPI.LoadCurrentPlaylistTracks(MyLastSelectedPlaylistName)
                                    arplaylist = MusicAPI.GetCurrentPlaylistTracks
                                    LoadPlayListBox(True)
                                Catch ex As Exception
                                    Log("Error in postBackProc for Player = " & ZoneName & " unable to load Playlist with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "SAVEPLAYLISTASNAMEBOX"
                                SavePlayListAsNameBox.defaultText = ObjectValue.ToString
                            Case Else
                                If g_bDebug Then Log("postBackProc for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_WARNING)
                                If g_bDebug Then Log("postBackProc for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_WARNING)
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in postBackProc for Player = " & ZoneName & " processing page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If g_bDebug Then Log("postBackProc for Player = " & ZoneName & " found parts to be empty", LogType.LOG_TYPE_INFO)
        End If
        Return MyBase.postBackProc(page, data, user, userRights)
    End Function

    Private Function GenerateTrackInfo() As String
        GenerateTrackInfo = ""
        Dim NbrOfLines As Integer = 0
        If LblTrackName <> "" Then
            NbrOfLines += 1
            GenerateTrackInfo &= "Title:  " & LblTrackName & "<br>"
        End If
        If LblArtistName <> "" Then
            Dim Lines As String() = Split(MyArtist, vbCrLf)
            If Lines IsNot Nothing Then
                If Lines.Count > 0 Then
                    For Each Line As String In Lines
                        NbrOfLines += 1
                        If NbrOfLines > 10 Then Exit Function
                        GenerateTrackInfo &= "Artist:  " & Line & "<br>"
                    Next
                End If
            End If
        End If
        If LblAlbumName <> "" Then
            NbrOfLines += 1
            If NbrOfLines > 10 Then Exit Function
            GenerateTrackInfo &= "Album:  " & LblAlbumName & "<br>"
        End If

        If LblDescr <> "" Then
            Dim Lines As String() = Split(MyTrackDesc, vbCrLf)
            If Lines IsNot Nothing Then
                If Lines.Count > 0 Then
                    For Each Line As String In Lines
                        NbrOfLines += 1
                        If NbrOfLines > 10 Then Exit Function
                        GenerateTrackInfo &= "Descr:  " & Line & "<br>"
                    Next
                End If
            End If
        End If
        If LblNextTrackName <> "" Then
            NbrOfLines += 1
            If NbrOfLines > 10 Then Exit Function
            GenerateTrackInfo &= "Next Title:  " & LblNextTrackName & "<br>"
        End If

        If LblNextArtistName <> "" Then
            NbrOfLines += 1
            If NbrOfLines > 10 Then Exit Function
            GenerateTrackInfo &= "Next Artist: " & LblNextArtistName & "<br>"
        End If

        If LblNextAlbumName <> "" Then
            NbrOfLines += 1
            If NbrOfLines > 10 Then Exit Function
            GenerateTrackInfo &= "Next Album:  " & LblNextAlbumName & "<br>"
        End If

        If LblNextDescr <> "" Then
            Dim Lines As String() = Split(MyNextTrackDesc, vbCrLf)
            If Lines IsNot Nothing Then
                If Lines.Count > 0 Then
                    For Each Line As String In Lines
                        NbrOfLines += 1
                        If NbrOfLines > 10 Then Exit Function
                        GenerateTrackInfo &= "Next Descr:  " & Line & "<br>"
                    Next
                End If
            End If
        End If
    End Function

    Private Sub LoadPlayListSelector(SelectedValue As String)
        If g_bDebug Then Log("LoadPlayListSelector called for Player = " & ZoneName & " with Value = " & SelectedValue, LogType.LOG_TYPE_INFO)
        Dim Index As Integer = 0
        PlayListSelector.items.Clear()
        PlayListSelector.AddItem("No Selection", "", SelectedValue = "")
        Dim PLS As Object = Nothing
        Try
            PLS = PIReference.GetPlayLists()
        Catch ex As Exception
            Log("Error in LoadPlayListSelector for Player = " & ZoneName & " getting PlayList with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If PLS IsNot Nothing Then
                Dim MaxIndex As Integer = UBound(PLS)
                If MaxIndex > 10000 Then
                    Log("Warning in LoadPlayListSelector  for Player = " & ZoneName & "getting PlayList. Too many entried. Entries found = " & MaxIndex.ToString & ". Set to 10,000", LogType.LOG_TYPE_WARNING)
                    MaxIndex = 10000
                End If
                For Index = 0 To MaxIndex
                    Dim Items As String()
                    Items = Split(PLS(Index), ":;:-:")
                    PlayListSelector.AddItem(Items(0), Items(0), SelectedValue = Items(0).ToString)
                    If SelectedValue = Items(0).ToString Then
                        SavePlayListAsNameBox.defaultText = SelectedValue
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in LoadPlayListSelector for Player = " & ZoneName & "reading Playlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        LoadPlayListBox(True)
    End Sub

    Private Sub LoadNavigationBox(Optional GenerateDiv As Boolean = True)
        If g_bDebug Then Log("LoadNavigationBox called for Player = " & ZoneName & " and GenerateDiv= " & GenerateDiv.ToString & " and InitDone = " & InitDone.ToString, LogType.LOG_TYPE_INFO)
        DescriptionCell = ""
        Try
            ButAddToPlaylist.enabled = False
            NavigationBox.items.Clear()
            NavigationBox.selectedItemIndex = -1
            Dim MaxLoad As Integer = 0
            If ar IsNot Nothing Then
                If g_bDebug Then Log("LoadNavigationBox called for Player = " & ZoneName & " and has " & (UBound(ar) + 1).ToString & " entries in the NavigationBox", LogType.LOG_TYPE_INFO)
                For Index = 0 To UBound(ar)
                    NavigationBox.AddItem(ar(Index).Title.ToString, HttpUtility.UrlEncode(ar(Index).Title.ToString & ":;:-:" & ar(Index).Id.ToString & ":;:-:" & ar(Index).ClassType.ToString & ":;:-:" & ar(Index).ItemOrContainer.ToString), False) ' & ":;:-:" & ar(Index).ClassType & ":;:-:" & ar(Index).IconURL & ":;:-:" & ar(Index).ArtistName & ":;:-:" & ar(Index).AlbumName & ":;:-:" & ar(Index).Genre), False)
                    If ar(Index).ItemOrContainer.ToString <> "CONTAINER" Then
                        ButAddToPlaylist.enabled = True
                    End If
                Next
            End If
            If InitDone Then
                If GenerateDiv Then Me.divToUpdate.Add("NavigationDiv", NavigationBox.Build)
                If GenerateDiv Then Me.divToUpdate.Add("LabelDiv", LblRoot.Build & ">" & LblLevel1.Build & ">" & LblLevel2.Build & ">" & LblLevel3.Build & ">" & LblLevel4.Build & ">" & LblLevel5.Build & ">" & LblLevel6.Build & ">" & LblLevel7.Build & ">" & LblLevel8.Build & ">" & LblLevel9.Build & ">" & LblLevel10.Build)
                Me.divToUpdate.Add("DescriptionCellDiv", "<table><tr>" & DescriptionCell & "</tr></table>")
                Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
            End If
        Catch ex As Exception
            Log("Error in LoadNavigationBox for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub LoadPlayListBox(Optional GenerateDiv As Boolean = False)
        If g_bDebug Then Log("LoadPlayListBox called for Player = " & ZoneName & "and GenerateDiv = " & GenerateDiv.ToString, LogType.LOG_TYPE_INFO)
        Try
            PlaylistBox.items.Clear()
            If arplaylist IsNot Nothing Then
                If g_bDebug Then Log("LoadPlayListBox called for Player = " & ZoneName & " and has " & (UBound(arplaylist) + 1).ToString & " entries in the PlaylistBox", LogType.LOG_TYPE_INFO)
                For Index = 0 To UBound(arplaylist)
                    Dim QItems As String() = Split(arplaylist(Index), ":;:-:")
                    PlaylistBox.AddItem(QItems(0).ToString, (Index + 1).ToString, QItems(0).ToString = MyTrack) 'HttpUtility.UrlEncode(QItems(0).ToString & ":;:-:" & QItems(1).ToString & ":;:-:" & QItems(2).ToString & ":;:-:" & QItems(3).ToString & ":;:-:" & QItems(4).ToString & ":;:-:" & QItems(5).ToString & ":;:-:" & QItems(6).ToString), QItems(0).ToString = MyTrack)
                Next
            End If
        Catch ex As Exception
            Log("Error in LoadPlayListBox for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If GenerateDiv And InitDone Then Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
    End Sub

    Private Sub UpdateStatus()
        Try
            Dim st As String
            st = MusicAPI.CurrentTrack
            MyTrack = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblTrackName = st

            st = MusicAPI.CurrentArtist
            MyArtist = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblArtistName = st

            st = MusicAPI.CurrentAlbum
            MyAlbum = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblAlbumName = st

            st = MusicAPI.NextTrack
            MyNextTrack = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblNextTrackName = st

            st = MusicAPI.NextArtist
            MyNextArtist = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblNextArtistName = st

            st = MusicAPI.NextAlbum
            MyNextAlbum = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblNextAlbumName = st

            LblDuration = Val(MusicAPI.CurrentTrackDuration)

            st = MusicAPI.CurrentTrackDescription
            MyTrackDesc = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblDescr = st

            st = MusicAPI.NextTrackDescription
            MyNextTrackDesc = st
            If st.Length > 86 Then st = st.Substring(0, 74) & "..."
            LblNextDescr = st

            MyArt = MusicAPI.CurrentAlbumArtPath.ToString
            If g_bDebug Then Log("UpdateStatus for PlayerControl for player = " & MusicAPI.DeviceName & " has Queueshuffle state = " & MusicAPI.QueueShuffleStatus.ToLower, LogType.LOG_TYPE_INFO)
            MyShuffle = False
            Select Case MusicAPI.QueueShuffleStatus.ToLower
                Case "shuffled"
                    ButShuffleImagePath = ImagesPath & "shuffle.png"
                    MyShuffle = True
                Case "ordered"
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                Case "sorted"
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                Case Else
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
            End Select
            If g_bDebug Then Log("UpdateStatus for PlayerControl for player = " & MusicAPI.DeviceName & " has Queuerepeat state = " & MusicAPI.QueueRepeat.ToString, LogType.LOG_TYPE_INFO)
            MyRepeat = False
            Select Case MusicAPI.QueueRepeat
                Case repeat_modes.repeat_all
                    ButRepeatImagePath = ImagesPath & "Repeat.png"
                    MyRepeat = True
                Case repeat_modes.repeat_off
                    ButRepeatImagePath = ImagesPath & "NoRepeat.png"
                Case repeat_modes.repeat_one
                    ButRepeatImagePath = ImagesPath & "Repeat.png"
                    MyRepeat = True
            End Select
            If MusicAPI.PlayerMute = True Then
                MyMute = True
                ButMuteImagePath = ImagesPath & "Muted.png"
            Else
                MyMute = False
                ButMuteImagePath = ImagesPath & "UnMuted.png"
            End If
        Catch ex As Exception
            Log("Error in UpdateStatus for PlayerControl with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub


    Private Sub CheckForChanges()
        'If g_bDebug Then Log("PlayerControl.CheckForChanges for Zoneplayer = " & ZoneName & " called and PlayerState = " & MusicAPI.PlayerState.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim trackHasChanged As Boolean = False
            Try
                If LblDuration <> Val(MusicAPI.CurrentTrackDuration) Then
                    LblDuration = Val(MusicAPI.CurrentTrackDuration)
                    trackHasChanged = True
                End If
            Catch ex As Exception
                LblDuration = 0
            End Try
            If MyPlayerState <> MusicAPI.PlayerState Then
                MyPlayerState = MusicAPI.PlayerState
                Select Case MyPlayerState
                    Case HSPI.player_state_values.Playing
                        ButPlay.imagePathNormal = ImagesPath & "player-pause.png"
                    Case HSPI.player_state_values.Forwarding
                        ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                    Case HSPI.player_state_values.Rewinding
                        ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                    Case Else
                        ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                End Select
                ButPlay.style = "height:95px;width:95px"
                Me.divToUpdate.Add("PlayBtnDiv", ButPlay.Build)
            End If
            If (MusicAPI.PlayerState = HSPI.player_state_values.Playing) Or (MusicAPI.PlayerState = HSPI.player_state_values.Forwarding) Or (MusicAPI.PlayerState = HSPI.player_state_values.Rewinding) Or trackHasChanged Then
                Dim PlayerPosition As Integer = MusicAPI.PlayerPosition
                Dim AlreadyPlayedSpan As TimeSpan = TimeSpan.FromSeconds(PlayerPosition)
                Dim AlreadyPlayedString As String = FormatMyTimeString(AlreadyPlayedSpan, LblDuration)
                Dim ToPlaySpan As TimeSpan = TimeSpan.FromSeconds(LblDuration - PlayerPosition)
                Dim ToPlayString As String = FormatMyTimeString(ToPlaySpan, LblDuration)
                If LblDuration = 0 Then
                    ToPlayString = "0:00"
                End If
                Dim PositionSlider As New clsJQuery.jqSlider("PositionSlider", 0, LblDuration, PlayerPosition, clsJQuery.jqSlider.jqSliderOrientation.horizontal, 155, MyPageName, False)
                PositionSlider.toolTip = "Shows track position. Drag to change position"
                Me.divToUpdate.Add("PositionDiv", AlreadyPlayedString & "&nbsp;&nbsp;" & PositionSlider.build & "&nbsp;&nbsp;-" & ToPlayString)
            End If
            If MusicAPI.VolumeIsConfigurable Then
                If MusicAPI.Volume <> MyVolume Then
                    MyVolume = MusicAPI.Volume
                    Dim VolumeSlider As New clsJQuery.jqSlider("VolumeSlider", MusicAPI.MyMinimumVolume, MusicAPI.MyMaximumVolume, MyVolume, clsJQuery.jqSlider.jqSliderOrientation.vertical, 180, MyPageName, False)
                    VolumeSlider.toolTip = "Shows volume. Drag to set player volume"
                    Me.divToUpdate.Add("VolumeDiv", VolumeSlider.build)
                End If
            End If
            If MusicAPI.CurrentTrack <> MyTrack Then
                MyTrack = MusicAPI.CurrentTrack
                Dim st As String
                st = MyTrack
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblTrackName = st
                trackHasChanged = True
                LoadPlayListBox(True)
                Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
            End If
            If MusicAPI.CurrentAlbum <> MyAlbum Then
                MyAlbum = MusicAPI.CurrentAlbum
                Dim st As String
                st = MyAlbum
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblAlbumName = st
                trackHasChanged = True
            End If
            If MusicAPI.CurrentArtist <> MyArtist Then
                MyArtist = MusicAPI.CurrentArtist
                Dim st As String
                st = MyArtist
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblArtistName = st
                trackHasChanged = True
            End If
            If MusicAPI.CurrentAlbumArtPath.ToString <> MyArt Then
                MyArt = MusicAPI.CurrentAlbumArtPath.ToString
                ArtImageBox.imagePathNormal = MyArt
                ArtImageBox.style = "height:180px;width:180px"
                Me.divToUpdate.Add("ArtImageDiv", ArtImageBox.Build)
            End If
            If MusicAPI.NextTrack <> MyNextTrack Then
                MyNextTrack = MusicAPI.NextTrack
                Dim st As String
                st = MyNextTrack
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblNextTrackName = st
                trackHasChanged = True
            End If
            If MusicAPI.NextAlbum <> MyNextAlbum Then
                MyNextAlbum = MusicAPI.NextAlbum
                Dim st As String
                st = MyNextAlbum
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblNextAlbumName = st
                trackHasChanged = True
            End If
            If MusicAPI.NextArtist <> MyNextArtist Then
                MyNextArtist = MusicAPI.NextArtist
                Dim st As String
                st = MyNextArtist
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblNextArtistName = st
                trackHasChanged = True
            End If
            If MusicAPI.CurrentTrackDescription <> MyTrackDesc Then
                MyTrackDesc = MusicAPI.CurrentTrackDescription
                Dim st As String
                st = MyTrackDesc
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblDescr = st
                trackHasChanged = True
            End If
            If MusicAPI.NextTrackDescription <> MyNextTrackDesc Then
                MyNextTrackDesc = MusicAPI.NextTrackDescription
                Dim st As String
                st = MyNextTrackDesc
                If st.Length > 86 Then st = st.Substring(0, 74) & "..."
                LblNextDescr = st
                trackHasChanged = True
            End If
            If MusicAPI.MuteIsConfigurable Then
                If MyMute <> MusicAPI.PlayerMute Then
                    MyMute = MusicAPI.PlayerMute
                    If MyMute Then
                        ButMute.imagePathNormal = ImagesPath & "Muted.png"
                    Else
                        ButMute.imagePathNormal = ImagesPath & "UnMuted.png"
                    End If
                    ButMute.style = "height:35px;width:35px"
                    Me.divToUpdate.Add("MuteDiv", ButMute.Build)
                End If
            End If
            Dim ShuffleChanged As Boolean = False
            Select Case MusicAPI.QueueShuffleStatus.ToLower
                Case "shuffled"
                    ButShuffleImagePath = ImagesPath & "shuffle.png"
                    If Not MyShuffle Then
                        ShuffleChanged = True
                        MyShuffle = True
                    End If
                Case Else
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                    If MyShuffle Then
                        ShuffleChanged = True
                        MyShuffle = False
                    End If
            End Select
            ButShuffle.style = "height:40px;width:40px"
            ButShuffle.imagePathNormal = ButShuffleImagePath
            If ShuffleChanged Then Me.divToUpdate.Add("ShuffleDiv", ButShuffle.Build)
            Dim RepeatChanged As Boolean = False
            Select Case MusicAPI.QueueRepeat
                Case repeat_modes.repeat_all, repeat_modes.repeat_one
                    ButRepeatImagePath = ImagesPath & "Repeat.png"
                    If Not MyRepeat Then
                        RepeatChanged = True
                        MyRepeat = True
                    End If
                Case repeat_modes.repeat_off
                    ButRepeatImagePath = ImagesPath & "NoRepeat.png"
                    If MyRepeat Then
                        RepeatChanged = True
                        MyRepeat = False
                    End If
            End Select
            ButRepeat.style = "height:40px;width:40px"
            ButRepeat.imagePathNormal = ButRepeatImagePath
            If RepeatChanged Then Me.divToUpdate.Add("RepeatDiv", ButRepeat.Build)
            If trackHasChanged Then Me.divToUpdate.Add("TrackDiv", GenerateTrackInfo)
            If MusicAPI.MyQueueHasChanged Then
                MusicAPI.MyQueueHasChanged = False
                arplaylist = MusicAPI.GetCurrentPlaylistTracks
                LoadPlayListBox(True)
            End If
        Catch ex As Exception
            Log("Error in CheckForChanges with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function FormatMyTimeString(AlreadyPlayedSpan As TimeSpan, LblDuration As Integer) As String
        If LblDuration >= (24 * 60 * 10) Then
            ' more than one hour. so report in hour format
            FormatMyTimeString = Format(AlreadyPlayedSpan.Hours, "00") & ":" & Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        ElseIf LblDuration >= (24 * 60) Then
            FormatMyTimeString = Format(AlreadyPlayedSpan.Hours, "0") & ":" & Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        ElseIf LblDuration >= (10 * 60) Then
            FormatMyTimeString = Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        Else
            FormatMyTimeString = Format(AlreadyPlayedSpan.Minutes, "0") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        End If
    End Function



    Protected Sub ButPlay_Click()
        Try
            If MusicAPI.PlayerState = player_state_values.paused Then
                'music was paused so resume playing
                MusicAPI.PlayIfPaused()
            ElseIf MusicAPI.PlayerState = player_state_values.playing Then
                'music was playing so pause it
                MusicAPI.Pause()
            Else
                'nothing was playing, but they hit play.
                'if the current playlist has tracks in it, play that list
                Dim tracks() As String = MusicAPI.GetCurrentPlaylistTracks
                If UBound(tracks) = 0 And tracks(0) = "" Then
                    'MusicAPI.PlayMusic(LblArtist.label, LblAlbum.label, LblPlaylist.label, LblGenre.label, "", "", "", ListBox1.SelectedValue, "", LblAudiobooks.label, LblPodcasts.label) ' changed for SonosController
                    arplaylist = MusicAPI.GetCurrentPlaylistTracks
                    ar = Nothing
                    LoadNavigationBox()
                Else
                    ' build a playlist with selected tracks and play
                    MusicAPI.Play()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub


    Protected Sub AddTrack_Click()
        ' add selected track to existing playlist
        If g_bDebug Then Log("AddTrack_Click called for Player = " & ZoneName & " and MyLastSelecteNavBoxItems = " & MyLastSelecteNavBoxItems.ToString & " and MyLastSelectedNavBoxClass = " & MyLastSelectedNavBoxClass.ToString, LogType.LOG_TYPE_INFO)
        If MyLastSelecteNavBoxItems = "" And MyLastSelectedNavBoxClass = "" Then Exit Sub
        Try
            If ServerAPI IsNot Nothing Then
                MusicAPI.AddTrackToCurrentPlaylist(MyLastSelecteNavBoxItems, ServerAPI.DeviceUDN, HSPI.QueueActions.qaPlayLast)
                arplaylist = MusicAPI.GetCurrentPlaylistTracks
                LoadPlayListBox(True)
                MusicAPI.SaveCurrentPlaylistTracks("")
            End If
            UpdateStatus()
        Catch ex As Exception
            Log("Error in AddTrack_Click for Player = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Protected Sub ButClearList_Click()
        If g_bDebug Then Log("ButClearList_Click called for Player = " & ZoneName, LogType.LOG_TYPE_INFO)
        MusicAPI.StopPlay()
        MusicAPI.ClearCurrentPlayList()
        MusicAPI.SaveCurrentPlaylistTracks("")
        arplaylist = MusicAPI.GetCurrentPlaylistTracks
        LoadPlayListBox(True)
        UpdateStatus()
    End Sub

    Protected Sub SavePlayListAs_Click()
        If g_bDebug Then Log("SavePlayListAs_Click called for Player = " & ZoneName & " and PlaylistName = " & SavePlayListAsNameBox.defaultText, LogType.LOG_TYPE_INFO)
        Try
            MusicAPI.SaveCurrentPlaylistTracks(SavePlayListAsNameBox.defaultText)
            MyLastSelectedPlaylistName = SavePlayListAsNameBox.defaultText
            WriteStringIniFile("DevicePage", "playlistSelector_" & MyZoneUDN, SavePlayListAsNameBox.defaultText)
            LoadPlayListSelector(SavePlayListAsNameBox.defaultText)
            Me.divToUpdate.Add("PlaylistSelectorDiv", "Select your Playlist &nbsp;" & PlayListSelector.Build)
        Catch ex As Exception
            Log("Error in SavePlaylistBtn_Click for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Protected Sub PlaylistBox_SelectedIndexChanged(value As String)
        If g_bDebug Then Log("PlaylistBox_SelectedIndexChanged called for Player = " & ZoneName & " with Value = " & value, LogType.LOG_TYPE_INFO)
        If value = "" Then Exit Sub
        Try
            MusicAPI.SkipToTrack(CInt(Val(value)))
            arplaylist = MusicAPI.GetCurrentPlaylistTracks
            Exit Sub
        Catch ex As Exception
            Log("Error in ListBox1_SelectedIndexChanged with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        LoadNavigationBox()
        Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
    End Sub

    Private Sub AddAjaxDivForNavBox()
        If g_bDebug Then Log("AddAjaxDivForNavBox called for Player = " & ZoneName, LogType.LOG_TYPE_INFO)
        Me.divToUpdate.Add("NavigationDiv", NavigationBox.Build)
        'Me.divToUpdate.Add("LabelDiv", LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
    End Sub

    Protected Sub NavigationBox_SelectedIndexChanged(Value As String, DoubleClick As Boolean)
        If g_bDebug Then Log("NavigationBox_SelectedIndexChanged called for Player = " & ZoneName & " with Value = " & Value & ", Double Click = " & DoubleClick.ToString & " and WhatsLoaded = " & LblWhatsLoaded, LogType.LOG_TYPE_INFO)
        If Value = "" Then Exit Sub
        MyLastSelecteNavBoxItems = Value
        MyLastSelectedNavBoxClass = LblWhatsLoaded

        Try
            Dim ValueString As String() = Split(Value, ":;:-:")
            Dim ObjectName As String = ValueString(0)
            Dim ObjectID As String = ValueString(1)
            Dim ObjectClass As String = ValueString(2)
            Dim ItemOrContainer As String = ValueString(3)
            If g_bDebug Then Log("NavigationBox_SelectedIndexChanged called for Player = " & ZoneName & " and found ObjectID = " & ObjectID, LogType.LOG_TYPE_INFO)
            If g_bDebug Then Log("NavigationBox_SelectedIndexChanged called for Player = " & ZoneName & " and found ObjectClass = " & ObjectClass, LogType.LOG_TYPE_INFO)
            ar = Nothing
            If (ItemOrContainer.ToUpper <> "CONTAINER") And (ServerAPI IsNot Nothing) And (Not DoubleClick) Then
                DescriptionCell = ServerAPI.GetDescriptionFromObject(ObjectID)
                ButAddToPlaylist.enabled = True
                Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
                Me.divToUpdate.Add("DescriptionCellDiv", "<table><tr>" & DescriptionCell & "</tr></table>")
                Exit Sub
            Else
                DescriptionCell = ""
                Me.divToUpdate.Add("DescriptionCellDiv", "<table><tr>" & DescriptionCell & "</tr></table>")
            End If
            If ItemOrContainer.ToUpper = "ITEM" Then
                ButAddToPlaylist.enabled = True
            Else
                ButAddToPlaylist.enabled = False
            End If
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
            If DoubleClick Then
                AddTrack_Click()
                Exit Sub
            End If
            Dim LevelIndex As Integer = 0
            Try
                Select Case LblWhatsLoaded
                    Case ITEM_TYPE.Root.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level1.ToString
                        LblLevel1.label = ObjectName
                        ObjectIDLevel1 = ObjectID
                        LevelIndex = 1
                    Case ITEM_TYPE.Level1.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level2.ToString
                        LblLevel2.label = ObjectName
                        ObjectIDLevel2 = ObjectID
                        LevelIndex = 2
                    Case ITEM_TYPE.Level2.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level3.ToString
                        LblLevel3.label = ObjectName
                        ObjectIDLevel3 = ObjectID
                        LevelIndex = 3
                    Case ITEM_TYPE.Level3.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level4.ToString
                        LblLevel4.label = ObjectName
                        ObjectIDLevel4 = ObjectID
                        LevelIndex = 4
                    Case ITEM_TYPE.Level4.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level5.ToString
                        LblLevel5.label = ObjectName
                        ObjectIDLevel5 = ObjectID
                        LevelIndex = 5
                    Case ITEM_TYPE.Level5.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level6.ToString
                        LblLevel6.label = ObjectName
                        ObjectIDLevel6 = ObjectID
                        LevelIndex = 6
                    Case ITEM_TYPE.Level6.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level7.ToString
                        LblLevel7.label = ObjectName
                        ObjectIDLevel7 = ObjectID
                        LevelIndex = 7
                    Case ITEM_TYPE.Level7.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level8.ToString
                        LblLevel8.label = ObjectName
                        ObjectIDLevel8 = ObjectID
                        LevelIndex = 8
                    Case ITEM_TYPE.Level8.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level9.ToString
                        LblLevel9.label = ObjectName
                        ObjectIDLevel9 = ObjectID
                        LevelIndex = 9
                    Case ITEM_TYPE.Level9.ToString
                        LblWhatsLoaded = ITEM_TYPE.Level10.ToString
                        LblLevel10.label = ObjectName
                        ObjectIDLevel10 = ObjectID
                        LevelIndex = 10
                    Case Else
                End Select
                UpdateIniFile()
                ar = ServerAPI.GetContainerAtLevel(MyZoneUDN, LevelIndex)
                LoadNavigationBox()
                UpdateStatus()
            Catch ex As Exception
                Log("Error in NavigationBox_SelectedIndexChanged for Player = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("Error in NavigationBox_SelectedIndexChanged (1) for Player = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub UpdateIniFile()
        WriteStringIniFile("DevicePage", MyZoneUDN & "_loaded", LblWhatsLoaded)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_root", LblRoot.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level1", LblLevel1.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level2", LblLevel2.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level3", LblLevel3.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level4", LblLevel4.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level5", LblLevel5.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level6", LblLevel6.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level7", LblLevel7.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level8", LblLevel8.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level9", LblLevel9.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_level10", LblLevel10.label)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDroot", ObjectIDRoot)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel1", ObjectIDLevel1)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel2", ObjectIDLevel2)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel3", ObjectIDLevel3)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel4", ObjectIDLevel4)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel5", ObjectIDLevel5)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel6", ObjectIDLevel6)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel7", ObjectIDLevel7)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel8", ObjectIDLevel8)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel9", ObjectIDLevel9)
        WriteStringIniFile("DevicePage", MyZoneUDN & "_ObjectIDlevel10", ObjectIDLevel10)
    End Sub

    Private Sub ClearSelections()
        LblRoot.label = "Root"
        LblLevel1.label = ""
        LblLevel2.label = ""
        LblLevel3.label = ""
        LblLevel4.label = ""
        LblLevel5.label = ""
        LblLevel6.label = ""
        LblLevel7.label = ""
        LblLevel8.label = ""
        LblLevel9.label = ""
        LblLevel10.label = ""
        ObjectIDRoot = "0"
        ObjectIDLevel1 = ""
        ObjectIDLevel2 = ""
        ObjectIDLevel3 = ""
        ObjectIDLevel4 = ""
        ObjectIDLevel5 = ""
        ObjectIDLevel6 = ""
        ObjectIDLevel7 = ""
        ObjectIDLevel8 = ""
        ObjectIDLevel9 = ""
        ObjectIDLevel10 = ""
    End Sub
End Class

<Serializable()> Public Class GroupArrayElement
    Public UDN As String
    Public Master As Boolean = False
    Public Added As Boolean = False
    Public Confirmed As Boolean = False
End Class
