Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Text
Imports System.IO.Path
Imports System.Xml
Imports System.Drawing

Partial Public Class HSPI

    Implements IMediaAPI_3

    Friend WithEvents MyEventTimer As Timers.Timer

    Const MaxTOActionArray = 10
    Private MyTimeoutActionArray(MaxTOActionArray) As Integer
    '
    ' timeout indexes
    ' 
    Const TORediscover = 0
    Const TORediscoverValue = 600 ' 10 minutes 
    Const TOCheckChange = 1
    Const TOCheckChangeValue = 600 ' 10 minutes
    Const TOCheckAnnouncement = 2
    Const TOCheckAnnouncementValue = 1


    Public MyHSDeviceLinkedList As New LinkedList(Of MyUPnpDeviceInfo)
    Private UPnPControllerConf As UPnPController_Config
    Private ActAsSpeakerProxy As Boolean = False
    Private ProxySpeakerActive As Boolean = False
    Private CapabilitiesCalledFlag As Boolean = False
    Private MyPingAddressLinkedList As New LinkedList(Of PingArrayElement)
    Private MyPingReEntry As Boolean = False
    Private MyLocalMACAddress As String = ""
    Private MyIPMask As String = ""
    Private DeviceConnectFlag As Boolean = False
    Private PlayModeisConfigurable As Boolean = False

    Private HSRefDevice As Integer = -1
    Private HSRefPlayer As Integer = -1
    Private HSRefTrack As Integer = -1
    Private HSRefNextTrack As Integer = -1
    Private HSRefArtist As Integer = -1
    Private HSRefNextArtist As Integer = -1
    Private HSRefAlbum As Integer = -1
    Private HSRefNextAlbum As Integer = -1
    Private HSRefArt As Integer = -1
    Private HSRefNextArt As Integer = -1
    Private HSRefPlayState As Integer = -1
    Private HSRefVolume As Integer = -1
    Private HSRefMute As Integer = -1
    Private HSRefLoudness As Integer = -1
    Private HSRefBalance As Integer = -1
    Private HSRefTrackLength As Integer = -1
    Private HSRefTrackPos As Integer = -1
    Private HSRefTrackDescr As Integer = -1
    Private HSRefRepeat As Integer = -1
    Private HSRefShuffle As Integer = -1
    Private HSRefQueueRepeat As Integer = -1
    Private HSRefQueueShuffle As Integer = -1
    Private HSRefGenre As Integer = -1
    Private HSRefSpeed As Integer = -1
    Private HSRefRemote As Integer = -1
    Private HSRefServiceRemote As Integer = -1
    Private HSRefLGVolume As Integer = -1
    Private HSRefParty As Integer = -1
    Private HSRefMessage As Integer = -1
    Private HSRefServer As Integer = -1
    Private MyMissedPings As Integer = 0
    Private CurrentLibEntry As Lib_Entry
    Private NextLibEntry As Lib_Entry
    Private CurrentLibKey As Lib_Entry_Key
    Private NextLibKey As Lib_Entry_Key

    Public Enum part
        Version = 1
        PlayerUDN = 2
        PlaylistName = 3
        PlaylistItem = 4
        PlayListItemPosition = 5
        Volume = 6
        Mute = 7
        Repeat = 8
        Shuffle = 9
        SlideshowSpeed = 10
    End Enum

    Public Enum QueueActions
        qaDontPlay
        qaPlayLast
        qaPlayNext
        qaPlayNow
    End Enum



    Private RemoteUIServer As MyUPnPService = Nothing
    Private RemoteControlService As MyUPnPService = Nothing
    Private MessageBoxService As MyUPnPService = Nothing

    Private MyUPnPDevice As MyUPnPDevice
    Private ConnectUPnPDevice As Boolean = False

    Friend WithEvents MyMusicAPITimer As Timers.Timer           ' used for Position Updates
    Friend WithEvents MyUPnPDeviceTimer As Timers.Timer         ' used to check device status and re-connect
    Friend WithEvents MyCheckQueueTimer As Timers.Timer
    Friend WithEvents MyPollRenderStateTimer As Timers.Timer
    Friend WithEvents MyPollAVTransportTimer As Timers.Timer

    Private MyUPnPDeviceName As String = ""
    Private MyUPnPDeviceServiceType As String = "" ' could be DMR, PMR, RCR
    Private MyReferenceToMyController As HSPI = Me
    Private MyUPnPModelName As String = ""
    Private MyUPnPModelNumber As String = ""
    Private MyUPnPDeviceManufacturer As String = ""

    Private WaitingToReConnect As Boolean = False
    Private TimerReEntry As Boolean = False
    Private MyMessageServiceActive As Boolean = False
    Private MyRemoteServiceActive As Boolean = False




    Private LastRetrieveAlbumArtURL As String = ""
    Private MyHSTMusicIndex As Integer = 0
    Private MyAPIIndex As Integer = 0
    Private MyDeviceStatus As String = "Offline"
    Private MyFailedPingCount As Integer = 0
    Private MyUDN As String = ""
    Private MyIPAddress As String = ""
    Private MyIPPort As String = ""
    Private MyMacAddress As String = ""
    Private MyIConURL As String = ""
    Private MyPreviousAlbumArtPath As String = ""
    Private MyPreviousNextAlbumArtPath As String = ""
    Private MyPreviousAlbumURI As String = ""
    Private MyPreviousNextAlbumURI As String = ""
    Private MyDocumentURL As String = ""
    Private MyAdminStateActive As Boolean = False
    'Private MyPollForTransportChangeFlag As Boolean = False
    'Private MyPollForVolumeChangeFlag As Boolean = False
    Private MyQueueRepeatState As Boolean = False
    Private MyQueueShuffleState As Boolean = False
    Private MyRandomNumberGenerator As New Random()
    Private MyServerUDN As String = ""
    Private MyPlayerWentThroughPlayState As Boolean = False
    Private MyPlayerWentThroughTrackChange As Boolean = False
    Private MyTimeIveBeenWaitingForPlayState As Integer = 0
    Private MaxWaitTimeToGoThroughPlayState As Integer = 20 ' 20 seconds
    Private MyTrackLength As Integer = 0
    Private MyCurrentTrackYear As Integer = 1962
    Private MyPictureSize As PictureSize
    Private MyRefreshAVGetPositionInfo As Boolean = False
    Private MyCurrentTrackBPM As Integer = 0
    Private MyCurrentTrackDescriptor As String = ""
    Private MyCurrentTrackLastPlayedDate As String = ""
    Private MyCurrentTrackLyrics As String = ""
    Private MyCurrentTrackPlayedCount As Integer = 0
    Private MyCurrentTrackSampleRate As Integer = 0
    Private MyCurrentTrackRating As String = ""
    Private MyCurrentTrackBitRate As Integer = 0
    Private MyCurrentTrackTime As String = ""
    Private MyPossibleFFSpeeds As Integer() = Nothing
    Private MyPossibleREWSpeeds As Integer() = Nothing
    Private MyNextTrackDescriptor As String = ""

    '
    ' Timeout indexes
    Const TOReachable = 0
    Const TOPositionUpdate = 1
    Const TOQueue = 2
    Const TOPollRenderState = 3
    Const TOProcessBuildDB = 4
    Const TOPollAVTransportState = 5
    '
    ' Timeout Values
    Const TOReachableValue = 10
    Const TOPositionUpdateValue = 1
    Const TOQueueValue = 1
    Const TOPollRenderStateValue = 5
    Const TOProcessBuildDBValue = 60
    Const TOPollAVTransportStateValue = 2


    Public WriteOnly Property pDevice As MyUPnPDevice
        Set(value As MyUPnPDevice)
            MyUPnPDevice = value
        End Set
    End Property

    Public Enum SearchOperatorTypes
        soContains = 0
        soIsEqual = 1
    End Enum


    Private Sub MyMusicAPITimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyMusicAPITimer.Elapsed
        Dim Index As Integer
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyMusicAPITimer_Elapsed called for UPnPDevice " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO, LogColorNavy)
        If MyRefreshAVGetPositionInfo Then
            MyRefreshAVGetPositionInfo = False
            If PIDebuglevel > DebugLevel.dlEvents Then Log("MyMusicAPITimer_Elapsed called for UPnPDevice " & MyUPnPDeviceName & " is refreshing the PositionInfo", LogType.LOG_TYPE_INFO)
            AVTGetPositionInfo(0)
        End If
        For Index = 0 To MaxTOActionArray
            If MyTimeoutActionArray(Index) <> 0 Then
                Select Case Index
                    Case TOReachable
                        'MyTimeoutActionArray(Index) = MyTimeoutActionArray(Index) - 1
                        'If MyTimeoutActionArray(Index) <= 0 Then
                        'Reachable()
                        'MyTimeoutActionArray(Index) = TOReachableValue
                        'End If
                    Case TOPositionUpdate
                        MyTimeoutActionArray(Index) = MyTimeoutActionArray(Index) - 1
                        If MyTimeoutActionArray(Index) <= 0 Then
                            Try
                                UpdatePositionInfo()
                            Catch ex As Exception
                            End Try
                            MyTimeoutActionArray(Index) = TOPositionUpdateValue
                        End If
                    Case TOQueue
                    Case TOPollRenderState
                    Case TOPollAVTransportState
                    Case TOProcessBuildDB
                        MyTimeoutActionArray(Index) = MyTimeoutActionArray(Index) - 1
                        If MyTimeoutActionArray(Index) <= 0 Then
                            Try
                                If GetStringIniFile(MyUDN, DeviceInfoIndex.diSystemUpdateID.ToString, "") <> GetStringIniFile(MyUDN, DeviceInfoIndex.diSystemUpdateIDAtDBCreation.ToString, "") Then
                                    'BuildDatabase(CurrentAppPath & DBPath & MyUDN & ".mdb")
                                End If
                            Catch ex As Exception
                            End Try
                            MyTimeoutActionArray(Index) = 0
                        End If
                End Select
            End If
        Next
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyUPnPDeviceTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyUPnPDeviceTimer.Elapsed
        If TimerReEntry Then Exit Sub
        If Not MyAdminStateActive Then Exit Sub
        If MyUPnPDeviceServiceType = "HST" Then Exit Sub ' no need to ping this device
        TimerReEntry = True
        Try
            If ConnectUPnPDevice Then
                Connect("uuid:" & MyUDN)
                ConnectUPnPDevice = False
            Else
                Reachable()
            End If
            sender = Nothing
            e = Nothing
        Catch ex As Exception
        End Try
        TimerReEntry = False
    End Sub

    Private Sub MyEventTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyEventTimer.Elapsed
        If DeviceConnectFlag Then
            DeviceConnectFlag = False
        End If
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyCheckQueueTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyCheckQueueTimer.Elapsed
        Try
            'If piDebuglevel > DebugLevel.dlEvents Then Log("MyCheckQueueTimer_Elapsed called for UPnPDevice " & MyUPnPDeviceName & " is checking the queue", LogType.LOG_TYPE_INFO)
            If DeviceStatus = "Offline" Then Exit Sub
            CheckQueue()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyCheckQueueTimer_Elapsed for UPnPDevice " & MyUPnPDeviceName & " while checking the queue with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub MyPollRenderStateTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyPollRenderStateTimer.Elapsed
        Try
            If PIDebuglevel > DebugLevel.dlEvents Then Log("MyPollRenderStateTimer_Elapsed called for UPnPDevice " & MyUPnPDeviceName & " is polling the RenderState", LogType.LOG_TYPE_INFO)
            If DeviceStatus = "Offline" Then Exit Sub
            If Not MyPollRenderStateReEntry Then CheckRenderState()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyPollRenderStateTimer_Elapsed for UPnPDevice " & MyUPnPDeviceName & " while checking the RenderState with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub MyPollAVTransportTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyPollAVTransportTimer.Elapsed
        Try
            If ProcessAVXMLReEntryFlag Then Exit Sub ' prevent reentries. This proc is called each second so we can skip one time
            If DeviceStatus = "Offline" Then Exit Sub
            MyTrackInfoHasChanged = False
            MyTransportStateHasChanged = False
            MyAlbumArtURIHasChanged = False
            MyNextTrackInfoHasChanged = False
            MyNextAlbumArtURIHasChanged = False
            MyDurationInfoHasChanged = False
            MyTittleWasPresent = False
            If PIDebuglevel > DebugLevel.dlEvents Then Log("MyPollAVTransportTimer_Elapsed called for UPnPDevice " & MyUPnPDeviceName & " is polling the transportstate", LogType.LOG_TYPE_INFO)
            AVTGetTransportInfo(0)
            AVTGetPositionInfo(0)
            'If MyCurrentPlayerState <> ConvertTransportStateToPlayerState(MyCurrentTransportState) Then
            'MyCurrentPlayerState = ConvertTransportStateToPlayerState(MyCurrentTransportState)
            'AVTGetPositionInfo()
            'End If
            If MyTrackInfoHasChanged Then MyPlayerWentThroughTrackChange = True
            If NextAvTransportIsAvailable And UseNextAvTransport Then CheckNextURIisPresent()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyPollAVTransportTimer_Elapsed for UPnPDevice " & MyUPnPDeviceName & " while checking the AVTransportState with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub InitMusicAPI()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitMusicAPI called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        MyMusicAPITimer = New Timers.Timer
        MyMusicAPITimer.Interval = 1000 ' every second
        MyMusicAPITimer.AutoReset = True
        MyMusicAPITimer.Enabled = True
        Dim Index As Integer
        For Index = 0 To MaxTOActionArray
            MyTimeoutActionArray(Index) = 0
        Next
        MyUPnPDeviceTimer = New Timers.Timer
        MyUPnPDeviceTimer.Interval = 20000 ' every 20 seconds
        MyUPnPDeviceTimer.AutoReset = True
        MyUPnPDeviceTimer.Enabled = True

        MyTimeoutActionArray(TOReachable) = TOReachableValue
        MyTimeoutActionArray(TOPositionUpdate) = TOPositionUpdateValue

        MyCheckQueueTimer = New Timers.Timer
        MyCheckQueueTimer.Interval = 1000 ' every second
        MyCheckQueueTimer.AutoReset = True
        MyCheckQueueTimer.Enabled = True

        If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diPollTransportChanges.ToString, False) Then StartAVPollTimer()     ' changed 12/2/2018 v.039
        If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diPollVolumeChanges.ToString, False) Then StartRenderPollTimer()    ' changed 12/2/2018 v0.39

        With CurrentLibKey
            .iKey = 1
            .Library = 1
            .sKey = "1"
            .Title = ""
            .WhichKey = eKey_Type.eEither
        End With

        With NextLibKey
            .iKey = 1
            .Library = 1
            .sKey = "1"
            .Title = ""
            .WhichKey = eKey_Type.eEither
        End With

        With CurrentLibEntry
            .Title = ""
            .Album = ""
            .Artist = ""
            .Cover_path = NoArtPath
            .Cover_Back_path = NoArtPath
            .Genre = ""
            .Key = CurrentLibKey
            .Kind = ""
            .LengthSeconds = 0
            .Lib_Media_Type = eLib_Media_Type.Music
            .Lib_Type = 1
            .PlayedCount = 0
            .Rating = 0
            .Year = 0
        End With

        With NextLibEntry
            .Title = ""
            .Album = ""
            .Artist = ""
            .Cover_path = NoArtPath
            .Cover_Back_path = NoArtPath
            .Genre = ""
            .Key = NextLibKey
            .Kind = ""
            .LengthSeconds = 0
            .Lib_Media_Type = eLib_Media_Type.Music
            .Lib_Type = 1
            .PlayedCount = 0
            .Rating = 0
            .Year = 0
        End With
        DeviceStatus = "Offline"
    End Sub

    Private Sub StartAVPollTimer()
        If MyPollAVTransportTimer Is Nothing Then
            MyPollAVTransportTimer = New Timers.Timer
        End If
        MyPollAVTransportTimer.Interval = 1000 ' every second
        MyPollAVTransportTimer.AutoReset = True
        MyPollAVTransportTimer.Enabled = True
    End Sub

    Private Sub StartRenderPollTimer()
        If MyPollRenderStateTimer Is Nothing Then
            MyPollRenderStateTimer = New Timers.Timer
        End If
        MyPollRenderStateTimer.Interval = 2000 ' every 2 second
        MyPollRenderStateTimer.AutoReset = True
        MyPollRenderStateTimer.Enabled = True
    End Sub

    Private Sub StopAVPollTimer()
        If MyPollAVTransportTimer Is Nothing Then Exit Sub
        Try
            MyPollAVTransportTimer.AutoReset = False
            MyPollAVTransportTimer.Enabled = False
            MyPollAVTransportTimer.Dispose()
        Catch ex As Exception
        End Try
        MyPollAVTransportTimer = Nothing
    End Sub

    Private Sub StopRenderPollTimer()
        If MyPollRenderStateTimer Is Nothing Then Exit Sub
        Try
            MyPollRenderStateTimer.AutoReset = False
            MyPollRenderStateTimer.Enabled = False
            MyPollRenderStateTimer.Dispose()
        Catch ex As Exception
        End Try
        MyPollRenderStateTimer = Nothing
    End Sub

    Public Sub DestroyPlayer(disposing As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroyPlayer called for device = " & MyUPnPDeviceName & " and Disposing = " & disposing.ToString, LogType.LOG_TYPE_INFO)
        If disposing Then
            ' Free other state (managed objects).
            Try
                If DeviceStatus.ToUpper = "ONLINE" And MyUPnPDeviceServiceType = "DMR" Then
                    StopPlay()
                End If
            Catch ex As Exception
            End Try
            Try
                Disconnect(True)
            Catch ex As Exception
            End Try
            Try
                ClearQueue()
            Catch ex As Exception
            End Try
            Try
                CurrentLibEntry.Key = Nothing
                NextLibEntry.Key = Nothing
                CurrentLibEntry = Nothing
                NextLibEntry = Nothing
            Catch ex As Exception

            End Try
            Try
                If AVTransport IsNot Nothing Then AVTransport.RemoveCallback()
            Catch ex As Exception
            End Try
            AVTransport = Nothing
            Try
                If ContentDirectory IsNot Nothing Then ContentDirectory.RemoveCallback()
            Catch ex As Exception
            End Try
            ContentDirectory = Nothing
            Try
                If SonyPartyService IsNot Nothing Then SonyPartyService.RemoveCallback()
            Catch ex As Exception
            End Try
            SonyPartyService = Nothing
            Try
                If ConnectionManager IsNot Nothing Then ConnectionManager.RemoveCallback()
            Catch ex As Exception
            End Try
            ConnectionManager = Nothing
            StopAVPollTimer()           ' added 12/2/2018 v.039
            StopRenderPollTimer()
            Try
                MyMusicAPITimer.Enabled = False
                MyUPnPDeviceTimer.Enabled = False
                MyCheckQueueTimer.Enabled = False
                'MyPollAVTransportTimer.Enabled = False ' removed 12/2/2018 v.039
                'MyPollRenderStateTimer.Enabled = False
                MyMusicAPITimer = Nothing
                MyUPnPDeviceTimer = Nothing
                MyCheckQueueTimer = Nothing
                MyPollAVTransportTimer = Nothing
                MyPollRenderStateTimer = Nothing
                RemoteUIServer = Nothing
                RemoteControlService = Nothing
                MessageBoxService = Nothing
                MyUPnPDevice = Nothing
                MyReferenceToMyController = Nothing
                MyRandomNumberGenerator = Nothing
                myConnectionManagerCallback = Nothing
                myContentDirectoryCallback = Nothing
                myAVTransportCallback = Nothing
                MyQueueLinkedList = Nothing
                MySourceProtocolInfo = Nothing
                MySinkProtocolInfo = Nothing
            Catch ex As Exception
            End Try
            Try
                If Not MyPossibleFFSpeeds Is Nothing Then
                    For index = 0 To UBound(MyPossibleFFSpeeds, 1)
                        MyPossibleFFSpeeds(index) = ""
                    Next
                    MyPossibleFFSpeeds = Nothing
                End If
            Catch ex As Exception
            End Try
            Try
                If Not MyPossibleREWSpeeds Is Nothing Then
                    For index = 0 To UBound(MyPossibleREWSpeeds, 1)
                        MyPossibleREWSpeeds(index) = ""
                    Next
                    MyPossibleREWSpeeds = Nothing
                End If
            Catch ex As Exception
            End Try
        End If
        ' Free your own state (unmanaged objects).
        ' Set large fields to null.
    End Sub

    Public Sub TreatSetIOEx(CC As CAPIControl)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOEx called for device = " & MyUPnPDeviceName & " with  Ref = " & CC.Ref.ToString & ", Index " & CC.CCIndex.ToString & ", controlFlag = " & CC.ControlFlag.ToString &
                 ", ControlString" & CC.ControlString.ToString & ", ControlType = " & CC.ControlType.ToString & ", ControlValue = " & CC.ControlValue.ToString &
                  ", Label = " & CC.Label.ToString, LogType.LOG_TYPE_INFO)
        Dim UPnPDevice As HSPI = Me

        ' treat with "local" requests
        Select Case CC.Ref
            Case HSRefDevice
                Select Case CC.ControlValue
                    Case dsActivate
                        SetAdministrativeState(True)
                    Case dsDeactivate
                        SetAdministrativeState(False)
                    Case dsWOL
                        SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), PlugInIPAddress, GetSubnetMask())
                        SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, ""), PlugInIPAddress, GetSubnetMask())
                    Case Else
                        Select Case MyUPnPDeviceServiceType
                            Case "DIAL"
                                TreatSetIOExRoku(CC.ControlValue)
                        End Select
                End Select
            Case HSRefServer
                Select Case CC.ControlValue
                    Case dsBuildDB
                        'BuildDatabase(CurrentAppPath & DBPath & MyUDN & ".mdb")
                End Select
            Case HSRefBalance
                Select Case CC.ControlValue
                    Case vpDown
                        UPnPDevice.ChangeBalanceLevel("LF", 10)
                    Case vpUp
                        UPnPDevice.ChangeBalanceLevel("RF", 10)
                    Case Else
                        UPnPDevice.SetBalance(CC.ControlValue)
                        ' this should be the slider
                End Select
            Case HSRefLoudness
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.SetLoudness = True
                    Case tpOff
                        UPnPDevice.SetLoudness = False
                    Case tpToggle
                        UPnPDevice.ToggleLoudnessState("Master")
                End Select
            Case HSRefMute
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.SetMute("Master", True)
                    Case tpOff
                        UPnPDevice.SetMute("Master", False)
                    Case tpToggle
                        UPnPDevice.ToggleMuteState("Master")
                    Case tpLGOn, tpLGOff, tpLGToggle
                        UPnPDevice.SetLGMute(CC.ControlValue)
                End Select
            Case HSRefPlayState
                Select Case CC.ControlValue
                    Case psPlay
                        UPnPDevice.SetPlayState("Play")
                    Case psStop
                        UPnPDevice.SetPlayState("Stop")
                    Case psPause
                        UPnPDevice.TogglePause()
                    Case psFF
                        DoFastForward()
                    Case psREW
                        DoRewind()
                End Select
            Case HSRefPlayer
                Select Case CC.ControlValue
                    Case psPlay
                        UPnPDevice.SetPlayState("Play")
                    Case psStop
                        UPnPDevice.SetPlayState("Stop")
                    Case psPause
                        UPnPDevice.TogglePause()
                    Case psPrevious
                        UPnPDevice.SetPlayState("Previous")
                    Case psNext
                        UPnPDevice.SetPlayState("Next")
                    Case psShuffle  'Shuffle
                        UPnPDevice.ToggleShuffle()
                    Case psRepeat  ' Repeat
                        UPnPDevice.ToggleRepeat()
                    Case psQueueShuffle
                        UPnPDevice.ToggleQueueShuffle()
                    Case psQueueRepeat
                        UPnPDevice.ToggleQueueRepeat()
                    Case psVolUp  ' Up
                        UPnPDevice.ChangeVolumeLevel("Master", MyVolumeStep)
                    Case psVolDown   ' Down
                        UPnPDevice.ChangeVolumeLevel("Master", -MyVolumeStep)
                    Case psMute   ' mute
                        UPnPDevice.ToggleMuteState("Master")
                    Case psBalanceLeft    ' Left
                        UPnPDevice.ChangeBalanceLevel("LF", 10)
                    Case psBalanceRight   ' Right
                        UPnPDevice.ChangeBalanceLevel("RF", 10)
                    Case psLoudness   ' Loudness
                        UPnPDevice.ToggleLoudnessState("Master")
                    Case psClearQueue
                        ClearQueue()
                        SaveCurrentPlaylistTracks("")
                    Case psFF
                        DoFastForward()
                    Case psREW
                        DoRewind()
                    Case Else
                        If CC.ControlValue <= MyMaximumVolume Then
                            UPnPDevice.SetVolumeLevel("Master", CC.ControlValue)
                        ElseIf CC.ControlValue >= 200 Then
                            UPnPDevice.SetBalance(CC.ControlValue - vpMidPoint)
                        End If
                End Select
            Case HSRefRepeat
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.Repeat = repeat_modes.repeat_all
                    Case tpOff
                        UPnPDevice.Repeat = repeat_modes.repeat_off
                    Case tpToggle
                        UPnPDevice.ToggleRepeat()
                End Select
            Case HSRefShuffle
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.Shuffle = Shuffle_modes.Shuffled
                    Case tpOff
                        UPnPDevice.Shuffle = Shuffle_modes.Ordered
                    Case tpToggle
                        UPnPDevice.ToggleShuffle()
                End Select
            Case HSRefQueueRepeat
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.QueueRepeat = repeat_modes.repeat_all
                    Case tpOff
                        UPnPDevice.QueueRepeat = repeat_modes.repeat_off
                    Case tpToggle
                        UPnPDevice.ToggleQueueRepeat()
                End Select
            Case HSRefQueueShuffle
                Select Case CC.ControlValue
                    Case tpOn
                        UPnPDevice.QueueShuffle = Shuffle_modes.Shuffled
                    Case tpOff
                        UPnPDevice.QueueShuffle = Shuffle_modes.Ordered
                    Case tpToggle
                        UPnPDevice.ToggleQueueShuffle()
                End Select
            Case HSRefTrackPos
                Select Case CC.ControlValue
                    Case vpDown
                        UPnPDevice.AVTSeek(AVT_SeekMode.REL_TIME, ConvertSecondsToTimeFormat(MyCurrentPlayerPosition - 10))
                    Case vpUp
                        UPnPDevice.AVTSeek(AVT_SeekMode.REL_TIME, ConvertSecondsToTimeFormat(MyCurrentPlayerPosition + 10))
                    Case Else
                        Try
                            If MyHSTrackPositionFormat = HSSTrackPositionSettings.TPSPercentage Then
                                UPnPDevice.AVTSeek(AVT_SeekMode.REL_TIME, ConvertSecondsToTimeFormat(Val(CC.ControlValue / 100 * MyTrackLength)))
                            Else
                                UPnPDevice.AVTSeek(AVT_SeekMode.REL_TIME, ConvertSecondsToTimeFormat(Val(CC.ControlValue)))
                            End If
                        Catch ex As Exception
                        End Try
                End Select
                'If Not MyPollForTransportChangeFlag Then GetPositionDurationInfoOnly(0)
            Case HSRefVolume
                Select Case CC.ControlValue
                    Case vpDown
                        UPnPDevice.ChangeVolumeLevel("Master", -MyVolumeStep)
                    Case vpUp
                        UPnPDevice.ChangeVolumeLevel("Master", MyVolumeStep)
                    Case vpLGUp, vpLGDown
                        UPnPDevice.setLGVolume(CC.ControlValue)
                    Case Else
                        ' this should be the slider
                        If HSRefVolume <> HSRefLGVolume Then
                            UPnPDevice.SetVolumeLevel("Master", CC.ControlValue)
                        Else
                            UPnPDevice.setLGVolume(CC.ControlValue)
                        End If
                End Select
            Case HSRefSpeed
                ' the speed is cc.controlValue - gsdefault
                Try
                    AVTPlay(CC.ControlValue - gsDefault)
                Catch ex As Exception
                    Log("Error in TreatSetIOEx for UPnPDevice = " & MyUPnPDeviceName & " while speed set to = " & (CC.ControlValue - gsDefault).ToString & " and with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Case HSRefRemote, HSRefServiceRemote
                TreatSetIOExRemoteControl(CC.ControlValue)
            Case HSRefParty
                TreatSetIOExSonyParty(CC.ControlValue)
            Case HSRefMessage
                Select Case CC.ControlValue
                    Case dsActivate
                        SetAdministrativeState(True)
                    Case dsDeactivate
                        SetAdministrativeState(False)
                End Select
            Case MyBrightnessHSRef
                ' the values are percentages between 0 = no and 20 = max on a scale from 0 to 100
                Select Case CC.ControlValue
                    Case psVolDown
                        BrightnessDown()
                    Case psVolUp
                        BrightnessUp()
                    Case Else
                        If CC.ControlValue >= 0 And CC.ControlValue <= 20 Then
                            RCSetBrightness(CC.ControlValue * 5)
                        End If
                End Select

            Case MyColorHSRef
                ' the values are percentages between 0 = no and 20 = max on a scale from 0 to 100
                Select Case CC.ControlValue
                    Case psVolUp
                        ColorDown()
                    Case psVolDown
                        ColorUp()
                    Case Else
                        If CC.ControlValue >= 0 And CC.ControlValue <= 20 Then
                            RCSetColorTemperature(CC.ControlValue * 5)
                        End If
                End Select
            Case MyContrastHSRef
                ' the values are percentages between 0 = no and 20 = max on a scale from 0 to 100
                Select Case CC.ControlValue
                    Case psVolUp
                        ContrastDown()
                    Case psVolDown
                        ContrastUp()
                    Case Else
                        If CC.ControlValue >= 0 And CC.ControlValue <= 20 Then
                            RCSetContrast(CC.ControlValue * 5)
                        End If
                End Select
            Case MySharpnessHSRef
                ' the values are percentages between 0 = no and 20 = max on a scale from 0 to 100
                Select Case CC.ControlValue
                    Case psVolDown
                        SharpnessDown()
                    Case psVolUp
                        SharpnessUp()
                    Case Else
                        If CC.ControlValue >= 0 And CC.ControlValue <= 20 Then
                            RCSetSharpness(CC.ControlValue * 5)
                        End If
                End Select
            Case MySlideshowEffectHSRef
                Select Case CC.ControlValue
                    Case psVolDown
                        'SlideShowEffectDown()
                    Case psVolUp
                        'SlideShowEffectUp()
                    Case Else
                End Select
            Case MyImageScaleHSRef
                Select Case CC.ControlValue
                    Case psVolDown
                        ImageScaleDown()
                    Case psVolUp
                        ImageScaleUp()
                    Case Else
                End Select
            Case MyImageRotationHSRef
                Select Case CC.ControlValue
                    Case psVolDown
                        ImageRotationDown()
                    Case psVolUp
                        ImageRotationUp()
                    Case Else
                End Select

            Case Else

        End Select



    End Sub

    Public Property DeviceName As String ' 
        Get
            DeviceName = MyUPnPDeviceName
        End Get
        Set(value As String)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceName called for device - " & MyUPnPDeviceName & " with value = " & value.ToString, LogType.LOG_TYPE_INFO)
            MyUPnPDeviceName = value
        End Set
    End Property

    Public Property DeviceServiceType As String
        Get
            DeviceServiceType = MyUPnPDeviceServiceType
        End Get
        Set(value As String)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceServiceType Set called  for device - " & MyUPnPDeviceName & " with ServiceType = " & value.ToString, LogType.LOG_TYPE_INFO)
            MyUPnPDeviceServiceType = value
            If MyUPnPDeviceServiceType = "HST" Then
                'DeviceStatus = "Online"
                MyPlayFromQueue = True
                MyIConURL = ImagesPath & "HSTLogo.gif"
            End If
        End Set
    End Property

    Public Property DeviceHSRef As Integer
        Get
            DeviceHSRef = HSRefDevice
        End Get
        Set(value As Integer)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceHSCode Set called  for device - " & MyUPnPDeviceName & "  with DeviceRef = " & value.ToString, LogType.LOG_TYPE_INFO)
            HSRefDevice = value
        End Set
    End Property

    Public Property DeviceAPIIndex As Integer
        Get
            DeviceAPIIndex = MyAPIIndex
        End Get
        Set(value As Integer)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceAPIIndex called for device = " & MyUPnPDeviceName & " with API Index = " & value.ToString, LogType.LOG_TYPE_INFO)
            MyAPIIndex = value
        End Set
    End Property

    Public WriteOnly Property ReferenceToController As HSPI
        Set(value As HSPI)
            MyReferenceToMyController = value
        End Set
    End Property

    Public Property DeviceUDN As String
        Get
            DeviceUDN = MyUDN
        End Get
        Set(value As String)
            MyUDN = value
        End Set
    End Property

    Public Property DeviceIPAddress As String
        Get
            DeviceIPAddress = MyIPAddress
        End Get
        Set(value As String)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("DeviceIPAddress called for UPnPDevice = " & MyUPnPDeviceName & " with IPAddress = " & value.ToString, LogType.LOG_TYPE_INFO)
            MyIPAddress = value
        End Set
    End Property

    Public Property DeviceIPPort As String
        Get
            DeviceIPPort = MyIPPort
        End Get
        Set(value As String)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("DeviceIPPort called for UPnPDevice = " & MyUPnPDeviceName & " with IPPort = " & value.ToString, LogType.LOG_TYPE_INFO)
            MyIPPort = value
        End Set
    End Property

    Public Property DeviceModelName As String
        Get
            DeviceModelName = MyUPnPModelName
        End Get
        Set(value As String)
            MyUPnPModelName = value
        End Set
    End Property

    Public Property DeviceModelNumber As String
        Get
            DeviceModelNumber = MyUPnPModelNumber
        End Get
        Set(value As String)
            MyUPnPModelNumber = value
        End Set
    End Property

    Public Property DeviceAdminStateActive As Boolean
        Get
            DeviceAdminStateActive = MyAdminStateActive
        End Get
        Set(value As Boolean)
            MyAdminStateActive = value
        End Set
    End Property


    Public Property DeviceManufacturer As String
        Get
            DeviceManufacturer = MyUPnPDeviceManufacturer
        End Get
        Set(value As String)
            MyUPnPDeviceManufacturer = value
        End Set
    End Property

    Public Property Track As String
        Get
            Track = MyCurrentTrack
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then log( "Track Get for device = " & MyUPnPDeviceName & ". Track = " & MyCurrentTrack)
        End Get
        Set(ByVal value As String)
            If MyCurrentTrack <> value And HSRefTrack <> -1 Then
                hs.SetDeviceString(HSRefTrack, value, True)
            End If
            MyCurrentTrack = value
            CurrentLibEntry.Title = value
            CurrentLibKey.Title = value
            If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("Track Set for device = " & MyUPnPDeviceName & ". Track = " & MyCurrentTrack, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextTrack As String
        Get
            'Returns the name of the next track.
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "NextTrack called for device - " & MyUPnPDeviceName & ". Value= " & MyNextTrack.ToString)
            NextTrack = MyNextTrack
        End Get
        Set(ByVal value As String)
            If MyNextTrack <> value And HSRefNextTrack <> -1 Then
                hs.SetDeviceString(HSRefNextTrack, value, True)
            End If
            MyNextTrack = value
            NextLibEntry.Title = value
            NextLibKey.Title = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("NextTrack Set for device = " & MyUPnPDeviceName & ". NextTrack = " & MyNextTrack, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property Artist As String
        Get
            Artist = MyCurrentArtist
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then log( "Artist Get for device = " & MyUPnPDeviceName & ". Artist = " & MyCurrentArtist)
        End Get
        Set(ByVal value As String)
            If MyCurrentArtist <> value And HSRefArtist <> -1 Then
                hs.SetDeviceString(HSRefArtist, value, True)
            End If
            MyCurrentArtist = value
            CurrentLibEntry.Artist = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("Artist Set for device = " & MyUPnPDeviceName & ". Artist = " & MyCurrentArtist, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextArtist As String
        Get
            'Returns the artist's name of the next track.
            NextArtist = MyNextArtist
        End Get
        Set(ByVal value As String)
            If MyNextArtist <> value And HSRefNextArtist <> -1 Then
                hs.SetDeviceString(HSRefNextArtist, value, True)
            End If
            MyNextArtist = value
            NextLibEntry.Artist = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("NextArtist Set for device = " & MyUPnPDeviceName & ". NextArtist = " & MyNextArtist, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property Album As String
        Get
            Album = MyCurrentAlbum
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then log( "Album Get for device = " & MyUPnPDeviceName & ". Album = " & MyCurrentAlbum)
        End Get
        Set(ByVal value As String)
            If MyCurrentAlbum <> value And HSRefAlbum <> -1 Then
                hs.SetDeviceString(HSRefAlbum, value, True)
            End If
            MyCurrentAlbum = value
            CurrentLibEntry.Album = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("Album Set for device = " & MyUPnPDeviceName & ". Album = " & MyCurrentAlbum, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextAlbum As String
        Get
            'Returns the album name of the next track.
            NextAlbum = MyNextAlbum
        End Get
        Set(ByVal value As String)
            If MyNextAlbum <> value And HSRefNextAlbum <> -1 Then
                hs.SetDeviceString(HSRefNextAlbum, value, True)
            End If
            MyNextAlbum = value
            NextLibEntry.Album = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("NextAlbum Set for device = " & MyUPnPDeviceName & ". NextAlbum = " & MyNextAlbum, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property CurrentTrackDuration As String
        Get
            CurrentTrackDuration = MyTrackLength
        End Get
    End Property

    Public Property CurrentTrackDescription As String
        Get
            CurrentTrackDescription = MyCurrentTrackDescriptor
        End Get
        Set(ByVal value As String)
            If MyCurrentTrackDescriptor <> value And HSRefTrackDescr <> -1 Then
                hs.SetDeviceString(HSRefTrackDescr, value, True)
            End If
            MyCurrentTrackDescriptor = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("CurrentTrackDescription for device = " & MyUPnPDeviceName & ". Track Descriptor = " & MyCurrentTrackDescriptor, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextTrackDescription As String
        Get
            NextTrackDescription = MyNextTrackDescriptor
        End Get
        Set(ByVal value As String)
            MyNextTrackDescriptor = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("NextTrackDescription for device = " & MyUPnPDeviceName & ". NextTrack Descriptor = " & MyNextTrackDescriptor, LogType.LOG_TYPE_INFO)
        End Set
    End Property


    Public WriteOnly Property SetVolume As Integer
        Set(ByVal value As Integer)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetVolume is setting HS Status Volume for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefVolume.ToString, LogType.LOG_TYPE_INFO)
            If MyCurrentVolumeLevel <> value Then
                If HSRefVolume <> -1 Then
                    hs.SetDeviceValueByRef(HSRefVolume, value, True)
                End If
            End If
            MyCurrentVolumeLevel = value
        End Set
    End Property

    Public WriteOnly Property SetMuteState As Boolean
        Set(ByVal value As Boolean)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SetMuteState is setting HS Status Mute for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefMute.ToString, LogType.LOG_TYPE_INFO) ' changed on 11/30/2018. Entry shows up due to polling
            If HSRefMute <> -1 Then
                If MyCurrentMuteState <> value Then
                    If value Then
                        hs.SetDeviceValueByRef(HSRefMute, msMuted, True)
                    Else
                        hs.SetDeviceValueByRef(HSRefMute, msUnmuted, True)
                    End If
                End If
                MyCurrentMuteState = value
            End If
        End Set
    End Property

    Public WriteOnly Property SetLoudness As Boolean
        Set(value As Boolean)
            If value <> MyCurrentLoudness Then
                If HSRefLoudness <> -1 Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("SetLoudnessState is setting HS Status Loundness for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefVolume.ToString, LogType.LOG_TYPE_INFO)
                    If value Then
                        hs.SetDeviceValueByRef(HSRefLoudness, lsLoudnessOn, True)
                        'hs.SetDeviceString(HSRefLoudness, "Loudness On", True)
                    Else
                        hs.SetDeviceValueByRef(HSRefLoudness, lsLoudnessOff, True)
                        'hs.SetDeviceString(HSRefLoudness, "Loudness Off", True)
                    End If
                End If
                MyCurrentLoudness = value
            End If
        End Set
    End Property

    Public Property Genre As String
        Get
            Genre = MyCurrentGenre
        End Get
        Set(ByVal value As String)
            If MyCurrentGenre <> value And HSRefGenre <> -1 Then
                hs.SetDeviceString(HSRefGenre, value, True)
            End If
            MyCurrentGenre = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("Genre Set for device = " & MyUPnPDeviceName & ". Genre = " & MyCurrentGenre, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property CurrentPlayerState As player_state_values
        Get
            If MyPlayFromQueue Then
                CurrentPlayerState = MyQueuePlayState
            Else
                CurrentPlayerState = MyCurrentPlayerState
            End If
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then log( "CurrentPlayerState Get for device = " & MyUPnPDeviceName & ". Value = " & CurrentPlayerState.ToString)
            'log( "CurrentPlayerState Get for device = " & MyUPnPDeviceName & ". Value = " & MyCurrentPlayerState.ToString)
        End Get
        Set(ByVal value As player_state_values)
            'If MyPlayFromQueue Then
            'MyQueuePlayState = value  ' max need fixing
            'Else
            If MyCurrentPlayerState <> value And HSRefPlayState <> -1 Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CurrentPlayerState is setting HS Status for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefPlayState.ToString, LogType.LOG_TYPE_INFO)
                hs.SetDeviceValueByRef(HSRefPlayState, value, True)
            End If
            MyCurrentPlayerState = value
            'End If
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled And Not MyPollForTransportChangeFlag Then Log("CurrentPlayerState Set for device = " & MyUPnPDeviceName & ". Value = " & MyCurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property CurrentPlayerSpeed As Integer
        Get
            CurrentPlayerSpeed = MyCurrentTransportPlaySpeed
        End Get
        Set(value As Integer)
            If MyCurrentTransportPlaySpeed <> value And HSRefSpeed <> -1 Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CurrentPlayerSpeed is setting HS Status for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefSpeed.ToString, LogType.LOG_TYPE_INFO)
                hs.SetDeviceValueByRef(HSRefSpeed, gsDefault + value, True)
            End If
            MyCurrentTransportPlaySpeed = value
            If PIDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("CurrentPlayerSpeed Set for device = " & MyUPnPDeviceName & ". Value = " & MyCurrentTransportPlaySpeed.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property NoQueuePlayerState As player_state_values
        Get
            NoQueuePlayerState = MyCurrentPlayerState
        End Get
    End Property

    Private Sub UpdateBalance()
        If MyCurrentLFVolumeLevel < MyMaximumVolume Then
            MyBalance = MyMaximumVolume - MyCurrentLFVolumeLevel
        ElseIf MyCurrentRFVolumeLevel < MyMaximumVolume Then
            MyBalance = -MyMaximumVolume + MyCurrentRFVolumeLevel
        Else
            MyBalance = 0
        End If
        If HSRefBalance <> -1 Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateBalance is setting HS Status for device - " & MyUPnPDeviceName & " with Value = " & MyBalance.ToString & " and HSRef = " & HSRefBalance.ToString, LogType.LOG_TYPE_INFO)
            'hs.SetDeviceString(HSRefBalance, MyBalance.ToString, True)
            hs.SetDeviceValueByRef(HSRefBalance, MyBalance, True)
        End If
    End Sub


    Public Property ArtworkURL As String
        Get
            ArtworkURL = MyCurrentArtworkURL
        End Get
        Set(ByVal value As String)
            If Not (value.ToLower().StartsWith("http://") Or value.ToLower().StartsWith("https://") Or value.ToLower().StartsWith("file:")) Then
                Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                'MyCurrentArtworkURL = "http://" & hs.GetIPAddress & HTTPPort & value
            End If
            If MyCurrentArtworkURL <> value And HSRefArt <> -1 Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ArtworkURL is setting HS Status for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefArt.ToString, LogType.LOG_TYPE_INFO)
                hs.DeviceVSP_ClearAll(HSRefArt, True)
                hs.DeviceVGP_ClearAll(HSRefArt, True)
                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = gsDefault
                Pair.Status = value
                hs.DeviceVSP_AddPair(HSRefArt, Pair)
                Dim GraphicsPair As VGPair
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = value
                GraphicsPair.Set_Value = gsDefault
                hs.DeviceVGP_AddPair(HSRefArt, GraphicsPair)
                hs.SetDeviceValueByRef(HSRefArt, gsDefault, True)
                hs.SetDeviceString(HSRefArt, value, True)
            End If
            MyCurrentArtworkURL = value
            CurrentLibEntry.Cover_path = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("ArtworkURL Set for device = " & MyUPnPDeviceName & ". Album = " & MyCurrentArtworkURL, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextArtworkURL As String
        Get
            NextArtworkURL = MyNextAlbumURI
        End Get
        Set(ByVal value As String)
            If Not (value.ToLower().StartsWith("http://") Or value.ToLower().StartsWith("https://") Or value.ToLower().StartsWith("file:")) Then
                Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                ' MyNextAlbumURI = "http://" & hs.GetIPAddress & HTTPPort & value
            End If
            If MyNextAlbumURI <> value And HSRefNextArt <> -1 Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("NextArtworkURL is setting HS Status for device - " & MyUPnPDeviceName & " with Value = " & value.ToString & " and HSRef = " & HSRefNextArt.ToString, LogType.LOG_TYPE_INFO)
                hs.DeviceVSP_ClearAll(HSRefNextArt, True)
                hs.DeviceVGP_ClearAll(HSRefNextArt, True)
                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = gsDefault
                Pair.Status = value
                hs.DeviceVSP_AddPair(HSRefNextArt, Pair)
                Dim GraphicsPair As VGPair
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = value
                GraphicsPair.Set_Value = gsDefault
                hs.DeviceVGP_AddPair(HSRefNextArt, GraphicsPair)
                hs.SetDeviceValueByRef(HSRefNextArt, gsDefault, True)
                hs.SetDeviceString(HSRefNextArt, value, True)
            End If
            MyNextAlbumURI = value
            NextLibEntry.Cover_path = value
            If PIDebuglevel > DebugLevel.dlEvents And gIOEnabled Then Log("NextArtworkURL Set for device = " & MyUPnPDeviceName & ". URL = " & MyNextAlbumURI, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property PlayerIconURL As String
        Get
            PlayerIconURL = MyIConURL
        End Get
        Set(ByVal value As String)
            MyIConURL = value
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerIconURL called for device = " & MyUPnPDeviceName.ToString & " and IConURL = " & MyIConURL.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property InstanceName As String
        Get
            InstanceName = MyUPnPDeviceName
        End Get
    End Property

    Public ReadOnly Property CurrentStreamTitle As String 'Implements MediaCommon.MusicAPI.CurrentStreamTitle
        Get
            'The title of the currently playing music stream (e.g. from an Internet music source)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CurrentStreamTitle called for device - " & MyUPnPDeviceName & " and Value = " & MyCurrentTrack.ToString, LogType.LOG_TYPE_INFO)
            CurrentStreamTitle = MyCurrentTrack
            'log( "CurrentStreamTitle called for device - " & MyUPnPDeviceName & "Value= " & MyCurrentTrack.ToString)
        End Get
    End Property

    Public ReadOnly Property CurrentTrackInfo As Object 'Implements MediaCommon.MusicAPI.CurrentTrackInfo
        Get
            'IITTrack
            'IWMPMedia3 	This returns the information on the currently playing track in iTunes (IIT Track) or Media Player (IWMPMedia3) API format.
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CurrentTrackInfo called for device - " & MyUPnPDeviceName & "Value= " & MyCurrentTrack.ToString, LogType.LOG_TYPE_INFO)
            Return Nothing
            'Dim MyCurrentTrackInfo As New IITTRack
            'CurrentTrackInfo = MyCurrentTrackInfo
        End Get
    End Property

    Public ReadOnly Property CurrentArtworkFile(Optional ByVal sPath As String = "") As String 'Implements MediaCommon.MusicAPI.CurrentArtworkFile
        Get
            ' implementation is slightly different and probably reverse. I return path here wheras in CurrentAlbumArtPath I download file and return path
            CurrentArtworkFile = MyCurrentArtworkURL
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CurrentArtworkFile called for device - " & MyUPnPDeviceName & "Value= " & MyCurrentArtworkURL.ToString, LogType.LOG_TYPE_INFO)
        End Get
    End Property

    Public ReadOnly Property CurrentTrack As String 'Implements MediaCommon.MusicAPI.CurrentTrack
        Get
            'Returns the name of the currently playing track.
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CurrentTrack called for device - " & MyUPnPDeviceName & ". Value= " & MyCurrentTrack.ToString)
            CurrentTrack = MyCurrentTrack
            'log( "CurrentTrack called for device - " & MyUPnPDeviceName & ". Value= " & MyCurrentTrack.ToString)
        End Get
    End Property

    Public ReadOnly Property CurrentAlbum As String 'Implements MediaCommon.MusicAPI.CurrentAlbum
        Get
            'Returns the album name of the currently playing track.
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CurrentAlbum called for device - " & MyUPnPDeviceName & "Value= " & MyCurrentAlbum.ToString)
            CurrentAlbum = MyCurrentAlbum
        End Get
    End Property

    Public ReadOnly Property CurrentAlbumArtPath As String
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CurrentAlbumArtPath called for device - " & MyUPnPDeviceName & " and CurrentAlbumArtPath = " & MyCurrentArtworkURL.ToString)
            CurrentAlbumArtPath = MyCurrentArtworkURL
        End Get
    End Property

    Public ReadOnly Property CurrentArtist As String 'Implements MediaCommon.MusicAPI.CurrentArtist
        Get
            'Returns the artist's name of the currently playing track.
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CurrentArtist called for device - " & MyUPnPDeviceName & "Value= " & MyCurrentArtist.ToString)
            CurrentArtist = MyCurrentArtist
        End Get
    End Property

    Public Function ToggleLoudnessState(ByVal Channel As String) As String
        ToggleLoudnessState = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim LoudnessState As Boolean = GetLoudnessState(Channel)
        If LoudnessState Then
            LoudnessState = False
        Else
            LoudnessState = True
        End If
        SetLoudnessState(Channel, LoudnessState)
    End Function

    Public Function SetLoudnessState(ByVal Channel As String, ByVal NewState As Boolean) As String
        SetLoudnessState = ""
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetLoudnessState called for zoneplayer = " & ZoneName & " with Channel = " & Channel & " and NewState = " & NewState & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function

        Try
            Dim InArg(2)            'InstanceID UI4
            Dim OutArg(0)           'Channel String
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
                InArg(2) = NewState ' Disired Loudness Boolean
            Else
                SetLoudnessState = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            RenderingControl.InvokeAction("SetLoudness", InArg, OutArg)
            SetLoudnessState = "OK"
        Catch ex As Exception
            'Log("ERROR in SetLoudnessState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetLoudnessState(ByVal Channel As String) As Boolean
        GetLoudnessState = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLoudnessState called for Device = " & MyUPnPDeviceName & " with values Channel = " & Channel & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(1)                'InstanceID UI4
            Dim OutArg(0)               'Channel String
            InArg(0) = 0
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetLoudness", InArg, OutArg)
            GetLoudnessState = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetLoudnessState for Device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Public Function SetMute(ByVal Channel As String, ByVal NewState As Boolean) As String
        SetMute = ""
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetMute called for zoneplayer = " & ZoneName & " with Channel = " & Channel & " and NewState = " & NewState & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)

            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
                InArg(2) = NewState
            Else
                SetMute = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If

            RenderingControl.InvokeAction("SetMute", InArg, OutArg)

            SetMute = "OK"
        Catch ex As Exception
            'Log("ERROR in SetMute for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ToggleMuteState(ByVal Channel As String) As String
        ToggleMuteState = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim MuteState As Boolean
        RCGetMute(Channel)
        If MyCurrentMuteState Then
            MuteState = False
        Else
            MuteState = True
        End If
        ToggleMuteState = SetMute(Channel, MuteState)
    End Function



    Public ReadOnly Property ShuffleStatus As String 'Implements MediaCommon.MusicAPI.ShuffleStatus
        Get
            'Returns the current shuffle mode as a string value: Shuffled, Ordered, Sorted, or Unknown
            'Dim CurrentShuffleState = MyControllerRef.GetPlayMode
            ShuffleStatus = "Ordered"
            Select Case UCase(MyCurrentPlayMode)
                Case "NORMAL"
                    ShuffleStatus = "Ordered"
                Case "SHUFFLE_NOREPEAT"
                    ShuffleStatus = "Shuffled"
                Case "REPEAT_ALL"
                    ShuffleStatus = "Ordered"
                Case "SHUFFLE"
                    ShuffleStatus = "Shuffled"
                Case Else
                    ShuffleStatus = "Ordered"
            End Select
        End Get
    End Property

    Public ReadOnly Property QueueShuffleStatus As String 'Implements MediaCommon.MusicAPI.ShuffleStatus
        Get
            'Returns the current shuffle mode as a string value: Shuffled, Ordered, Sorted, or Unknown
            'Dim CurrentShuffleState = MyControllerRef.GetPlayMode
            QueueShuffleStatus = "Ordered"
            If MyQueueShuffleState Then
                QueueShuffleStatus = "Shuffled"
            Else
                QueueShuffleStatus = "Ordered"
            End If
        End Get
    End Property


    Public Sub ShuffleToggle() 'Implements MediaCommon.MusicAPI.ShuffleToggle
        'Toggles through the 3 states for playlist shuffling: Shuffle, Order, Sort
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ShuffleToggle called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        AVTGetTransportSettings()
        Dim CurrentShuffleState = MyCurrentPlayMode
        Select Case UCase(CurrentShuffleState)
            Case "NORMAL"
                AVTSetPlayMode("SHUFFLE")
            Case "SHUFFLE NO REPEAT"
                AVTSetPlayMode("REPEAT_ALL")
            Case "REPEAT ALL"
                AVTSetPlayMode("NORMAL")
            Case "SHUFFLE"
                AVTSetPlayMode("SHUFFLE_NOREPEAT")
            Case "UNKNOWN"
                AVTSetPlayMode("NORMAL")
        End Select
        PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
    End Sub

    Public Property Shuffle As Integer 'Implements MediaCommon.MusicAPI.Shuffle
        Get '(Property Get) 	  	Short Integer 	Returns the current shuffle status: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            Shuffle = Shuffle_modes.Ordered
            Select Case UCase(MyCurrentPlayMode)
                Case "NORMAL"
                    Shuffle = Shuffle_modes.Ordered
                Case "SHUFFLE_NOREPEAT"
                    Shuffle = Shuffle_modes.Shuffled
                Case "REPEAT_ALL"
                    Shuffle = Shuffle_modes.Ordered
                Case "SHUFFLE"
                    Shuffle = Shuffle_modes.Shuffled
                Case Else
                    Shuffle = Shuffle_modes.Ordered
            End Select
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get Shuffle called for device - " & MyUPnPDeviceName & " with Value : " & Shuffle.ToString)
        End Get
        Set(ByVal value As Integer) '(Property Set) 	Short Integer 	  	Sets the shuffle status to the indicated value: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Shuffle set called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Dim RepeatState As repeat_modes = Repeat()
            Select Case value
                Case Shuffle_modes.Shuffled  ' Shuffled
                    If RepeatState = repeat_modes.repeat_all Then
                        AVTSetPlayMode("SHUFFLE")
                        MyCurrentPlayMode = "SHUFFLE" ' set the states here, prevents wrong setting when shuffle/repeat are set in one action
                    Else
                        AVTSetPlayMode("SHUFFLE_NOREPEAT")
                        MyCurrentPlayMode = "SHUFFLE_NOREPEAT"
                    End If
                Case Shuffle_modes.Ordered ' Ordered
                    If RepeatState = repeat_modes.repeat_all Then
                        AVTSetPlayMode("REPEAT_ALL")
                        MyCurrentPlayMode = "REPEAT_ALL"
                    Else
                        AVTSetPlayMode("NORMAL")
                        MyCurrentPlayMode = "NORMAL"
                    End If
                Case Shuffle_modes.Sorted ' Sorted
                    If RepeatState = repeat_modes.repeat_all Then
                        AVTSetPlayMode("REPEAT_ALL")
                        MyCurrentPlayMode = "REPEAT_ALL"
                    Else
                        AVTSetPlayMode("NORMAL")
                        MyCurrentPlayMode = "NORMAL"
                    End If
            End Select
        End Set
    End Property

    Public Property QueueShuffle As Integer 'Implements MediaCommon.MusicAPI.Shuffle
        Get '(Property Get) 	  	Short Integer 	Returns the current shuffle status: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            If MyQueueShuffleState Then
                QueueShuffle = Shuffle_modes.Shuffled
            Else
                QueueShuffle = Shuffle_modes.Ordered
            End If
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get QueueShuffle called for device - " & MyUPnPDeviceName & " with Value : " & Shuffle.ToString)
        End Get
        Set(ByVal value As Integer) '(Property Set) 	Short Integer 	  	Sets the shuffle status to the indicated value: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("QueueShuffle set called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Select Case value
                Case Shuffle_modes.Shuffled  ' Shuffled
                    MyQueueShuffleState = True
                    hs.SetDeviceValueByRef(HSRefQueueShuffle, ssShuffled, True)
                Case Shuffle_modes.Ordered ' Ordered
                    MyQueueShuffleState = False
                    hs.SetDeviceValueByRef(HSRefQueueShuffle, ssNoShuffle, True)
                Case Shuffle_modes.Sorted  ' Sorted
                    MyQueueShuffleState = False
                    hs.SetDeviceValueByRef(HSRefQueueShuffle, ssNoShuffle, True)
            End Select
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diQueueShuffle.ToString, MyQueueShuffleState)
        End Set
    End Property

    Public Sub ToggleRepeat()
        If Repeat() = 0 Then Repeat = 2 Else Repeat = 0
    End Sub

    Public Sub ToggleQueueRepeat()
        If QueueRepeat() = 0 Then QueueRepeat = 2 Else QueueRepeat = 0
    End Sub

    Public Sub ToggleShuffle()
        If Shuffle = Shuffle_modes.Ordered Then Shuffle = Shuffle_modes.Shuffled Else Shuffle = Shuffle_modes.Ordered
    End Sub

    Public Sub ToggleQueueShuffle()
        If QueueShuffle() = Shuffle_modes.Ordered Then QueueShuffle = Shuffle_modes.Shuffled Else QueueShuffle = Shuffle_modes.Ordered
    End Sub

    Public Property Repeat As repeat_modes 'Implements MediaCommon.MusicAPI.Repeat
        '(Enum) 	Returns the current repeat setting using the following Enum values:
        ' Public Enum repeat_modes
        '      repeat_off = 0
        '      repeat_one = 1
        '     repeat_all = 2
        ' End Enum
        Get
            If UCase(MyCurrentPlayMode) = "REPEAT_ALL" Or UCase(MyCurrentPlayMode) = "SHUFFLE" Then
                Repeat = repeat_modes.repeat_all
            Else
                Repeat = repeat_modes.repeat_off
            End If
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get Repeat called for device - " & MyUPnPDeviceName  & " with Value : " & Repeat.ToString)
        End Get
        Set(ByVal value As repeat_modes)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set Repeat called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString & " and MyPlayFromQueue = " & MyPlayFromQueue.ToString, LogType.LOG_TYPE_INFO)
            If value = repeat_modes.repeat_all Then
                Select Case UCase(MyCurrentPlayMode)
                    Case "NORMAL"
                        AVTSetPlayMode("REPEAT_ALL")
                        MyCurrentPlayMode = "REPEAT_ALL"
                    Case "SHUFFLE_NOREPEAT"
                        AVTSetPlayMode("SHUFFLE")
                        MyCurrentPlayMode = "SHUFFLE"
                End Select
            Else ' no repeat_off or repeat_one
                Select Case UCase(MyCurrentPlayMode)
                    Case "REPEAT_ALL"
                        AVTSetPlayMode("NORMAL")
                        MyCurrentPlayMode = "NORMAL"
                    Case "SHUFFLE"
                        AVTSetPlayMode("SHUFFLE_NOREPEAT")
                        MyCurrentPlayMode = "SHUFFLE_NOREPEAT"
                End Select
            End If
        End Set
    End Property

    Public Property QueueRepeat As repeat_modes 'Implements MediaCommon.MusicAPI.Repeat
        '(Enum) 	Returns the current repeat setting using the following Enum values:
        ' Public Enum repeat_modes
        '      repeat_off = 0
        '      repeat_one = 1
        '     repeat_all = 2
        ' End Enum
        Get
            If MyQueueRepeatState Then
                QueueRepeat = repeat_modes.repeat_all
            Else
                QueueRepeat = repeat_modes.repeat_off
            End If
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get Repeat called for device - " & MyUPnPDeviceName  & " with Value : " & Repeat.ToString)
        End Get
        Set(ByVal value As repeat_modes)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set QueueRepeat called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString & " and MyPlayFromQueue = " & MyPlayFromQueue.ToString, LogType.LOG_TYPE_INFO)
            If value = repeat_modes.repeat_all Then
                MyQueueRepeatState = True
                hs.SetDeviceValueByRef(HSRefQueueRepeat, rsRepeat, True)
            Else
                MyQueueRepeatState = False
                hs.SetDeviceValueByRef(HSRefQueueRepeat, rsnoRepeat, True)
            End If
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diQueueRepeat.ToString, MyQueueRepeatState)
        End Set
    End Property

    Public Sub Play() 'Implements MediaCommon.MusicAPI.Play
        'Starts the player playing the currently loaded HomeSeer playlist.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Play called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            If MyQueuePlayState = player_state_values.Paused Or MyCurrentPlayerState = player_state_values.Forwarding Or MyCurrentPlayerState = player_state_values.Rewinding Then
                CurrentPlayerState = player_state_values.Playing
                MyQueuePlayState = player_state_values.Playing
                MyTransportStateHasChanged = True
                MyTrackInfoHasChanged = False
                UpdateTransportState()
            Else
                If MyUPnPDeviceServiceType = "HST" Then
                    CurrentPlayerState = player_state_values.Playing
                    MyQueueDelay = 0
                End If
                CurrentPlayerState = player_state_values.Playing
                MyQueuePlayState = player_state_values.Playing
                MyTransportStateHasChanged = True
                MyTrackInfoHasChanged = False
                PlayFromQueue(1)
                UpdateTransportState()
                Exit Sub
            End If
        End If
        Try
            AVTPlay()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Play called for device - " & MyUPnPDeviceName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub Pause() 'Implements MediaCommon.MusicAPI.Pause
        'Toggles the state of the pause function of the player.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Pause called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            If MyQueuePlayState <> player_state_values.Stopped Then
                CurrentPlayerState = player_state_values.Paused
                MyQueuePlayState = player_state_values.Paused
            End If
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            UpdateTransportState()
            If MyUPnPDeviceServiceType = "HST" Then Exit Sub
        End If
        Try
            AVTPause()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Pause called for device - " & MyUPnPDeviceName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub PlayIfPaused() 'Implements MediaCommon.MusicAPI.PlayIfPaused
        'If the current state of the player is paused, the player will be resumed.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayIfPaused called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        ' states are Play Stop Pause Next Previous
        If MyPlayFromQueue Then
            If MyUPnPDeviceServiceType = "HST" Then MyCurrentPlayerState = player_state_values.Playing
            MyQueuePlayState = player_state_values.Playing
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            UpdateTransportState()
            If MyUPnPDeviceServiceType = "HST" Then Exit Sub
        End If
        Try
            AVTGetTransportInfo()
        Catch ex As Exception
        End Try
        If ConvertTransportStateToPlayerState(MyCurrentTransportState) <> player_state_values.Stopped Then
            Try
                AVTPlay()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub PauseIfPlaying() 'Implements MediaCommon.MusicAPI.PauseIfPlaying
        'If the current state of the player is playing, the player will be paused.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PauseIfPlaying called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            If MyQueuePlayState <> player_state_values.Stopped Then
                CurrentPlayerState = player_state_values.Paused
                MyQueuePlayState = player_state_values.Paused
            End If
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            UpdateTransportState()
            If MyUPnPDeviceServiceType = "HST" Then Exit Sub
        End If
        ' states are Play Stop Pause Next Previous
        Try
            AVTGetTransportInfo()
        Catch ex As Exception
        End Try
        If ConvertTransportStateToPlayerState(MyCurrentTransportState) = player_state_values.Playing Or ConvertTransportStateToPlayerState(MyCurrentTransportState) = player_state_values.Forwarding Or ConvertTransportStateToPlayerState(MyCurrentTransportState) = player_state_values.Rewinding Then
            Try
                AVTPause()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub ResumeFromPause() 'Implements MediaCommon.MusicAPI.PauseIfPlaying
        'If the current state of the player is playing, the player will be paused.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ResumeFromPause called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        ' states are Play Stop Pause Next Previous
        If MyPlayFromQueue Then
            If MyUPnPDeviceServiceType = "HST" Then MyCurrentPlayerState = player_state_values.Playing
            MyQueuePlayState = player_state_values.Playing
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            UpdateTransportState()
            If MyUPnPDeviceServiceType = "HST" Then Exit Sub
        End If
        Try
            AVTGetTransportInfo()
        Catch ex As Exception
        End Try
        If ConvertTransportStateToPlayerState(MyCurrentTransportState) = player_state_values.Paused Then
            Try
                AVTPlay()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub TogglePause()
        'Toggles the state of the pause function of the player.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TogglePause called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        Try
            If PlayState = "Pause" Then
                SetPlayState("Play")
            ElseIf PlayState = "Play" Then
                SetPlayState("Pause")
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TogglePause for device - " & MyUPnPDeviceName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub TogglePlay()
        'Toggles the state of the play function of the player.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TogglePlay called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        Try
            If PlayState = "Play" Then
                If SpeedIsConfigurable And CurrentPlayerSpeed <> 1 Then
                    AVTPlay(1) ' set regular speed
                Else
                    SetPlayState("Pause")
                End If
            Else
                SetPlayState("Play")
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TogglePlay for device - " & MyUPnPDeviceName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub




    Public Sub TrackNext() 'Implements MediaCommon.MusicAPI.TrackNext
        'Causes the player to jump to the next track in the playlist and begin playing it.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TrackNext called for Zone - " & MyUPnPDeviceName & " and MyQueuePlayState = " & MyQueuePlayState.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            MyQueueDelay = 0
            If NextElementPrefetchState = AVT_PrefetchState.PfSLoaded And NextAvTransportIsAvailable And UseNextAvTransport Then
                If UCase(AVTNext(0)) <> "OK" Then
                    PlayDBItem(MyNextQueueObjectID, MyNextQueueServerUDN, True)
                Else
                    MyCurrentQueueIndex = MyNextQueueIndex
                    MyCurrentQueueObjectID = MyNextQueueObjectID
                    MyCurrentQueueServerUDN = MyNextQueueServerUDN
                End If
                NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            Else
                NextQueue()
            End If
        Else
            Try
                AVTNext()
            Catch ex As Exception
            End Try
        End If

    End Sub

    Public Sub TrackPrev() 'Implements MediaCommon.MusicAPI.TrackPrev
        'Causes the player to start playing from the previous track in the playlist.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TrackPrev called for Zone - " & MyUPnPDeviceName & " and MyQueuePlayState = " & MyQueuePlayState.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            MyQueueDelay = 0
            PreviousQueue()
        Else
            Try
                AVTPrevious()
            Catch ex As Exception
            End Try
        End If

    End Sub

    Public Sub SkipToTrack(ByVal track_num As Integer) 'Implements MediaCommon.MusicAPI.SkipToTrack
        '(Overloaded) 	Integer 	  	Jumps the player to the track number in the current HomeSeer playlist given in the Integer parameter.  Track numbers less than 0 or greater than the number of entries in the playlist are ignored.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SkipToTrack(integer) called for device - " & MyUPnPDeviceName & " with value = " & track_num.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            PlayFromQueue(track_num, True, True)
            Exit Sub
        End If
        Try
            'PlayFromQueue("Q:")
            AVTSeek(AVT_SeekMode.TRACK_NR, Str(track_num + 1))
            PlayIfPaused()
        Catch ex As Exception
            Log("Error in SkipToTrack with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SkipToTrack(ByVal track_name As String) 'Implements MediaCommon.MusicAPI.SkipToTrack
        '(Overloaded) 	String 	  	Jumps the player to track matching the track name provided as the parameter.  If the track name does not match any entries in the current playlist, nothing happens.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SkipToTrack(String) called for device - " & MyUPnPDeviceName & " with value = " & track_name.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        ' dcor to be fixed
        Dim QueueInformation() As String = Nothing
        'QueueInformation = GetCurrentPlaylistTracks()
        If QueueInformation Is Nothing Then Exit Sub
        Dim Index As Integer
        For Index = 0 To UBound(QueueInformation)
            If QueueInformation(Index) = track_name Then
                ' we found it, use the index to start the track
                'PlayFromQueue("Q:")
                AVTSeek(AVT_SeekMode.TRACK_NR, Str(Index + 1)) ' Sonos starts at 1 not 0 as an Index
                PlayIfPaused()
                Exit Sub
            End If
        Next
        Log("SkipToTrack(String) called for device - " & MyUPnPDeviceName & " but couldn't find = " & track_name.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub SelectTrackInPlayList(TrackInfo As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SelectTrackInPlayList called for device - " & MyUPnPDeviceName & " with value = " & TrackInfo.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        ' this is in the form of Dim QItems As String() = Split(arplaylist(Index), ":;:-:")
        Try
            Dim TrackElements As String() = Split(TrackInfo, ":;:-:")
            Dim ObjectId = TrackElements(0)
            Dim ServerUDN = TrackElements(1)
            If ServerUDN = "" Then
                Log("Error in SelectTrackInPlayList for device - " & MyUPnPDeviceName & " no ServerUDN", LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            Dim QElement As myQueueElement = GetElementFromQueueByObjectId(ObjectId)
            If QElement Is Nothing Then
                Log("Error in SelectTrackInPlayList for device - " & MyUPnPDeviceName & " no Queue Element found for ObjectID = " & ObjectId.ToString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            'MyQueuePlayState = player_state_values.Playing
            MyCurrentQueueIndex = GetQueueIndex(ObjectId)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SelectTrackInPlayList found queueindex for device - " & MyUPnPDeviceName & " with value = " & MyCurrentQueueIndex.ToString, LogType.LOG_TYPE_INFO)
            MyNbrOfQueueItemsPlayed = 0
            MyPlayerWentThroughPlayState = False
            PlayDBItem(ObjectId, ServerUDN, True)
            MyQueuePlayState = player_state_values.Playing
            If (NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST" Then
                NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            End If
        Catch ex As Exception
            Log("Error in SelectTrackInPlayList for device - " & MyUPnPDeviceName & " with value = " & TrackInfo.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub StartPlay() 'Implements MediaCommon.MusicAPI.StartPlay
        'Like Play, this command starts the player playing the currently loaded HomeSeer playlist, but StartPlay always starts at playlist entry 0 (the beginning).
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("StartPlay called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper <> "ONLINE" Then Exit Sub
        If MyPlayFromQueue Then
            'MyCurrentPlayerState = player_state_values.Playing
            MyQueuePlayState = player_state_values.Playing
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            PlayFromQueue(1)
            UpdateTransportState()
            Exit Sub
        End If
        Try
            AVTPlay()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub StopPlay() 'Implements MediaCommon.MusicAPI.StopPlay
        'Stops the player.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("StopPlay called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        MyQueuePlayState = player_state_values.Stopped
        If MyPlayFromQueue Then
            CurrentPlayerState = player_state_values.Stopped
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = False
            UpdateTransportState()
            If MyUPnPDeviceServiceType = "HST" Then Exit Sub
        End If
        MyCurrentQueueIndex = 1
        MyNbrOfQueueItemsPlayed = 0
        If DeviceStatus.ToUpper = "ONLINE" Then
            Try
                AVTStop()
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public ReadOnly Property LibLoading As Boolean 'Implements MediaCommon.MusicAPI.LibLoading
        Get '(Property Get) 	  	Boolean 	When True is returned, the internal track library is in the process of being updated (loading).
            LibLoading = False
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LibLoading called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
            'PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
        End Get
    End Property

    Public Property PlayerMute As Boolean 'Implements MediaCommon.MusicAPI.PlayerMute
        Get ' (Property Get) Boolean Returns True if the player is muted, False if it is not.
            PlayerMute = MyCurrentMuteState
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get PlayerMute called for device - " & MyUPnPDeviceName & " State = " & PlayerMute.ToString)
        End Get
        Set(ByVal value As Boolean)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set PlayerMute called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            If value Then
                RCSetMute(True)
            Else
                RCSetMute(False)
            End If
        End Set
    End Property

    Public Property HasAnnouncementStarted As Boolean
        Get
            HasAnnouncementStarted = MyPlayerWentThroughPlayState
        End Get
        Set(value As Boolean)
            MyPlayerWentThroughPlayState = value
        End Set
    End Property


    Public Property Volume As Integer 'Implements MediaCommon.MusicAPI.Volume
        Get ' Returns the current volume setting of the player from 0 to 100.
            Volume = MyCurrentVolumeLevel
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get Volume called for device - " & MyUPnPDeviceName & " with value = " & Volume.ToString)
        End Get
        Set(ByVal value As Integer) '(Property Set) 	Integer 	  	Sets the volume of the player to the level indicated by the parameter, in the range 0-100.
            Dim Result As String
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set Volume called for device - " & MyUPnPDeviceName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            If Not VolumeIsConfigurable Then Exit Property
            Try
                Result = SetVolumeLevel("Master", value)
                If Result <> "OK" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Volume called for device - " & MyUPnPDeviceName & " but ended in Error: " & Result.ToString, LogType.LOG_TYPE_ERROR)
                End If
            Catch ex As Exception
                Log("Volume called for device - " & MyUPnPDeviceName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Set
    End Property

    Public Sub PlayMusic(NavigationString As String, ServerUDN As String, ByVal ClearPlayerQueue As Boolean, QueueAction As QueueActions)

        '    Public Sub PlayMusic(ByVal Artist As String, Optional ByVal Album As String = "", Optional ByVal PlayList As String = "", Optional ByVal Genre As String = "", Optional ByVal Track As String = "", Optional ByVal URL As String = "", Optional ByVal StartWithArtist As String = "", Optional ByVal StartWithTrack As String = "", Optional ByVal TrackMatch As String = "", Optional ByVal AddTrackToQueue As Boolean = False)
        'This causes the player to create a HomeSeer playlist matching the criteria provided and begin playing it.  
        'At least one parameter must be provided, although it may be delivered as a null string ("") if it is not desired to specify the artist.
        'Examples:
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("Phil Collins")                  Plays music by Phil Collins
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("", "", "", "Rock")             Plays music in the Rock genre
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("", "", "My Top Rated")   Plays music from the 'My Top Rated' playlist.
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "PlayMusic called for device - " & MyUPnPDeviceName & " with Artist=" & Artist & " and Album=" & Album & " and Playlist=" & PlayList & " and Genre=" & Genre & " and Track=" & Track & " and URL=" & URL & " and StartWithArtist=" & StartWithArtist & " and StartWithTrack=" & StartWithTrack & " and TrackMatch=" & TrackMatch & " and AddTrackToQueue = " & AddTrackToQueue.ToString)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayMusic called for device - " & MyUPnPDeviceName & " with NavigationString = " & NavigationString & " and ServerUDN = " & ServerUDN & " and ClearPlayerQueue = " & ClearPlayerQueue.ToString & " and QueueAction = " & QueueAction.ToString, LogType.LOG_TYPE_INFO)

        If ServerUDN = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & ". There is no Server assigned.", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        If Trim(NavigationString) = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & ". The NavigationString is empty.", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim ServerAPI As HSPI = MyReferenceToMyController.GetAPIByUDN(ServerUDN)
        If ServerAPI Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & ". ServerAPI could not be retrieved with ServerUDN = " & ServerUDN, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim ObjectNavigationParts As String() = Nothing

        Try
            ObjectNavigationParts = Split(NavigationString, ";--::")
            If ObjectNavigationParts Is Nothing Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & ". There are no ObjectNavigation parts found in NavigationString = " & NavigationString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            If ObjectNavigationParts.Count = 0 Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & ". There are no ObjectNavigation parts found in NavigationString = " & NavigationString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayMusic for device = " & MyUPnPDeviceName & " retrieving the ObjectNavigation parts in NavigationString = " & NavigationString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Dim ObjectParts As String() = Nothing
        Dim ObjectTitle As String = ""
        Dim ObjectID As String = ""
        ObjectParts = Split(ObjectNavigationParts(ObjectNavigationParts.Count - 1), ";::--") ' take the last one
        If ObjectParts IsNot Nothing Then
            ObjectTitle = ObjectParts(0)
            If ObjectParts.Count >= 2 Then
                ObjectID = ObjectParts(1)
            Else
                ObjectID = ""
            End If
        End If

        Dim ar As DBRecord() = Nothing
        Dim SomethingAdded As Boolean = False

        Try
            If ObjectID <> "" And ObjectTitle <> "" Then
                ar = ServerAPI.GetObjectFromServer(ObjectID, ObjectTitle)
                If ar IsNot Nothing Then
                    If ClearPlayerQueue Then ClearCurrentPlayList()
                    For Each arrecord As DBRecord In ar
                        AddObjectToCurrentPlaylist(arrecord.Id, ServerUDN, NavigationString, QueueAction)
                        SomethingAdded = True
                    Next
                    If SomethingAdded And QueueAction <> QueueActions.qaDontPlay Then Play()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Log("Error in PlayMusic retrieving the Object for device - " & MyUPnPDeviceName & " with NavigationString = " & NavigationString & ", ServerUDN = " & ServerUDN & ", ObjectID = " & ObjectID & ", ObjectTitle = " & ObjectTitle & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' we got here because the stored object ID is not valid anymore. Generate a warning
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in PlayMusic retrieving the Object for device - " & MyUPnPDeviceName & " with NavigationString = " & NavigationString & ", ServerUDN = " & ServerUDN & ", ObjectID = " & ObjectID & ", ObjectTitle = " & ObjectTitle, LogType.LOG_TYPE_WARNING)

        Try
            ' go look for it using the Titles stored in the Navigation Tree
            ar = ServerAPI.NavigateIntoContainerByTitle(NavigationString)
            SomethingAdded = False
            If ClearPlayerQueue Then ClearCurrentPlayList()
            If ar IsNot Nothing Then
                For Each arrecord As DBRecord In ar
                    AddObjectToCurrentPlaylist(arrecord.Id, ServerUDN, NavigationString, QueueAction)
                    SomethingAdded = True
                Next
            End If
            If SomethingAdded And QueueAction <> QueueActions.qaDontPlay Then Play()
        Catch ex As Exception
            Log("Error in PlayMusic retrieving the ObjectContainer for device - " & MyUPnPDeviceName & " with NavigationString = " & NavigationString & " and ServerUDN = " & ServerUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Function GetCurrentPlaylist() As track_desc() 'Implements MediaCommon.MusicAPI.GetCurrentPlaylist
        '  Array of track_desc 	
        'Returns the last HomeSeer playlist created when the player began playing whether by event action, control web page, or script command. 
        ' The return is an array of type track_desc, which is a class defined as follows:
        'Public Class track_desc
        '        Public name As String
        '        Public artist As String
        '        Public album As String
        '        Public length As String
        'End Class
        GetCurrentPlaylist = Nothing
        Dim Queue() As track_desc
        ReDim Queue(0)
        Dim TrackDescriptor As New track_desc
        TrackDescriptor.name = MyCurrentTrack  '"My Playlist Test Album" ' dcor to be fixed
        TrackDescriptor.artist = MyCurrentArtist '"My Playlist Test Artist"
        TrackDescriptor.length = MyCurrentTrackDuration.ToString
        TrackDescriptor.album = MyCurrentAlbum
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetCurrentPlaylist called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Queue(0) = TrackDescriptor
            GetCurrentPlaylist = Queue
        Catch ex As Exception

        End Try

    End Function

    Private Function CreateHSPlayerDevice(ByVal HSRef As Integer, ByVal DeviceName As String, NewDevice As Boolean) As Integer
        CreateHSPlayerDevice = -1
        Dim dv As Scheduler.Classes.DeviceClass
        Dim DevName As String = DeviceName
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSPlayerDevice called with device " & DeviceName & " with reference " & HSRef.ToString & " and NewDevice = " & NewDevice.ToString, LogType.LOG_TYPE_INFO)
        Try
            If HSRef = -1 Then
                HSRef = hs.NewDeviceRef("Player") '(DevName)
                Log("CreateHSPlayerDevice created device " & DevName & " with reference " & HSRef.ToString, LogType.LOG_TYPE_INFO)
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
                If HSRef <> -1 Then WriteStringIniFile("UPnP HSRef to UDN", HSRef, MyUDN)
                NewDevice = True
            Else
                ' Return HSRef
            End If
            dv = hs.GetDeviceByRef(HSRef)
            If NewDevice Then
                dv.Interface(hs) = sIFACE_NAME
                dv.Location2(hs) = tIFACE_NAME
                dv.InterfaceInstance(hs) = MainInstance
                dv.Location(hs) = DeviceName
                dv.Device_Type_String(hs) = HSDevices.Player.ToString
                dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
                dv.Address(hs) = "Player"
                Dim DT As New DeviceTypeInfo
                DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
                DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Root
                DT.Device_SubType_Description = "Media Controller Player Master Control"
                dv.DeviceType_Set(hs) = DT
                dv.Status_Support(hs) = True
                hs.SetDeviceString(HSRef, "_", False)
                ' dv.Image(hs) is set in myDeviceFinderCallback_DeviceFound after the image was downloaded
                ' This device is a child device, the parent being the root device for the entire security system. 
                ' As such, this device needs to be associated with the root (Parent) device.
                dvParent = hs.GetDeviceByRef(HSRefDevice)
                If dvParent.AssociatedDevices_Count(hs) < 1 Then
                    ' There are none added, so it is OK to add this one.
                    dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    Dim Found As Boolean = False
                    For Each ref As Integer In dvParent.AssociatedDevices(hs)
                        If ref = HSRef Then
                            Found = True
                            Exit For
                        End If
                    Next
                    If Not Found Then
                        dvParent.AssociatedDevice_Add(hs, HSRef)
                    Else
                        ' This is an error condition likely as this device's reference ID should not already be associated.
                    End If
                End If

                ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
                '   the child's associations.
                dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
                dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
                dv.Relationship(hs) = Enums.eRelationship.Child

                hs.DeviceVSP_ClearAll(HSRef, True)
                hs.DeviceVGP_ClearAll(HSRef, True)

                Dim MyIcon As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, "")
                If MyIcon <> "" Then
                    dv.Image(hs) = MyIcon
                    dv.ImageLarge(hs) = MyIcon
                End If

                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psPlay
                Pair.Status = "Play"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psStop
                Pair.Status = "Stop"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psPause
                Pair.Status = "Pause"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psPrevious
                Pair.Status = "Prev"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psNext
                Pair.Status = "Next"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psClearQueue
                Pair.Status = "Clear Queue"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)

                If VolumeIsConfigurable Then
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psVolDown
                    Pair.Status = "Vol - Dn"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 3
                    Pair.Render_Location.Column = 1
                    hs.DeviceVSP_AddPair(HSRef, Pair)

                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.Range
                    Pair.Value = psVolSlider
                    Pair.RangeStart = MyMinimumVolume
                    Pair.RangeEnd = MyMaximumVolume
                    Pair.RangeStatusPrefix = "Volume "
                    Pair.RangeStatusSuffix = "%"
                    Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
                    Pair.Render_Location.Row = 3
                    Pair.Render_Location.Column = 2
                    'hs.DeviceVSP_AddPair(HSRef, Pair)  ' two sliders on same device doesn't work, neither can I set the value

                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psVolUp
                    Pair.Status = "Vol - Up"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 3
                    Pair.Render_Location.Column = 2
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)

                    If BalanceIsConfigurable Then
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = psBalanceLeft
                        Pair.Status = "Bal - Left"
                        Pair.Render = Enums.CAPIControlType.Button
                        Pair.Render_Location.Row = 4
                        Pair.Render_Location.Column = 1
                        hs.DeviceVSP_AddPair(HSRef, Pair)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                        Pair.PairType = VSVGPairType.Range
                        Pair.Value = psBalanceSlider
                        Pair.RangeStart = 200
                        Pair.RangeEnd = 200 + MyMaximumVolume * 2 '400
                        Pair.RangeStatusPrefix = "Balance L <-> R "
                        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
                        Pair.Render_Location.Row = 4
                        Pair.Render_Location.Column = 2
                        'hs.DeviceVSP_AddPair(HSRef, Pair) ' two sliders on same device doesn't work, neither can I set the value

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = psBalanceRight
                        Pair.Status = "Bal - Right"
                        Pair.Render = Enums.CAPIControlType.Button
                        Pair.Render_Location.Row = 4
                        Pair.Render_Location.Column = 2
                        hs.DeviceVSP_AddPair(HSRef, Pair)
                    End If

                End If

                If MuteIsConfigurable Then
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psMute
                    Pair.Status = "Mute"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 3
                    Pair.Render_Location.Column = 3
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                End If

                If LoudnessIsConfigurable Then
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psLoudness
                    Pair.Status = "Loudness"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 4
                    Pair.Render_Location.Column = 3
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                End If

                If PlayModeisConfigurable Then
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psShuffle
                    Pair.Status = "Shuffle"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 5
                    Pair.Render_Location.Column = 1
                    hs.DeviceVSP_AddPair(HSRef, Pair)

                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = psRepeat
                    Pair.Status = "Repeat"
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = 5
                    Pair.Render_Location.Column = 2
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                End If

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psQueueShuffle
                Pair.Status = "QueueShuffle"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 6
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = psQueueRepeat
                Pair.Status = "QueueRepeat"
                Pair.Render = Enums.CAPIControlType.Button
                Pair.Render_Location.Row = 6
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

            End If

            Dim DeviceType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diDeviceType.ToString, "")
            HSRefTrack = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackHSRef.ToString, -1)
            If HSRefTrack = -1 Then
                HSRefTrack = CreateHSDevice(DeviceInfoIndex.diTrackHSRef, "Track")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackHSRef.ToString, HSRefTrack)
            End If
            HSRefArtist = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diArtistHSRef.ToString, -1)
            If HSRefArtist = -1 Then
                HSRefArtist = CreateHSDevice(DeviceInfoIndex.diArtistHSRef, "Artist")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diArtistHSRef.ToString, HSRefArtist)
            End If
            HSRefAlbum = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diAlbumHSRef.ToString, -1)
            If HSRefAlbum = -1 Then
                HSRefAlbum = CreateHSDevice(DeviceInfoIndex.diAlbumHSRef, "Album")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diAlbumHSRef.ToString, HSRefAlbum)
            End If
            If DeviceType <> "HST" Then

                HSRefTrackLength = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackLengthHSRef.ToString, -1)
                If HSRefTrackLength = -1 Then
                    HSRefTrackLength = CreateHSDevice(DeviceInfoIndex.diTrackLengthHSRef, "Track Length")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackLengthHSRef.ToString, HSRefTrackLength)
                End If
                HSRefTrackPos = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackPosHSRef.ToString, -1)
                If HSRefTrackPos = -1 Then
                    HSRefTrackPos = CreateHSDevice(DeviceInfoIndex.diTrackPosHSRef, "Track Position")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackPosHSRef.ToString, HSRefTrackPos)
                End If
                HSRefTrackDescr = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackDescrHSRef.ToString, -1)
                If HSRefTrackDescr = -1 Then
                    HSRefTrackDescr = CreateHSDevice(DeviceInfoIndex.diTrackDescrHSRef, "Track Desc")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diTrackDescrHSRef.ToString, HSRefTrackDescr)
                End If
                HSRefGenre = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diGenreHSRef.ToString, -1)
                If HSRefGenre = -1 Then
                    HSRefGenre = CreateHSDevice(DeviceInfoIndex.diGenreHSRef, "Genre")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diGenreHSRef.ToString, HSRefGenre)
                End If
            End If
            HSRefArt = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diArtHSRef.ToString, -1)
            If HSRefArt = -1 Then
                HSRefArt = CreateHSDevice(DeviceInfoIndex.diArtHSRef, "Art")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diArtHSRef.ToString, HSRefArt)
                ArtworkURL = NoArtPath
            End If
            If NextAvTransportIsAvailable Then
                HSRefNextTrack = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diNextTrackHSRef.ToString, -1)
                If HSRefNextTrack = -1 Then
                    HSRefNextTrack = CreateHSDevice(DeviceInfoIndex.diNextTrackHSRef, "Next Track")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diNextTrackHSRef.ToString, HSRefNextTrack)
                End If
                HSRefNextArtist = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtistHSRef.ToString, -1)
                If HSRefNextArtist = -1 Then
                    HSRefNextArtist = CreateHSDevice(DeviceInfoIndex.diNextArtistHSRef, "Next Artist")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtistHSRef.ToString, HSRefNextArtist)
                End If
                HSRefNextAlbum = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diNextAlbumHSRef.ToString, -1)
                If HSRefNextAlbum = -1 Then
                    HSRefNextAlbum = CreateHSDevice(DeviceInfoIndex.diNextAlbumHSRef, "Next Album")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diNextAlbumHSRef.ToString, HSRefNextAlbum)
                End If
                HSRefNextArt = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtHSRef.ToString, -1)
                If HSRefNextArt = -1 Then
                    HSRefNextArt = CreateHSDevice(DeviceInfoIndex.diNextArtHSRef, "Next Art")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtHSRef.ToString, HSRefNextArt)
                    NextArtworkURL = NoArtPath
                End If
            End If
            HSRefPlayState = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diPlayStateHSRef.ToString, -1)
            If HSRefPlayState = -1 Then
                HSRefPlayState = CreateHSDevice(DeviceInfoIndex.diPlayStateHSRef, "State")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diPlayStateHSRef.ToString, HSRefPlayState)
            End If
            If VolumeIsConfigurable Then
                HSRefVolume = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diVolumeHSRef.ToString, -1)
                If HSRefVolume = -1 Then
                    HSRefVolume = CreateHSDevice(DeviceInfoIndex.diVolumeHSRef, "Volume")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diVolumeHSRef.ToString, HSRefVolume)
                End If
                hs.SetDeviceValueByRef(MyVolumeHSRef, MyCurrentVolumeLevel, True)
                If BalanceIsConfigurable Then
                    HSRefBalance = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diBalanceHSRef.ToString, -1)
                    If HSRefBalance = -1 Then
                        HSRefBalance = CreateHSDevice(DeviceInfoIndex.diBalanceHSRef, "Balance")
                        WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diBalanceHSRef.ToString, HSRefBalance)
                        hs.SetDeviceValueByRef(HSRefBalance, 0, True)
                    End If
                End If
            End If
            If MuteIsConfigurable Then
                HSRefMute = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diMuteHSRef.ToString, -1)
                If HSRefMute = -1 Then
                    HSRefMute = CreateHSDevice(DeviceInfoIndex.diMuteHSRef, "Mute")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diMuteHSRef.ToString, HSRefMute)
                End If
                If MyCurrentMuteState Then
                    hs.SetDeviceValueByRef(MyMuteHSRef, msMuted, True)
                Else
                    hs.SetDeviceValueByRef(MyMuteHSRef, msUnmuted, True)
                End If
            End If
            If LoudnessIsConfigurable Then
                HSRefLoudness = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diLoudnessHSRef.ToString, -1)
                If HSRefLoudness = -1 Then
                    HSRefLoudness = CreateHSDevice(DeviceInfoIndex.diLoudnessHSRef, "Loudness")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diLoudnessHSRef.ToString, HSRefLoudness)
                End If
                If MyCurrentLoudness Then
                    hs.SetDeviceValueByRef(MyLoudnessHSRef, lsLoudnessOn, True)
                Else
                    hs.SetDeviceValueByRef(MyLoudnessHSRef, lsLoudnessOff, True)
                End If
            End If
            If PlayModeisConfigurable Then
                HSRefRepeat = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diRepeatHSRef.ToString, -1)
                If HSRefRepeat = -1 Then
                    HSRefRepeat = CreateHSDevice(DeviceInfoIndex.diRepeatHSRef, "Repeat")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diRepeatHSRef.ToString, HSRefRepeat)
                    hs.SetDeviceValueByRef(HSRefRepeat, rsnoRepeat, True)
                End If
                HSRefShuffle = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diShuffleHSRef.ToString, -1)
                If HSRefShuffle = -1 Then
                    HSRefShuffle = CreateHSDevice(DeviceInfoIndex.diShuffleHSRef, "Shuffle")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diShuffleHSRef.ToString, HSRefShuffle)
                    hs.SetDeviceValueByRef(HSRefShuffle, ssNoShuffle, True)
                End If
            End If
            HSRefQueueRepeat = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diQueueRepeatHSRef.ToString, -1)
            If HSRefQueueRepeat = -1 Then
                HSRefQueueRepeat = CreateHSDevice(DeviceInfoIndex.diQueueRepeatHSRef, "QueueRepeat")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diQueueRepeatHSRef.ToString, HSRefQueueRepeat)
                hs.SetDeviceValueByRef(HSRefQueueRepeat, rsnoRepeat, True)
            End If
            HSRefQueueShuffle = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diQueueShuffleHSRef.ToString, -1)
            If HSRefQueueShuffle = -1 Then
                HSRefQueueShuffle = CreateHSDevice(DeviceInfoIndex.diQueueShuffleHSRef, "QueueShuffle")
                WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diQueueShuffleHSRef.ToString, HSRefQueueShuffle)
                hs.SetDeviceValueByRef(HSRefQueueShuffle, ssNoShuffle, True)
            End If
            If BrightnessIsConfigurable Then
                'MyBrightnessHSRef = CreateNewHSDevices("Brightness", Enums.dvMISC.STATUS_ONLY + Enums.dvMISC.SHOW_VALUES)
                'CreateVolumePairs(MyBrightnessHSRef)
                'hs.SetDeviceString(MyBrightnessHSRef, "Brightness = " & MyCurrentBrightness.ToString, True)
                'hs.SetDeviceValueByRef(MyBrightnessHSRef, MyCurrentBrightness, True)
            End If
            If ColorTemperatureIsConfigurable Then
                'MyColorHSRef = CreateNewHSDevices("Color", Enums.dvMISC.STATUS_ONLY + Enums.dvMISC.SHOW_VALUES)
                'CreateVolumePairs(MyColorHSRef)
                'hs.SetDeviceString(MyColorHSRef, "Color = " & MyCurrentColorTemperature.ToString, True)
                'hs.SetDeviceValueByRef(MyColorHSRef, MyCurrentColorTemperature, True)
            End If
            If ContrastIsConfigurable Then
                'MyContrastHSRef = CreateNewHSDevices("Contrast", Enums.dvMISC.STATUS_ONLY + Enums.dvMISC.SHOW_VALUES)
                'CreateVolumePairs(MyContrastHSRef)
                'hs.SetDeviceString(MyContrastHSRef, "Contrast = " & MyCurrentContrast.ToString, True)
                'hs.SetDeviceValueByRef(MyContrastHSRef, MyCurrentContrast, True)
            End If
            If SharpnessIsConfigurable Then
                'MySharpnessHSRef = CreateNewHSDevices("Sharpness", Enums.dvMISC.STATUS_ONLY + Enums.dvMISC.SHOW_VALUES)
                'CreateVolumePairs(MySharpnessHSRef)
                'hs.SetDeviceString(MySharpnessHSRef, "Sharpness = " & MyCurrentSharpness.ToString, True)
                'hs.SetDeviceValueByRef(MySharpnessHSRef, MyCurrentSharpness, True)
            End If
            If SpeedIsConfigurable Then
                HSRefSpeed = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diSpeedHSRef.ToString, -1)
                If HSRefSpeed = -1 Then
                    HSRefSpeed = CreateHSDevice(DeviceInfoIndex.diSpeedHSRef, "Speed")
                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diSpeedHSRef.ToString, HSRefSpeed)
                End If
            End If
            Return HSRef
        Catch ex As Exception
            Log("Error in CreateHSPlayerDevice with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function CreateHSServiceDevice(ByVal HSRef As Integer, DevName As String) As Integer
        CreateHSServiceDevice = -1
        Dim dv As Scheduler.Classes.DeviceClass
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        HSRef = GetIntegerIniFile(MyUDN, "di" & DevName & "HSCode", -1)
        If HSRef <> -1 Then
            Return HSRef
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSServiceDevice called for Device = " & MyUPnPDeviceName & " and DeviceFunction = " & DevName & " and HSRef = " & HSRef.ToString, LogType.LOG_TYPE_INFO)
        ' No HomeSeer Device Code yet
        Try
            HSRef = hs.NewDeviceRef(DevName)
            Log("CreateHSServiceDevice created device " & DevName & " with HSRef = " & HSRef.ToString, LogType.LOG_TYPE_INFO)
            ' Force HomeSeer to save changes to devices and events so we can find our new device
            hs.SaveEventsDevices()
            If HSRef <> -1 Then WriteStringIniFile("UPnP HSRef to UDN", HSRef, MyUDN)
            WriteIntegerIniFile(MyUDN, "di" & DevName & "HSCode", HSRef)
        Catch ex As Exception
            Log("Error in CreateHSServiceDevice getting free HS code for UPnPDevice = " & MyUPnPDeviceName & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If HSRef = -1 Then
            Log("Error in CreateHSServiceDevice getting a new HS device for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            dv = hs.GetDeviceByRef(HSRef)
            dv.Interface(hs) = sIFACE_NAME
            dv.InterfaceInstance(hs) = MainInstance
            dv.Location(hs) = MyUPnPDeviceName
            dv.Location2(hs) = tIFACE_NAME
            dv.Device_Type_String(hs) = DevName
            dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
            dv.Address(hs) = "Service"
            Dim DT As New DeviceTypeInfo
            DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
            DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Root
            DT.Device_SubType_Description = "Media Controller Service Control"
            dv.DeviceType_Set(hs) = DT
            dv.Status_Support(hs) = True
            'hs.SetDeviceString(HSRef, "_", False)
            ' This device is a child device, the parent being the root device for the entire security system. 
            ' As such, this device needs to be associated with the root (Parent) device.
            dvParent = hs.GetDeviceByRef(HSRefDevice)
            If dvParent.AssociatedDevices_Count(hs) < 1 Then
                ' There are none added, so it is OK to add this one.
                dvParent.AssociatedDevice_Add(hs, HSRef)
            Else
                Dim Found As Boolean = False
                For Each ref As Integer In dvParent.AssociatedDevices(hs)
                    If ref = HSRef Then
                        Found = True
                        Exit For
                    End If
                Next
                If Not Found Then
                    dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    ' This is an error condition likely as this device's reference ID should not already be associated.
                End If
            End If
            ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
            '   the child's associations.
            dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
            dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
            dv.Relationship(hs) = Enums.eRelationship.Child
            hs.DeviceVSP_ClearAll(HSRef, True)
            Dim MyIcon As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, "")
            If MyIcon <> "" Then
                dv.Image(hs) = MyIcon
                dv.ImageLarge(hs) = MyIcon
            End If
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, MyIConURL)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSServiceDevice updated HS info for device " & MyUPnPDeviceName & " - " & DevName & " with HSRef " & HSRef.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in CreateHSServiceDevice with error : " & ex.Message & " and Ref = " & HSRef.ToString, LogType.LOG_TYPE_ERROR)
        End Try
        CreateHSServiceDevice = HSRef
    End Function

    Public Sub CreateStateButtons(HSRef As Integer)
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlay
        Pair.Status = "Play"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psStop
        Pair.Status = "Stop"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPause
        Pair.Status = "Pause"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPrevious
        Pair.Status = "Prev"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 2
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psNext
        Pair.Status = "Next"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 2
        Pair.Render_Location.Column = 2
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psShuffle
        Pair.Status = "Shuffle"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 2
        Pair.Render_Location.Column = 3
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRepeat
        Pair.Status = "Repeat"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 2
        Pair.Render_Location.Column = 4
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psVolDown
        Pair.Status = "Dn"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 3
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.Range
        Pair.Value = psVolSlider
        Pair.RangeStart = 0
        Pair.RangeEnd = MyMaximumVolume
        Pair.RangeStatusPrefix = "Volume "
        Pair.RangeStatusSuffix = "%"
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Pair.Render_Location.Row = 3
        Pair.Render_Location.Column = 2
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psVolUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 3
        Pair.Render_Location.Column = 3
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psMute
        Pair.Status = "Mute"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 3
        Pair.Render_Location.Column = 4
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psBalanceLeft
        Pair.Status = "Left"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 4
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.Value = psBalanceSlider
        Pair.RangeStart = -100
        Pair.RangeEnd = 100
        Pair.RangeStatusPrefix = "Balance L <-> R "
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Pair.Render_Location.Row = 4
        Pair.Render_Location.Column = 2
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psBalanceRight
        Pair.Status = "Right"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 4
        Pair.Render_Location.Column = 3
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psLoudness
        Pair.Status = "Loudness"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 4
        Pair.Render_Location.Column = 4
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psClearQueue
        Pair.Status = "ClearQueue"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 4
        Pair.Render_Location.Column = 5
        hs.DeviceVSP_AddPair(HSRef, Pair)

        If Not MyPossibleFFSpeeds Is Nothing Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psFF
            Pair.Status = "FF"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Row = 5
            Pair.Render_Location.Column = 1
            hs.DeviceVSP_AddPair(HSRef, Pair)
        End If
        If Not MyPossibleREWSpeeds Is Nothing Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psREW
            Pair.Status = "REW"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Row = 5
            Pair.Render_Location.Column = 2
            hs.DeviceVSP_AddPair(HSRef, Pair)
        End If

    End Sub

    Private Sub CreateOnOffTogglePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOff
        Pair.Status = "Off"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOn
        Pair.Status = "On"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpToggle
        Pair.Status = "Toggle"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)
    End Sub

    Private Sub CreateActivateDeactivateButtons(HSRef As Integer, AddWOL As Boolean)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsDeactivate
        Pair.Status = "Deactivate"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivate
        Pair.Status = "Activate"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        If AddWOL Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = dsWOL
            Pair.Status = "WOL"
            Pair.Render = Enums.CAPIControlType.Button
            hs.DeviceVSP_AddPair(HSRef, Pair)
        End If


        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsDeactivated
        Pair.Status = "Deactivated"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Dim GraphicsPair As New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "NOKBtn.png"
        GraphicsPair.Set_Value = dsDeactivated
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOffLine
        Pair.Status = "Activated Offline"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "PartialOKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOffLine
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivateOnLine
        Pair.Status = "Activated Online"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivateOnLine
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOnLineUnregistered
        Pair.Status = "Activated Online Unregistered"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOnLineUnregistered
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsBuildingDB
        Pair.Status = "Building DB"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsBuildingDB
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

    End Sub

    Private Sub CreateContentDeviceButtons(HSRef As Integer)

        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsBuildDB
        Pair.Status = "Build DB"
        Pair.Render = Enums.CAPIControlType.Button
        'hs.DeviceVSP_AddPair(HSRef, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsBuildingDB
        Pair.Status = "Building DB"
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsDeactivated
        Pair.Status = "Deactivated"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Dim GraphicsPair As New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "NOKBtn.png"
        GraphicsPair.Set_Value = dsDeactivated
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOffLine
        Pair.Status = "Activated Offline"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "PartialOKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOffLine
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivateOnLine
        Pair.Status = "Activated Online"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivateOnLine
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOnLineUnregistered
        Pair.Status = "Activated Online Unregistered"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOnLineUnregistered
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsBuildingDB
        Pair.Status = "Building DB"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsBuildingDB
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

    End Sub

    Private Function CreateHSDevice(ByVal DeviceType As DeviceInfoIndex, DevString As String) As Integer
        CreateHSDevice = -1
        Dim dv As Scheduler.Classes.DeviceClass
        Dim DevName As String = DevString 'ZoneName & " - " & DevString
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        Dim HSRef As Integer = -1
        Try
            HSRef = hs.NewDeviceRef(DevName)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSDevice for Device = " & MyUPnPDeviceName & " created deviceType " & DeviceType.ToString & " with Ref " & HSRef.ToString, LogType.LOG_TYPE_INFO)
            ' Force HomeSeer to save changes to devices and events so we can find our new device
            hs.SaveEventsDevices()
            If HSRef <> -1 Then WriteStringIniFile("UPnP HSRef to UDN", HSRef, MyUDN)
            dv = hs.GetDeviceByRef(HSRef)
            dv.Interface(hs) = sIFACE_NAME
            dv.InterfaceInstance(hs) = MainInstance
            dv.Location2(hs) = tIFACE_NAME
            dv.Location(hs) = MyUPnPDeviceName
            'dv.MISC_Set(hs, Enums.dvMISC.HIDDEN)
            dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
            Dim DT As New DeviceTypeInfo
            DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
            Dim MyIcon As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, "")
            If MyIcon <> "" Then
                dv.Image(hs) = MyIcon
                dv.ImageLarge(hs) = MyIcon
            End If
            Select Case DeviceType
                Case DeviceInfoIndex.diTrackHSRef
                    dv.Device_Type_String(hs) = "Track"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Track
                    DT.Device_SubType_Description = "Player Track Name"
                    dv.Address(hs) = "S09"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Track"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Track", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diNextTrackHSRef
                    dv.Device_Type_String(hs) = "Next Track"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Next Track Name"
                    dv.Address(hs) = "S13"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Track"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Track", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diArtistHSRef
                    dv.Device_Type_String(hs) = "Artist"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Artist
                    DT.Device_SubType_Description = "Player Artist Name"
                    dv.Address(hs) = "S11"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Artist"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Artist", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diNextArtistHSRef
                    dv.Device_Type_String(hs) = "Next Artist"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Next Artist Name"
                    dv.Address(hs) = "S15"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Artist"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Artist", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)
                Case DeviceInfoIndex.diAlbumHSRef
                    dv.Device_Type_String(hs) = "Album"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Album
                    DT.Device_SubType_Description = "Player Album Name"
                    dv.Address(hs) = "S10"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Album"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Album", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)
                Case DeviceInfoIndex.diNextAlbumHSRef
                    dv.Device_Type_String(hs) = "Next Album"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Next Album Name"
                    dv.Address(hs) = "S14"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Album"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Album", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)
                Case DeviceInfoIndex.diArtHSRef
                    dv.Device_Type_String(hs) = "Art"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player AlbumArt URL"
                    dv.Address(hs) = "S16"
                    CreateArtImagePairs(HSRef)
                    hs.SetDeviceString(HSRef, NoArtPath, True)
                Case DeviceInfoIndex.diNextArtHSRef
                    dv.Device_Type_String(hs) = "Next Art"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player AlbumArt URL"
                    dv.Address(hs) = "S17"
                    CreateArtImagePairs(HSRef)
                    hs.SetDeviceString(HSRef, NoArtPath, True)
                Case DeviceInfoIndex.diPlayStateHSRef
                    dv.Device_Type_String(hs) = "State"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status
                    DT.Device_SubType_Description = "Player State"
                    dv.Address(hs) = "S01"
                    CreateStatePairs(HSRef)
                    hs.SetDeviceValueByRef(HSRef, player_state_values.Stopped, True)
                Case DeviceInfoIndex.diVolumeHSRef
                    dv.Device_Type_String(hs) = "Volume"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Volume
                    DT.Device_SubType_Description = "Player Volume"
                    dv.Address(hs) = "S02"
                    CreateVolumePairs(HSRef)
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                Case DeviceInfoIndex.diMuteHSRef
                    dv.Device_Type_String(hs) = "Mute"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Mute State"
                    dv.Address(hs) = "S06"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diMuteHSRef)
                Case DeviceInfoIndex.diLoudnessHSRef
                    dv.Device_Type_String(hs) = "Loudness"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Loudness State"
                    dv.Address(hs) = "S07"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diLoudnessHSRef)
                Case DeviceInfoIndex.diBalanceHSRef
                    dv.Device_Type_String(hs) = "Balance"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Balance State"
                    dv.Address(hs) = "S03"
                    CreateSliderPairs(HSRef, DeviceInfoIndex.diBalanceHSRef)
                    hs.SetDeviceValueByRef(HSRef, 100, False)
                Case DeviceInfoIndex.diTrackLengthHSRef
                    dv.Device_Type_String(hs) = "Track Length"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Track Length"
                    dv.Address(hs) = "S21"
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                    Select Case MyHSTrackPositionFormat
                        Case HSSTrackLengthSettings.TLSSeconds
                            hs.SetDeviceString(HSRef, "0", True)
                        Case HSSTrackLengthSettings.TLSHoursMinutesSeconds
                            hs.SetDeviceString(HSRef, "00:00:00", True)
                    End Select
                Case DeviceInfoIndex.diTrackPosHSRef
                    dv.Device_Type_String(hs) = "Track Position"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Track Position"
                    dv.Address(hs) = "S08"
                    CreateSliderPairs(HSRef, DeviceInfoIndex.diTrackPosHSRef)
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                    Select Case MyHSTrackPositionFormat
                        Case HSSTrackPositionSettings.TPSSeconds
                            hs.SetDeviceString(HSRef, "0", True)
                        Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                            hs.SetDeviceString(HSRef, "00:00:00", True)
                        Case HSSTrackPositionSettings.TPSPercentage
                            hs.SetDeviceString(HSRef, "0", True)
                    End Select
                Case DeviceInfoIndex.diTrackDescrHSRef
                    dv.Device_Type_String(hs) = "Track Descr"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Track Descriptor"
                    dv.Address(hs) = "S20"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Descr"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Descr", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)
                Case DeviceInfoIndex.diRepeatHSRef
                    dv.Device_Type_String(hs) = "Repeat"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Repeat
                    DT.Device_SubType_Description = "Player Repeat State"
                    dv.Address(hs) = "S04"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diRepeatHSRef)
                    hs.SetDeviceValueByRef(HSRef, rsnoRepeat, True)
                Case DeviceInfoIndex.diShuffleHSRef
                    dv.Device_Type_String(hs) = "Shuffle"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Shuffle
                    DT.Device_SubType_Description = "Player Shuffle State"
                    dv.Address(hs) = "S05"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diShuffleHSRef)
                    hs.SetDeviceValueByRef(HSRef, ssNoShuffle, True)
                Case DeviceInfoIndex.diGenreHSRef
                    dv.Device_Type_String(hs) = "Genre"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Genre
                    DT.Device_SubType_Description = "Player Track Genre"
                    dv.Address(hs) = "S19"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault
                    Pair.Status = "No Genre"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Genre", True)
                    hs.SetDeviceValueByRef(HSRef, gsDefault, True)
                Case DeviceInfoIndex.diSpeedHSRef
                    dv.Device_Type_String(hs) = "REW / FF"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Player Rewind / Fast Forward"
                    dv.Address(hs) = "S22"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = gsDefault + 1 ' this is the normal speed
                    Pair.Status = "Normal"
                    Pair.Render = Enums.CAPIControlType.Button
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceValueByRef(HSRef, gsDefault + 1, True)
                    Dim Index As Integer = 1
                    If MyPossibleFFSpeeds IsNot Nothing Then
                        For Each FFSpeed As Integer In MyPossibleFFSpeeds
                            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                            Pair.PairType = VSVGPairType.SingleValue
                            Pair.Value = gsDefault + FFSpeed
                            Pair.Status = "FF" & Index.ToString
                            Index = Index + 1
                            Pair.Render = Enums.CAPIControlType.Button
                            hs.DeviceVSP_AddPair(HSRef, Pair)
                        Next
                    End If
                    Index = 1
                    If MyPossibleREWSpeeds IsNot Nothing Then
                        For Each RewSpeed As Integer In MyPossibleREWSpeeds
                            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                            Pair.PairType = VSVGPairType.SingleValue
                            Pair.Value = gsDefault + RewSpeed
                            Pair.Status = "REW" & Index.ToString
                            Index = Index + 1
                            Pair.Render = Enums.CAPIControlType.Button
                            hs.DeviceVSP_AddPair(HSRef, Pair)
                        Next
                    End If
                Case DeviceInfoIndex.diQueueRepeatHSRef
                    dv.Device_Type_String(hs) = "QueueRepeat"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Repeat
                    DT.Device_SubType_Description = "Player QueueRepeat State"
                    dv.Address(hs) = "S23"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diQueueRepeatHSRef)
                    hs.SetDeviceValueByRef(HSRef, rsnoRepeat, True)
                Case DeviceInfoIndex.diQueueShuffleHSRef
                    dv.Device_Type_String(hs) = "QueueShuffle"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Shuffle
                    DT.Device_SubType_Description = "Player QueueShuffle State"
                    dv.Address(hs) = "S24"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diQueueShuffleHSRef)
                    hs.SetDeviceValueByRef(HSRef, ssNoShuffle, True)
            End Select

            dv.DeviceType_Set(hs) = DT
            dv.Status_Support(hs) = True
            ' This device is a child device, the parent being the root device for the entire security system. 
            ' As such, this device needs to be associated with the root (Parent) device.
            dvParent = hs.GetDeviceByRef(HSRefDevice)
            If dvParent.AssociatedDevices_Count(hs) < 1 Then
                ' There are none added, so it is OK to add this one.
                dvParent.AssociatedDevice_Add(hs, HSRef)
            Else
                Dim Found As Boolean = False
                For Each ref As Integer In dvParent.AssociatedDevices(hs)
                    If ref = HSRef Then
                        Found = True
                        Exit For
                    End If
                Next
                If Not Found Then
                    dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    ' This is an error condition likely as this device's reference ID should not already be associated.
                End If
            End If

            ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
            '   the child's associations.
            dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
            dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
            dv.Relationship(hs) = Enums.eRelationship.Child
            hs.SaveEventsDevices()
            Return HSRef
        Catch ex As Exception
            Log("Error in CreateHSDevice with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub CreateVolumePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        ' add a Down button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpDown
        Pair.Status = "Down"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)


        ' add Volume Slider
        Pair = New VSPair(ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.Value = vpSlider
        Pair.RangeStart = MyMinimumVolume
        Pair.RangeEnd = MyMaximumVolume
        Pair.RangeStatusPrefix = "Volume "
        Pair.RangeStatusSuffix = "%"
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add an Up button

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

    End Sub

    Private Sub CreateOnOffTogglePairs(HSRef As Integer, DeviceType As DeviceInfoIndex)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Dim GraphicsPair As VGPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOff
        Pair.Status = "Off"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOn
        Pair.Status = "On"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpToggle
        Pair.Status = "Toggle"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Select Case DeviceType
            Case DeviceInfoIndex.diMuteHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = msMuted
                Pair.Status = "Muted"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "muted.gif".ToLower
                GraphicsPair.Set_Value = msMuted
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = msUnmuted
                Pair.Status = "Unmuted"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "unmuted.gif"
                GraphicsPair.Set_Value = msUnmuted
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diLoudnessHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = lsLoudnessOn
                Pair.Status = "Loudness On"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "loudness.gif"
                GraphicsPair.Set_Value = lsLoudnessOn
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = lsLoudnessOff
                Pair.Status = "Loudness Off"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "Loudnessoff.gif"
                GraphicsPair.Set_Value = lsLoudnessOff
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diRepeatHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsnoRepeat
                Pair.Status = "No Repeat"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatoff.gif"
                GraphicsPair.Set_Value = rsnoRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsRepeat
                Pair.Status = "Repeat all"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatall.gif"
                GraphicsPair.Set_Value = rsRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diShuffleHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssNoShuffle
                Pair.Status = "Ordered"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "ordered.gif"
                GraphicsPair.Set_Value = ssNoShuffle
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssShuffled
                Pair.Status = "Shuffled"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "shuffled.gif"
                GraphicsPair.Set_Value = ssShuffled
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diQueueRepeatHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsnoRepeat
                Pair.Status = "No Repeat"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatoff.gif"
                GraphicsPair.Set_Value = rsnoRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsRepeat
                Pair.Status = "Repeat all"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatall.gif"
                GraphicsPair.Set_Value = rsRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diQueueShuffleHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssNoShuffle
                Pair.Status = "Ordered"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "ordered.gif"
                GraphicsPair.Set_Value = ssNoShuffle
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssShuffled
                Pair.Status = "Shuffled"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "shuffled.gif"
                GraphicsPair.Set_Value = ssShuffled
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
        End Select
    End Sub

    Private Sub CreateStatePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Dim GraphicsPair As VGPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psStop
        Pair.Status = "Stop"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlay
        Pair.Status = "Play"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPause '103
        Pair.Status = "Pause"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Dim DeviceType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diDeviceType.ToString, "")
        If DeviceType <> "HST" And SpeedIsConfigurable Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psREW
            Pair.Status = "REW"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Column = 1
            Pair.Render_Location.Row = 2
            hs.DeviceVSP_AddPair(HSRef, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psFF
            Pair.Status = "FF"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Column = 2
            Pair.Render_Location.Row = 2
            hs.DeviceVSP_AddPair(HSRef, Pair)
        End If

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlaying
        Pair.Status = "Playing"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "playing.gif"
        GraphicsPair.Set_Value = psPlaying
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psStopped
        Pair.Status = "Stopped"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "stopped.gif"
        GraphicsPair.Set_Value = psStopped
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPaused
        Pair.Status = "Paused"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "paused.jpg"
        GraphicsPair.Set_Value = psPaused
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psTransitioning
        Pair.Status = "Transitioning"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "paused.jpg"
        GraphicsPair.Set_Value = psTransitioning
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRewind
        Pair.Status = "Rewind"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "paused.jpg"
        GraphicsPair.Set_Value = psRewind
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psFastForward
        Pair.Status = "Fast Forward"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "paused.jpg"
        GraphicsPair.Set_Value = psFastForward
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

    End Sub

    Private Sub CreateSliderPairs(HSRef As Integer, deviceType As DeviceInfoIndex)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

        ' add a Dow button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpDown
        Pair.Status = "Down"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add Volume Slider
        Pair = New VSPair(ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Select Case deviceType
            Case DeviceInfoIndex.diBalanceHSRef
                Pair.Value = vpSlider
                Pair.RangeStart = -(MyMaximumVolume)
                Pair.RangeEnd = MyMaximumVolume
                Pair.RangeStatusPrefix = "Balance L (-100) <-> R (+100) : "
                'Pair.RangeStatusSuffix = "%"
            Case DeviceInfoIndex.diTrackPosHSRef
                Pair.Value = vpSlider
                Pair.RangeStart = 0
                Pair.RangeEnd = 200
                'Pair.RangeStatusPrefix = "Volume "
                'Pair.RangeStatusSuffix = "%"

        End Select
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add an Up button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)
    End Sub

    Private Sub CreateArtImagePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = gsDefault
        Pair.Status = NoArtPath
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Dim GraphicsPair As VGPair
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = NoArtPath
        GraphicsPair.Set_Value = gsDefault
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
    End Sub





    Public Sub SetHSMainState()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetHSMainState called for device - " & MyUPnPDeviceName & " and Current HSRef = " & HSRefDevice & " and AdminStateActive = " & MyAdminStateActive.ToString & " and DeviceStatus = " & DeviceStatus.ToString, LogType.LOG_TYPE_INFO)
        If HSRefDevice = -1 Then Exit Sub ' we don't have a code yet
        If MyAdminStateActive Then
            If UCase(DeviceStatus) = "ONLINE" Then
                hs.SetDeviceValueByRef(HSRefDevice, dsActivateOnLine, True)
            Else
                hs.SetDeviceValueByRef(HSRefDevice, dsActivatedOffLine, True)
            End If
        Else
            hs.SetDeviceValueByRef(HSRefDevice, dsDeactivated, True)
        End If
        If HSRefRemote = -1 Or HSRefRemote = HSRefDevice Then Exit Sub
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetHSMainState called for device - " & MyUPnPDeviceName & " and Current Remote HSRef = " & HSRefRemote & " and AdminStateActive = " & MyAdminStateActive.ToString & " and DeviceStatus = " & DeviceStatus.ToString, LogType.LOG_TYPE_INFO)
        Dim MyRemoteServiceAdminState As Boolean = GetBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminStateRemote.ToString, False)
        If MyAdminStateActive Then
            If UCase(DeviceStatus) = "ONLINE" Then
                Dim RemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                If (RemoteType.ToUpper = "SONYIRRC" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKETPIN" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKET" Or RemoteType.ToUpper = "LG") And Not GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                    hs.SetDeviceValueByRef(HSRefRemote, dsActivatedOnLineUnregistered, True)
                Else
                    hs.SetDeviceValueByRef(HSRefRemote, dsActivateOnLine, True)
                End If
            Else
                hs.SetDeviceValueByRef(HSRefRemote, dsActivatedOffLine, True)
            End If
        Else
            hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
        End If
    End Sub

    Public Sub SetHSRemoteState()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetHSRemoteState called for device - " & MyUPnPDeviceName & " and Current Remote HSRef = " & HSRefRemote & " and AdminStateActive = " & MyAdminStateActive.ToString & " and DeviceStatus = " & DeviceStatus.ToString & " and RemoteServiceActive = " & MyRemoteServiceActive.ToString, LogType.LOG_TYPE_INFO)
        If HSRefRemote = -1 Then Exit Sub
        Dim MyRemoteServiceAdminState As Boolean = GetBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminStateRemote.ToString, False)
        If MyAdminStateActive And MyRemoteServiceAdminState Then
            If UCase(DeviceStatus) = "ONLINE" Then
                Dim RemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                If (RemoteType.ToUpper = "SONYIRRC" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKETPIN" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKET" Or RemoteType.ToUpper = "LG") And Not GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                    hs.SetDeviceValueByRef(HSRefRemote, dsActivatedOnLineUnregistered, True)
                Else
                    hs.SetDeviceValueByRef(HSRefRemote, dsActivateOnLine, True)
                End If
            Else
                hs.SetDeviceValueByRef(HSRefRemote, dsActivatedOffLine, True)
            End If
        Else
            hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
        End If
    End Sub

    Public ReadOnly Property DeviceState As Integer
        Get
            If MyAdminStateActive Then
                If UCase(DeviceStatus) = "ONLINE" Then
                    Dim RemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                    If (RemoteType.ToUpper = "SONYIRRC" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKETPIN" Or RemoteType.ToUpper = "SAMSUNGWEBSOCKET" Or RemoteType.ToUpper = "LG") And Not GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                        DeviceState = 3
                    Else
                        DeviceState = 2
                    End If
                Else
                    DeviceState = 1
                End If
            Else
                DeviceState = 0
            End If
        End Get
    End Property

    Public Sub SetAdministrativeState(Active As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetAdministrativeState called for device - " & MyUPnPDeviceName & " and Active = " & Active.ToString, LogType.LOG_TYPE_INFO)
        Try
            If Active Then
                MyAdminStateActive = True
                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminState.ToString, True)
                If MyUPnPDeviceServiceType = "HST" Then
                    DeviceStatus = "Online"
                    SetHSMainState()
                    Exit Sub
                ElseIf MyUPnPDeviceServiceType = "DIAL" Then
                    If DeviceStatus.ToUpper <> "ONLINE" Then
                        Connect("uuid:" & MyUDN)
                        If MyUPnPDevice IsNot Nothing Then
                            If MyUPnPDevice.Alive Then 'MyReferenceToMyController.CheckDeviceIsOnLine(MyIPAddress) Then
                                If MyUPnPDeviceManufacturer.ToUpper.IndexOf("ROKU") <> -1 Then
                                    If RokuRetrieveDIALAppList(MyUPnPDevice.Location) Then '"http://" & MyIPAddress & ":" & MyIPPort) Then
                                        DeviceStatus = "Online"
                                        If HSRefRemote = -1 Then CreateHSRokuRemoteButtons(False)
                                    End If
                                Else
                                    If RetrieveDIALAppList(MyUPnPDevice.ApplicationURL) Then '"http://" & MyIPAddress & ":" & MyIPPort) Then
                                        DeviceStatus = "Online"
                                        If HSRefRemote = -1 Then CreateHSDIALRemoteButtons(False)
                                    End If
                                End If
                            End If
                        Else
                            DeviceStatus = "Offline"
                        End If
                    End If
                    SetHSMainState()
                    Exit Sub
                ElseIf MyUPnPDeviceServiceType = "RCR" Then
                    ' this device is created in extractallservices procedure
                    If DeviceStatus.ToUpper <> "ONLINE" Then
                        Connect("uuid:" & MyUDN)
                    End If
                    WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminState.ToString, True)
                    SetHSMainState()
                    Exit Sub
                ElseIf MyUPnPDeviceServiceType = "PMR" Then
                    WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminState.ToString, True)
                    WriteBooleanIniFile("Message Service by UDN", MyUDN, True)
                    'hs.SetDeviceValueByRef(HSRefMessage, dsActivateOnLine, True)
                    MyMessageServiceActive = True
                    Connect("uuid:" & MyUDN)
                    Exit Sub
                End If
                If DeviceStatus.ToUpper <> "ONLINE" And MyUPnPDevice IsNot Nothing Then 'MyReferenceToMyController.CheckDeviceIsOnLine(MyIPAddress) Then
                    If MyUPnPDevice.Alive Then
                        Connect("uuid:" & MyUDN)
                    End If
                End If
                SetHSMainState()
            Else ' Deactivate
                MyAdminStateActive = False
                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminState.ToString, False)
                If MyUPnPDeviceServiceType = "HST" Then
                    ' will have to add here what needs to be added if any
                    StopPlay()
                    DeviceStatus = "Offline"
                    SetHSMainState()
                ElseIf MyUPnPDeviceServiceType = "DIAL" Then
                    If DeviceStatus <> "Offline" Then
                        Disconnect(True)
                        DeviceStatus = "Offline"
                        SetHSMainState()
                    End If
                    Exit Sub
                ElseIf MyUPnPDeviceServiceType = "RCR" Then
                    If DeviceStatus <> "Offline" Then
                        Disconnect(True)
                        DeviceStatus = "Offline"
                        SetHSMainState()
                    End If
                    Exit Sub
                ElseIf MyUPnPDeviceServiceType = "PMR" Then
                    WriteBooleanIniFile("Message Service by UDN", MyUDN, False)
                    MyMessageServiceActive = False
                    'hs.SetDeviceValueByRef(HSRefMessage, dsDeactivated, True)
                    If DeviceStatus <> "Offline" Then
                        Disconnect(True)
                    Else
                        SetHSMainState()
                    End If
                    Exit Sub
                Else
                    If DeviceStatus <> "Offline" Then
                        If MyUPnPDeviceServiceType = "DMR" Then StopPlay()
                        Disconnect(True)
                    Else
                        SetHSMainState()
                    End If
                End If
            End If
        Catch ex As Exception
            Log("Error in SetAdministrativeState for device - " & MyUPnPDeviceName & " and Active = " & Active.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SetAdministrativeStateRemote(Activate As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetAdministrativeStateRemote called for device - " & MyUPnPDeviceName & " and Activate = " & Activate.ToString, LogType.LOG_TYPE_INFO)
        Dim RemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        If Activate Then
            If RemoteType.ToUpper = "SONYIRRC" Then
                SonySetupRemoteInfo()
                If HSRefRemote <> -1 Then
                    Dim RegisteredSetting As String = ""
                    RegisteredSetting = GetStringIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, "")
                    If RegisteredSetting = "" Then
                        ' never stored before, store now
                        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                        MyRemoteServiceActive = False
                    Else
                        MyRemoteServiceActive = True
                    End If
                End If
            ElseIf RemoteType.ToUpper = "SAMSUNGWEBSOCKETPIN" Then
                If SamsungActivateRemote() Then
                    If HSRefRemote <> -1 Then
                        Dim RegisteredSetting As String = ""
                        RegisteredSetting = GetStringIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, "")
                        If RegisteredSetting = "" Then
                            ' never stored before, store now
                            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                            MyRemoteServiceActive = False
                        Else
                            MyRemoteServiceActive = True
                        End If
                    End If
                End If
            End If
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminStateRemote.ToString, True)
        Else
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diAdminStateRemote.ToString, False)
        End If
        SetHSRemoteState()
    End Sub


    Public ReadOnly Property DeviceOnLine As Boolean
        Get
            DeviceOnLine = UCase(DeviceStatus) = "ONLINE"
        End Get
    End Property

    Public Property DeviceStatus As String
        Get
            DeviceStatus = MyDeviceStatus
        End Get
        Set(value As String)
            If value <> MyDeviceStatus Then
                If value.ToUpper = "ONLINE" Then
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Online)
                Else
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Offline)
                End If
            End If
            MyDeviceStatus = value
        End Set
    End Property

    Public Sub DirectConnect(ByVal pDevice As MyUPnPDevice)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect called for " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If pDevice Is Nothing Then
            Log("Error in DirectConnect. pDevice is Nothing", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Try
            MyUDN = pDevice.UniqueDeviceName
            MyUDN = Replace(MyUDN, "uuid:", "")
        Catch ex As Exception
            Log("Error in DirectConnect retrieving Unique Device Name with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If pDevice.ModelNumber IsNot Nothing Then
            MyUPnPModelNumber = pDevice.ModelNumber
        End If

        MyDocumentURL = pDevice.Location

        Try
            MyIConURL = pDevice.IconURL("image/png", 200, 200, 16) 'image/png image/x-png image/tiff image/bmp image/pjpeg image/jpeg
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect for Device = " & MyUPnPDeviceName & " found IconURL = " & MyIConURL, LogType.LOG_TYPE_INFO)
            If MyIConURL <> "" Then ' changed v3.0.0.32
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect for Device = " & MyUPnPDeviceName & " checking for file = " & CurrentAppPath & "/html" & URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_INFO)
                If Not File.Exists(CurrentAppPath & "dcor tralala" & "/html" & URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png") Then
                    Dim IConImage As Image
                    IConImage = GetPicture(MyIConURL)
                    If IConImage IsNot Nothing Then
                        Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                        MyIConURL = URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png"
                        Dim SuccesfullSave As Boolean = False
                        SuccesfullSave = hs.WriteHTMLImage(IConImage, FileArtWorkPath & "PlayerIcon_" & MyUDN & ".png", True)
                        If Not SuccesfullSave Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DirectConnect for Device = " & MyUPnPDeviceName & " had error storing Icon at " & FileArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_ERROR)
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect for Device = " & MyUPnPDeviceName & " stored Icon at " & FileArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_INFO)
                        End If
                        IConImage.Dispose()
                    End If
                Else
                    MyIConURL = URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png"
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect for Device = " & MyUPnPDeviceName & " found Icon already stored at " & MyIConURL, LogType.LOG_TYPE_INFO)
                End If

                Dim HSRef As Integer = HSRefDevice
                If HSRef <> -1 Then
                    Dim dv As Scheduler.Classes.DeviceClass
                    dv = hs.GetDeviceByRef(HSRef)
                    dv.Image(hs) = MyIConURL
                    dv.ImageLarge(hs) = MyIConURL
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DirectConnect for Device = " & MyUPnPDeviceName & " added Image = " & MyIConURL & " for HSRef = " & HSRef.ToString, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, MyIConURL)
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DirectConnect. Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        DeviceStatus = "Online"
        MyUPnPDevice = pDevice

        pDevice.AddHandlers(Me)

        MyMacAddress = MyReferenceToMyController.LocalMACAddress
        ExtractAllServices(pDevice)
        If pDevice.HasChildren Then
            For Each Child As MyUPnPDevice In pDevice.Children
                If Child IsNot Nothing Then
                    ExtractAllServices(Child)
                End If
            Next
        End If
        SetHSMainState()

    End Sub

    Public Function Connect(ByVal inUDN As String) As String
        Connect = ""
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Connect called for UPnPDevice - " & MyUPnPDeviceName & " and inUDN = " & inUDN, LogType.LOG_TYPE_INFO)
        If DeviceStatus.ToUpper = "ONLINE" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Connect called for UPnPDevice - " & MyUPnPDeviceName & " which is already on-line with UDN = " & inUDN & " and DeviceStatus = " & DeviceStatus.ToString, LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        Dim MyDevice As MyUPnPDevice = Nothing
        Try
            MyDevice = MySSDPDevice.Item(inUDN)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Connect for UPnPDevice - " & MyUPnPDeviceName & " and UDN = " & inUDN & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Connect = "Failed"
            Exit Function
        End Try
        Try
            If MyDevice IsNot Nothing Then
                MyUDN = Replace(inUDN, "uuid:", "")
                'If piDebuglevel > DebugLevel.dlEvents Then Log("Connect called for UPnPDevice - " & MyUPnPDeviceName & " with UDN = " & inUDN & " and Device to find = " & Val(DeviceToFind), LogType.LOG_TYPE_INFO)
                Connect = "OK"
                ConnectUPnPDevice = False
                myDeviceFinderCallback_DeviceFound(MyDevice)
            Else
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Connect for UPnPDevice - " & MyUPnPDeviceName & ". No device found with UDN = " & inUDN, LogType.LOG_TYPE_ERROR)
                Connect = "Failed"
            End If
        Catch ex As Exception
            Log("Error in Connect for UPnPDevice - " & MyUPnPDeviceName & " and UDN = " & inUDN & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub Disconnect(full As Boolean)
        Log("Disconnect: Disconnected from UPnPDevice - " & MyUPnPDeviceName & " and CurrentState = " & DeviceStatus & " Full = " & full.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Sub ' already disconnected
        DeviceStatus = "Offline"
        DestroyObjects(full)
        SetHSMainState()
    End Sub

    Private Sub DestroyObjects(Full As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroyObjects called for UPnPDevice - " & MyUPnPDeviceName & " with Full = " & Full.ToString, LogType.LOG_TYPE_INFO)
        DeviceStatus = "Offline"
        Try
            If AVTransport IsNot Nothing Then AVTransport.RemoveCallback()
        Catch ex As Exception
        End Try
        AVTransport = Nothing
        Try
            If RenderingControl IsNot Nothing Then RenderingControl.RemoveCallback()
        Catch ex As Exception
        End Try
        RenderingControl = Nothing
        Try
            If ContentDirectory IsNot Nothing Then ContentDirectory.RemoveCallback()
        Catch ex As Exception
        End Try
        ContentDirectory = Nothing
        Try
            If ConnectionManager IsNot Nothing Then ConnectionManager.RemoveCallback()
        Catch ex As Exception
        End Try
        ConnectionManager = Nothing
        Try
            If RemoteUIServer IsNot Nothing Then RemoteUIServer.RemoveCallback()
        Catch ex As Exception
        End Try
        RemoteUIServer = Nothing
        Try
            If SonyPartyService IsNot Nothing Then SonyPartyService.RemoveCallback()
        Catch ex As Exception
        End Try
        SonyPartyService = Nothing
        Try
            If RemoteControlService IsNot Nothing Then RemoteControlService.RemoveCallback()
        Catch ex As Exception
        End Try
        RemoteControlService = Nothing
        Try
            If MessageBoxService IsNot Nothing Then MessageBoxService.RemoveCallback()
        Catch ex As Exception
        End Try
        Try
            If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                Select Case GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                    Case "Onkyo"
                        OnkyoCloseTCPConnection()
                    Case "Samsungiapp", "SamsungWebSocket", "SamsungWebSocketPIN"
                        SamsungCloseTCPConnection(False)
                    Case "DIAL"
                    Case "SonyIRRC"
                    Case "LG"
                        LGCloseWebSocket()
                End Select
            End If
        Catch ex As Exception
        End Try
        MessageBoxService = Nothing
        Try
            MyUPnPDevice.Dispose(True)
        Catch ex As Exception
        End Try
        MyUPnPDevice = Nothing
    End Sub

    Private Sub myDeviceFinderCallback_DeviceFound(ByVal pDevice As MyUPnPDevice)

        If pDevice.ModelNumber IsNot Nothing Then
            MyUPnPModelNumber = pDevice.ModelNumber
        End If

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Device Finder Callback received for UPnPDevice " & MyUPnPDeviceName & " with device name = " & pDevice.UniqueDeviceName, LogType.LOG_TYPE_INFO)

        If (Mid(pDevice.UniqueDeviceName, 1, 12) = "uuid:RINCON_") And Not SonosDeviceIn Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Device Finder Call Back for UPnPDevice = " & MyUPnPDeviceName & " found Sonos device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_INFO)
            Exit Sub
        End If

        If pDevice.UniqueDeviceName <> "uuid:" & MyUDN Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning Device Finder Call Back for UPnPDevice = " & MyUPnPDeviceName & " received response from wrong device. If you are running XP, go to HS setup and turn off the 'HomeSeer is Discoverable Using UPNP'. The UDN that responded =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_WARNING)
            Exit Sub
        End If

        MyDocumentURL = pDevice.Location

        Try
            MyIConURL = pDevice.IconURL("image/png", 200, 200, 16)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Device Finder CallBack for Device = " & MyUPnPDeviceName & " found IconURL = " & MyIConURL, LogType.LOG_TYPE_INFO)
            If MyIConURL <> "" Then ' changed v3.0.0.32
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Device Finder CallBack for Device = " & MyUPnPDeviceName & " checking for file = " & CurrentAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_INFO)
                If Not File.Exists(CurrentAppPath & "dcor tralala" & "\html" & URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png") Then
                    Dim IConImage As Image
                    IConImage = GetPicture(MyIConURL)
                    If Not IConImage Is Nothing Then
                        Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                        MyIConURL = URLArtWorkPath & "PlayerIcon_" & MyUDN.ToString & ".png"
                        Dim SuccesfullSave As Boolean = False
                        SuccesfullSave = hs.WriteHTMLImage(IConImage, FileArtWorkPath & "PlayerIcon_" & MyUDN.ToString & ".png", True)
                        If Not SuccesfullSave Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in myDeviceFinderCallback_DeviceFound for Device = " & MyUPnPDeviceName & " had error storing Icon at " & FileArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_ERROR)
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("myDeviceFinderCallback_DeviceFound for Device = " & MyUPnPDeviceName & " stored Icon at " & FileArtWorkPath & "PlayerIcon_" & MyUDN & ".png", LogType.LOG_TYPE_INFO)
                        End If
                        'IConImage.Save(hs.GetAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & MyUPnPModelNumber.ToString & ".png", ImageFormat)
                        IConImage.Dispose()
                    End If
                Else
                    MyIConURL = URLArtWorkPath & "PlayerIcon_" & MyUDN & ".png"
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("myDeviceFinderCallback_DeviceFound for Device = " & MyUPnPDeviceName & " found Icon already stored at " & MyIConURL, LogType.LOG_TYPE_INFO)
                End If
                Dim HSRef As Integer = HSRefDevice
                If HSRef <> -1 Then
                    Dim dv As Scheduler.Classes.DeviceClass
                    dv = hs.GetDeviceByRef(HSRef)
                    dv.Image(hs) = MyIConURL
                    dv.ImageLarge(hs) = MyIConURL
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diDeviceIConURL.ToString, MyIConURL)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Device Finder Call Back  for Device = " & MyUPnPDeviceName & " added Image = " & MyIConURL, LogType.LOG_TYPE_INFO)
                End If

            End If
        Catch ex As Exception
            Log("Error in Device Finder Call Back for Device = " & MyUPnPDeviceName & ". Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            MyUDN = pDevice.UniqueDeviceName
            MyUDN = Replace(MyUDN, "uuid:", "")
        Catch ex As Exception
            Log("Error in myDeviceFinderCallback_DeviceFound for device = " & MyUPnPDeviceName & " retrieving Unique Device Name with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        DeviceStatus = "Online"
        MyUPnPDevice = pDevice

        pDevice.AddHandlers(Me)

        Try
            Dim ICon As Object
            ICon = pDevice.IconURL("image/jpeg", 200, 200, 16)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("'            IconIRL = " & ICon.ToString, LogType.LOG_TYPE_INFO)
            MyIConURL = ICon.ToString
        Catch ex As Exception
        End Try
        Try
            MyUPnPDeviceManufacturer = pDevice.ManufacturerName
        Catch ex As Exception
        End Try

        MyMacAddress = MyReferenceToMyController.LocalMACAddress
        ExtractAllServices(pDevice)
        If pDevice.HasChildren Then
            For Each Child As MyUPnPDevice In pDevice.Children
                If Child IsNot Nothing Then
                    ExtractAllServices(Child)
                End If
            Next
        End If
        SetHSMainState()

    End Sub

    Public Sub DeviceLostCallback()
        Try
            ' Device is Removed or Lost
            Log("UPnPDevice " & MyUPnPDeviceName & " has been disconnected from the network in DeviceLostCallback.", LogType.LOG_TYPE_WARNING)
            DeviceStatus = "Offline"
            DestroyObjects(False)
            ConnectUPnPDevice = True
            SetHSMainState()
        Catch ex As Exception
            Log("Error in DeviceLostCallback for UPnPDevicee = " & MyUPnPDeviceName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeviceAliveCallBack()
        Try
            ' Device is Removed or Lost
            Log("UPnPDevice " & MyUPnPDeviceName & " just became alive on the network in DeviceAliveCallBack.", LogType.LOG_TYPE_WARNING)
            'Reachable() let the timer add the device back so not to block the thread
        Catch ex As Exception
            Log("Error in DeviceAliveCallBack for UPnPDevicee = " & MyUPnPDeviceName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ExtractAllServices(pDevice As MyUPnPDevice)
        Dim Services As MyUPnPServices = Nothing
        Try
            Services = pDevice.Services
            If Services Is Nothing Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found no services", LogType.LOG_TYPE_WARNING)
                Exit Sub
            End If
            Try
                ProcessServiceDocument(pDevice)
            Catch ex As Exception
                Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " processing the Service Document with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found " & pDevice.Services.Count.ToString & " Services", LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Try
                    For Each objService As MyUPnPService In Services
                        If objService IsNot Nothing Then
                            Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found Service ID = " & objService.Id.ToString, LogType.LOG_TYPE_INFO)
                        End If
                    Next
                Catch ex As Exception
                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " discovering the Service IDs with error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            Try
                For Each objService As MyUPnPService In Services
                    Dim ObjectserviceID As String = ""
                    Dim objectserviceType As String = ""
                    Try
                        If Not objService Is Nothing Then
                            ObjectserviceID = objService.Id
                            objectserviceType = objService.ServiceType
                        End If
                    Catch ex As Exception
                        Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " extracting the Service ID with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If ObjectserviceID = "urn:upnp-org:serviceId:AVTransport" Or ObjectserviceID = "urn:schemas-upnp-org:service:AVTransport" Then
                        ' this was added for the wireless dock which is different from other zones
                        AVTransport = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found AVTransport", LogType.LOG_TYPE_INFO)
                        If Not AVTransport Is Nothing Then
                            Dim AllowedValueString As String = SearchServiceDocument(MyDocumentURL, ObjectserviceID, "Play", "Speed")
                            If AllowedValueString <> "" Then
                                Dim MyPossibleSpeedSettings As String()
                                MyPossibleSpeedSettings = GetXMLValues(AllowedValueString, "allowedValue")
                                If Not MyPossibleSpeedSettings Is Nothing Then ProcessSpeedSettings(MyPossibleSpeedSettings)
                            End If

                            HSRefPlayer = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diPlayerHSRef.ToString, -1)
                            If HSRefPlayer = -1 Then
                                HSRefPlayer = CreateHSPlayerDevice(HSRefPlayer, MyUPnPDeviceName, True)
                                If HSRefPlayer <> -1 Then
                                    WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diPlayerHSRef.ToString, HSRefPlayer)
                                End If
                            Else
                                CreateHSPlayerDevice(HSRefPlayer, MyUPnPDeviceName, False)
                            End If

                            'AVTGetCurrentTransportActions()
                            'AVTGetDeviceCapabilities()
                            'AVTGetTransportSettings()
                            'AVTX_DLNA_GetBytePositionInfo()

                            Try
                                AVTransport.AddCallback(myAVTransportCallback)
                                Log("AvTransportCallback added for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                            Catch ex As Exception
                                Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " adding Transport Call Back with error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                            'AVTGetMediaInfo(0)     ' removed in v11 I now have polling on all the time
                            'AVTGetPositionInfo(0)  ' removed in v11 I now have polling on all the time
                            AVTGetTransportInfo(0)
                        End If
                    ElseIf ObjectserviceID = "urn:upnp-org:serviceId:ConnectionManager" Or ObjectserviceID = "urn:schemas-upnp-org:service:ConnectionManager" Then
                        Try
                            ConnectionManager = objService
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found ConnectionManager for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                            If Not ConnectionManager Is Nothing Then
                                Try
                                    'CMGetCurrentConnectionIDs()
                                    'CMGetCurrentConnectionInfo()
                                    CMGetProtocolInfo()
                                Catch ex As Exception
                                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " getting Connections info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                Try
                                    ConnectionManager.AddCallback(myConnectionManagerCallback)
                                    Log("ConnectionManagerCallBack added for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                                Catch ex As Exception
                                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " adding ConnectionManager Call Back with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            End If
                        Catch ex As Exception

                        End Try

                    ElseIf ObjectserviceID = "urn:upnp-org:serviceId:ContentDirectory" Then 'Or ObjectserviceID = "urn:schemas-upnp-org:service:RenderingControl" Then removed in HS3, not sure why this is here?
                        Try
                            ContentDirectory = objService
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found ContentDirectory for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                            If Not ContentDirectory Is Nothing Then
                                Try
                                    ContentDirectory.AddCallback(myContentDirectoryCallback)
                                    Log("ContentDirectoryCallBack added for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                                    GetSearchCapabilities()
                                    GetSortCapabilities()
                                Catch ex As Exception
                                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " adding ContentDirectory Call Back with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            End If

                        Catch ex As Exception

                        End Try
                    ElseIf ObjectserviceID = "urn:upnp-org:serviceId:RenderingControl" Or ObjectserviceID = "urn:schemas-upnp-org:service:RenderingControl" Then ' added in HS3 
                        RenderingControl = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found RenderingControl for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        Try
                            Dim AllowedValueString As String = SearchServiceDocument(MyDocumentURL, ObjectserviceID, "SetVolume", "Channel")
                            If AllowedValueString <> "" Then
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found SetVolume/Channel allowed values = " & AllowedValueString, LogType.LOG_TYPE_INFO)
                                Dim MyPossibleBalanceSettings As String()
                                MyPossibleBalanceSettings = GetXMLValues(AllowedValueString, "allowedValue")
                                If Not MyPossibleBalanceSettings Is Nothing Then
                                    For index = 0 To UBound(MyPossibleBalanceSettings, 1)
                                        If MyPossibleBalanceSettings(index).ToUpper = "LF" Then
                                            BalanceIsConfigurable = True
                                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " set BalanceIsConfigurable", LogType.LOG_TYPE_INFO)
                                            Exit For
                                        ElseIf MyPossibleBalanceSettings(index).ToUpper = "RF" Then
                                            BalanceIsConfigurable = True
                                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " set BalanceIsConfigurable", LogType.LOG_TYPE_INFO)
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If

                            AllowedValueString = ""
                            AllowedValueString = SearchServiceDocument(MyDocumentURL, ObjectserviceID, "SetVolume", "DesiredVolume")

                            If AllowedValueString <> "" Then
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found SetVolume/DesiredVolume allowed values = " & AllowedValueString, LogType.LOG_TYPE_INFO)
                                Dim MyPossibleMinVolumeSettings As String()
                                Dim MyPossibleMaxVolumeSettings As String()
                                MyPossibleMinVolumeSettings = GetXMLValues(AllowedValueString, "minimum")
                                MyPossibleMaxVolumeSettings = GetXMLValues(AllowedValueString, "maximum")
                                If MyPossibleMinVolumeSettings IsNot Nothing Then
                                    If MyPossibleMinVolumeSettings(0) <> "" Then
                                        MyMinimumVolume = Val(MyPossibleMinVolumeSettings(0))
                                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " set MyMinimumVolume = " & MyMinimumVolume.ToString, LogType.LOG_TYPE_INFO)
                                    End If
                                End If
                                If MyPossibleMaxVolumeSettings IsNot Nothing Then
                                    If MyPossibleMaxVolumeSettings(0) <> "" Then
                                        MyMaximumVolume = Val(MyPossibleMaxVolumeSettings(0))
                                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " set MyMaximumVolume = " & MyMaximumVolume.ToString, LogType.LOG_TYPE_INFO)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Log("ExtractAllServices found RenderingControl for device = " & MyUPnPDeviceName & " but ran into an error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try

                    ElseIf ObjectserviceID = "urn:ce-org:serviceId:RemoteUIServer1" Then
                        RemoteUIServer = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found RemoteUIServerCallback for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

                    ElseIf ObjectserviceID = "urn:samsung.com:serviceId:TestRCRService" Then
                        RemoteControlService = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found TestRCRService for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

                    ElseIf ObjectserviceID = "urn:samsung.com:serviceId:MessageBoxService" Then
                        MessageBoxService = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found MessageBoxService for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        HSRefMessage = GetIntegerIniFile(MyUDN, "di" & HSDevices.Message.ToString & "HSCode", -1)
                        If HSRefMessage = -1 Then
                            HSRefMessage = CreateHSServiceDevice(HSRefMessage, HSDevices.Message.ToString)
                            CreateMessageButtons(HSRefMessage)
                        End If
                        If HSRefMessage <> -1 Then
                            Dim MessageServiceSetting As String = ""
                            MessageServiceSetting = GetStringIniFile("Message Service by UDN", MyUDN, "")
                            If MessageServiceSetting = "" Then
                                ' never stored before, store now
                                WriteBooleanIniFile("Message Service by UDN", MyUDN, False)
                                hs.SetDeviceValueByRef(HSRefMessage, dsDeactivated, True)
                                MyMessageServiceActive = False
                            Else
                                Try
                                    If Not GetBooleanIniFile("Message Service by UDN", MyUDN, False) Then
                                        hs.SetDeviceValueByRef(HSRefMessage, dsDeactivated, True)
                                        MyMessageServiceActive = False
                                    Else
                                        hs.SetDeviceValueByRef(HSRefMessage, dsActivateOnLine, True)
                                        MyMessageServiceActive = True
                                    End If
                                Catch ex As Exception
                                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " setting Message Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            End If
                        End If
                    ElseIf ObjectserviceID = "urn:schemas-sony-com:serviceId:IRCC" Then
                        RemoteControlService = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found IRCC for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "SonyIRRC")
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diSonyRemoteRegisterType.ToString, MySonyRegisterMode)
                        SonySetupRemoteInfo()
                        SetHSRemoteState()
                    ElseIf ObjectserviceID = "urn:schemas-sony-com:serviceId:Party" Then
                        SonyPartyService = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found Party for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        SonyUpdatePartyButtons()
                        If Not SonyPartyService Is Nothing Then
                            Try
                                SonyPartyService.AddCallback(mySonyPartyCallback)
                                Log("SonyPartyService added for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                            Catch ex As Exception
                                Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " adding SonyParyService Call Back with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                        End If
                    ElseIf objectserviceType = "urn:dial-multiscreen-org:service:dial:1" Or objectserviceType = "urn:dial-multiscreen-org:device:dialreceiver:1" Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found DIAL Service for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        If MyUPnPDeviceManufacturer.ToUpper.IndexOf("ROKU") <> -1 Then
                            If RokuRetrieveDIALAppList(MyUPnPDevice.Location) Then '"http://" & MyIPAddress & ":" & MyIPPort) Then
                                If HSRefRemote = -1 Then CreateHSRokuRemoteButtons(False)
                            End If
                        Else
                            If RetrieveDIALAppList(MyUPnPDevice.ApplicationURL) Then
                                If HSRefRemote = -1 Then CreateHSDIALRemoteButtons(False)
                            End If
                        End If
                    ElseIf objectserviceType = "urn:lge-com:service:webos-second-screen:1" Then
                        RemoteControlService = objService
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found LG WebOS Second Screen for device = " & MyUPnPDeviceName & " with DeviceServiceType= " & MyUPnPDeviceServiceType, LogType.LOG_TYPE_INFO)
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "LG")
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "3000")
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "")
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "ZTeaydfwEuaM5bkYY377PA==")
                        If HSRefRemote = -1 Then CreateHSLGRemoteButtons(False)
                        CreateHSLGRemoteServices()
                        If Not SendLGRegistration() Then
                            MyRemoteServiceActive = False
                            SetHSRemoteState()
                        Else
                            If HSRefRemote <> -1 Then
                                Dim RemoteServiceSetting As String = ""
                                RemoteServiceSetting = GetStringIniFile("Remote Service by UDN", MyUDN, "")
                                If RemoteServiceSetting = "" Then
                                    ' never stored before, store now
                                    WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                                    MyRemoteServiceActive = True
                                End If
                            End If
                            SetAdministrativeStateRemote(True)
                        End If
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found additional service with ID = " & ObjectserviceID, LogType.LOG_TYPE_INFO)
                    End If
                Next
            Catch ex As Exception
                Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " going through the services with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                If HSRefDevice = -1 Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ExtractAllServices for device = " & MyUPnPDeviceName & ". This needs fixing !! ", LogType.LOG_TYPE_WARNING)
                    ' update the IP adress if this was caused by a reconnect
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diIPAddress.ToString, MyIPAddress)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diIPPort.ToString, MyIPPort)
                End If
            Catch ex As Exception
                Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " creating HS Device Codes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If Not RenderingControl Is Nothing Then
                Try
                    RenderingControl.AddCallback(myRenderingControlCallback)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RenderingControlCallback added for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " adding RenderingControl Call Back  with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Try
                    'If BrightnessIsConfigurable Then RCGetBrightness()
                    'If ColorTemperatureIsConfigurable Then RCGetColorTemperature()
                    'If ContrastIsConfigurable Then RCGetContrast()
                    If LoudnessIsConfigurable Then RCGetLoudness()
                    If MuteIsConfigurable Then RCGetMute()
                    'If SharpnessIsConfigurable Then RCGetSharpness()
                    If VolumeIsConfigurable Then RCGetVolume()
                    'If VolumeDBIsConfigurable Then RCGetVolumeDB()
                    'RCGetVolumeDBRange()
                    'RCListPresets()
                    'If GetImageRotationIsConfigurable Then RCX_GetImageRotation()
                    'If GetImageScaleIsConfigurable Then RCX_GetImageScale()
                    'If GetSlideShowEffectIsConfigurable Then RCX_GetSlideShowEffect()
                Catch ex As Exception
                    Log("Error in ExtractAllServices1 for device = " & MyUPnPDeviceName & " adding RenderingControl Call Back  with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            If AVTransport IsNot Nothing Then
                MyTrackInfoHasChanged = True
                MyTransportStateHasChanged = True
                UpdateTransportState()
            End If
            Try
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices for device = " & MyUPnPDeviceName & " found DeviceManufacturer = " & MyUPnPDeviceManufacturer, LogType.LOG_TYPE_INFO)
                If MyUPnPDeviceServiceType = "RCR" And MyUPnPDeviceManufacturer.ToUpper.IndexOf("SAMSUNG") <> -1 Then
                    If ExtractSamsungInfoFromDeviceXML(MyUPnPDevice.DeviceUPnPDocument) Then
                        ' this means we are websocket based as opposed to legacy. Now it can be with or without PIN
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractAllServices found Samsung WebSocket based remote Service for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                        'If Not GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                        'SamsungCloseTCPConnection(False)
                        'Else
                        'If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
                        'SamsungOpenUnEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, ""), GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, ""))
                        'ElseIf GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocketPIN" Then
                        'If SamsungGetIdentifyParms() Then
                        'If SamsungAuthenticateUsePIN(GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, ""), True) <> "" Then
                        'SamsungOpenEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungSessionID.ToString, ""))
                        'End If
                        'End If
                        'End If
                        'End If
                    Else
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "Samsungiapp")
                        Try
                            If Not GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                                SamsungCloseTCPConnection(False)
                            Else
                                SamsungEstablishTCPConnection()
                            End If
                        Catch ex As Exception
                            Log("Error in ExtractAllServices for Samsung device = " & MyUPnPDeviceName & " setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If

                    If HSRefRemote = -1 Then
                        'WriteSamsungKeyInfoToInfoFile()
                        'CreateSamsungRemoteIniFileInfo()
                        CreateHSSamsungRemoteButtons(False)
                    End If

                    If HSRefRemote <> -1 Then
                        Dim RemoteServiceSetting As String = ""
                        RemoteServiceSetting = GetStringIniFile("Remote Service by UDN", MyUDN, "")
                        If RemoteServiceSetting = "" Then
                            ' never stored before, store now
                            WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                            MyRemoteServiceActive = True
                        End If
                        SetAdministrativeStateRemote(True)
                    End If
                    If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
                        CreateHSSamsungRemoteServices()
                    End If
                    If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
                        SamsungOpenUnEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, ""), GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, ""))
                    ElseIf GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocketPIN" Then
                        If SamsungGetIdentifyParms() Then
                            If SamsungAuthenticateUsePIN(GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, ""), True) <> "" Then
                                SamsungOpenEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungSessionID.ToString, ""))
                            End If
                        End If
                    End If
                End If
                If MyUPnPDeviceServiceType = "DMR" And MyUPnPDeviceManufacturer.ToUpper.IndexOf("ONKYO") <> -1 Then
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "Onkyo")
                    If GetStringIniFile(MyUDN, DeviceInfoIndex.diOnkyoPortNbr.ToString, "") <> "" Then              ' added in v24 to make it configurable
                        MyOnkyoPortNbr = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diOnkyoPortNbr.ToString, 60128)
                    Else
                        WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diOnkyoPortNbr.ToString, MyOnkyoPortNbr)
                    End If
                    If HSRefRemote = -1 Then
                        WriteOnkyoKeyInfoToInfoFile()
                        CreateHSOnkyoRemoteButtons(False)
                    End If
                    If HSRefRemote <> -1 Then
                        Dim RemoteServiceSetting As String = ""
                        RemoteServiceSetting = GetStringIniFile("Remote Service by UDN", MyUDN, "")
                        If RemoteServiceSetting = "" Then
                            ' never stored before, store now
                            WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                            MyRemoteServiceActive = True
                        End If
                    End If
                    Try
                        If Not GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                            OnkyoCloseTCPConnection()
                        Else
                            OnkyoEstablishTCPConnection()
                            OnkyoGetBasicInfo()
                        End If
                    Catch ex As Exception
                        Log("Error in ExtractAllServices for Onkyo device = " & MyUPnPDeviceName & " setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception

            End Try

        Catch ex As Exception
            Log("Error in ExtractAllServices for device = " & MyUPnPDeviceName & " with error=" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function CheckIPAddressChange() As Boolean
        CheckIPAddressChange = False
        If MyUPnPDevice Is Nothing Then Exit Function
        'Dim IPAddressInfo As IPAddressInfo
        'IPAddressInfo.IPAddress = ""
        'IPAddressInfo.IPPort = ""
        'Try
        'Dim pDescDoc As IUPnPDeviceDocumentAccess
        'Dim UPnPDocumentURL As String = ""
        ''pDescDoc = MyUPnPDevice
        'UPnPDocumentURL = MyUPnPDevice.Location & MyUPnPDevice.IPAddress 'pDescDoc.GetDocumentURL()
        'MyDocumentURL = UPnPDocumentURL
        'IPAddressInfo = ExtractIPInfo(UPnPDocumentURL)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckIPAddressChange for UPnPDevice = " & MyUPnPDeviceName & " found IPAddress = " & MyUPnPDevice.IPAddress, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckIPAddressChange for UPnPDevice = " & MyUPnPDeviceName & " found IPPort    = " & MyUPnPDevice.IPPort, LogType.LOG_TYPE_INFO)
        'UPnPDocumentURL = Nothing
        'Catch ex As Exception
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Error in CheckIPAddressChange for UPnPDevice = " & MyUPnPDeviceName & " the device is off-line or the plugin did not find the documentURL with error = " &ex.Message, LogType.LOG_TYPE_ERROR)
        'Exit Function
        'End Try
        If MyUPnPDevice.IPAddress <> MyIPAddress Then
            CheckIPAddressChange = True
            Log("IPAddress for UPnPDevice = " & MyUPnPDeviceName & " has changed. Old = " & MyIPAddress & ". New = " & MyUPnPDevice.IPAddress, LogType.LOG_TYPE_INFO)
            MyIPAddress = MyUPnPDevice.IPAddress
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diIPAddress.ToString, MyIPAddress)
        End If
    End Function

    Private Function Reachable() As Boolean

        Reachable = False
        If MyUPnPDevice Is Nothing Then Exit Function ' no device is alive on the network

        If MyReferenceToMyController Is Nothing Then
            Exit Function
        End If

        Try
            If MyUPnPDevice.Alive Then
                MyFailedPingCount = 0
                If DeviceStatus = "Offline" Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is reachable on network. Attempting to reconnect", LogType.LOG_TYPE_INFO)
                    If MyUPnPDeviceServiceType = "RCR" Then
                        ' just put on line, no need to do anything else and update state
                        DeviceStatus = "Online"
                        SetHSMainState()
                    ElseIf MyUPnPDeviceServiceType = "DIAL" Then
                        If MyUPnPDeviceManufacturer.ToUpper.IndexOf("ROKU") <> -1 Then
                            If RokuRetrieveDIALAppList(MyUPnPDevice.Location) Then '"http://" & MyIPAddress & ":" & MyIPPort) Then
                                If HSRefRemote = -1 Then CreateHSRokuRemoteButtons(False)
                            End If
                            DeviceStatus = "Online"
                        Else
                            If RetrieveDIALAppList(MyUPnPDevice.ApplicationURL) Then '"http://" & MyIPAddress & ":" & MyIPPort) Then
                                If HSRefRemote = -1 Then CreateHSDIALRemoteButtons(False)
                            End If
                            DeviceStatus = "Online"
                        End If
                    Else
                        Connect("uuid:" & MyUDN)
                    End If
                Else
                    'If piDebuglevel > DebugLevel.dlEvents Then log( "Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is reachable on network")
                End If
                Reachable = True
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is not reachable and DeviceStatus = " & DeviceStatus, LogType.LOG_TYPE_INFO)
                If DeviceStatus = "Online" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is not reachable but still on-line", LogType.LOG_TYPE_INFO)
                    If CheckIPAddressChange() Then
                        Reachable = True
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is not reachable but still on-line and we shouldn't be here", LogType.LOG_TYPE_INFO)
                        Exit Function
                    End If
                    Log("Reachable called for UPnPDevice " & MyUPnPDeviceName & " which is not reachable anymore and set Off-line", LogType.LOG_TYPE_WARNING)
                    If MyUPnPDeviceServiceType = "RCR" Or MyUPnPDeviceServiceType = "DIAL" Then
                        DeviceStatus = "Offline"
                        SetHSMainState()
                    Else
                        Disconnect(False)
                    End If
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Reachable for UPnPDevice " & MyUPnPDeviceName & " calling the ping status with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function GetUDNTag(ByVal InString As String) As String
        GetUDNTag = InString ' if there is no tag, return unmodified
        Dim Index As Integer = InString.IndexOf("<uuid:")
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetUDNTag called with String = " & InString & " and UDN = " & GetUDNTag & " and Index " & Index.ToString)
        If Index <= 0 Then Exit Function
        Try
            GetUDNTag = InString.Substring(Index, InString.Length - Index)
            GetUDNTag = GetUDNTag.Trim("<", ">")
        Catch ex As Exception
        End Try
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetUDNTag called with String = " & InString & " and UDN = " & GetUDNTag)
    End Function

    Public Sub DeviceTrigger(ByVal triggerEvent As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceTrigger called for device - " & MyUPnPDeviceName & " with Trigger = " & triggerEvent.ToString, LogType.LOG_TYPE_INFO)
        Dim TrigsToCheck() As IPlugInAPI.strTrigActInfo
        'Dim TC As IPlugInAPI.strTrigActInfo

        Dim strTrig As strTrigger = Nothing
        Try
            If MainInstance <> "" Then
                TrigsToCheck = callback.GetTriggersInst(sIFACE_NAME, MainInstance)
            Else
                TrigsToCheck = callback.GetTriggers(sIFACE_NAME)
            End If
        Catch ex As Exception
            TrigsToCheck = Nothing
        End Try
        If TrigsToCheck IsNot Nothing AndAlso TrigsToCheck.Count > 0 Then
            For Each TC As IPlugInAPI.strTrigActInfo In TrigsToCheck
                If PIDebuglevel > DebugLevel.dlEvents Then Log("DeviceTrigger found Trigger: EvRef=" & TC.evRef.ToString & ", Trig/SubTrig=" & TC.TANumber.ToString & "/" & TC.SubTANumber.ToString & ", UID=" & TC.UID.ToString, LogType.LOG_TYPE_INFO)
                'Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = Nothing
                'TrigsToCheck = callback.TriggerMatches(sIFACE_NAME, Info.TANumber, Info.SubTANumber)
                'callback.TriggerFire(IFACE_NAME, Info)
                If Not (TC.DataIn Is Nothing) Then
                    Dim trigger As New trigger
                    DeSerializeObject(TC.DataIn, trigger)
                    Dim Command As String = ""
                    Dim PlayerUDN As String = ""
                    For Each sKey In trigger.Keys
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
                        Select Case True
                            Case InStr(sKey, "PlayerListTrigger") > 0 AndAlso trigger(sKey) <> ""
                                If trigger(sKey) <> MyUDN Then
                                    Exit For ' not for this player
                                End If
                                PlayerUDN = trigger(sKey)
                                If Command <> "" Then
                                    callback.TriggerFire(sIFACE_NAME, TC)
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceTrigger called TriggerFire for device - " & MyUPnPDeviceName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                    Exit For
                                End If
                            Case InStr(sKey, "CommandListTrigger") > 0 AndAlso trigger(sKey) <> ""
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceTrigger for Zone - " & ZoneName & " command = " & trigger(sKey) & " while looking for triggerEvent = " & triggerEvent, LogType.LOG_TYPE_INFO)
                                If trigger(sKey) <> triggerEvent Then
                                    Exit For ' not right state
                                End If
                                Command = trigger(sKey)
                                If PlayerUDN <> "" Then
                                    callback.TriggerFire(sIFACE_NAME, TC)
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeviceTrigger called TriggerFire for device - " & MyUPnPDeviceName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                    Exit For
                                End If
                        End Select
                    Next
                End If
            Next
        End If
    End Sub


    Public Sub PlayChangeNotifyCallback(ByVal ChangeType As player_status_change, ByVal ChangeValue As player_state_values, Optional SendDeviceTrigger As Boolean = True)
        ' Raised by a Music plug-in whenever various music plug-in status changes.  
        'HSTouch and other applications can add an event handler for this event in your plug-in to be notified of changes in the status.
        ' Public Enum player_status_change
        '        SongChanged = 1           raised whenever the current song changes
        '        PlayStatusChanged = 2     raised when pause, stop, play, etc. pressed.
        '        PlayList = 3              raised whenever the current playlist changes
        '        Library = 4               raised when the library changes
        '        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        '    End Enum
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayChangeNotifyCallback called for device - " & MyUPnPDeviceName & " with ChangeType = " & ChangeType.ToString & " and Changevalue = " & ChangeValue.ToString, LogType.LOG_TYPE_INFO)
        'log( "PlayChangeNotifyCallback called for device - " & MyUPnPDeviceName & " with ChangeType = " & ChangeType.ToString & " and Changevalue = " & ChangeValue.ToString)
        If gInterfaceStatus <> ERR_NONE Then
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Warning PlayChangeNotifyCallback called for device - " & MyUPnPDeviceName & " before plugin is initialized. Nothing sent")
            Exit Sub ' no updates to be sent until completely intialized. Else the multizone API is hosed.
        End If
        Dim TriggerEvent As String = ""
        ' trigger Names
        '   Track Change
        '   Player Stop
        '   Player Paused
        '   "Player Start Playing
        If ChangeType = player_status_change.SongChanged Then
            TriggerEvent = "Player Track Change"
        ElseIf ChangeType = player_status_change.PlayStatusChanged Then
            If MyCurrentPlayerState = player_state_values.Playing Then
                TriggerEvent = "Player Start Playing"
            ElseIf MyCurrentPlayerState = player_state_values.Paused Then
                TriggerEvent = "Player Paused"
            ElseIf MyCurrentPlayerState = player_state_values.Stopped Then
                TriggerEvent = "Player Stop"
            ElseIf MyCurrentPlayerState = player_state_values.Forwarding Then
                TriggerEvent = "Player Forwarding"
            ElseIf MyCurrentPlayerState = player_state_values.Rewinding Then
                TriggerEvent = "Player Rewinding"
            End If
        ElseIf ChangeType = player_status_change.DeviceStatusChanged Then
            If ChangeValue = player_state_values.Online Then
                TriggerEvent = "Player Online"
            ElseIf ChangeValue = player_state_values.Offline Then
                TriggerEvent = "Player Offline"
            End If
        ElseIf ChangeType = player_status_change.AlarmStart Then
            TriggerEvent = "Player Alarm Start"
        ElseIf ChangeType = player_status_change.ConfigChange Then
            TriggerEvent = "Player Config Change"
        ElseIf ChangeType = player_status_change.NextSong Then
            TriggerEvent = "Player Next Track Change"
        End If

        If TriggerEvent <> "" And SendDeviceTrigger Then DeviceTrigger(TriggerEvent)

    End Sub

    Private Function GetAlbumArtPath(ByVal AlbumURI As String, ByVal NextTrack As Boolean) As String
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetAlbumArtPath called for device " & MyUPnPDeviceName & " with AlbumURI = " & AlbumURI & " and NextTrack = " & NextTrack.ToString, LogType.LOG_TYPE_INFO)
        Dim AlbumArtImage As Image = Nothing
        GetAlbumArtPath = NoArtPath
        If AlbumURI = NoArtPath Or AlbumURI = "" Then
            GetAlbumArtPath = NoArtPath
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        ' prevent multiple saves and avoid this pesky GDI+ error
        If Not (AlbumURI.ToLower().StartsWith("http://") Or AlbumURI.ToLower().StartsWith("https://") Or AlbumURI.ToLower().StartsWith("file:")) Then AlbumURI = "http://" & MyUPnPDevice.IPAddress & ":" & MyUPnPDevice.IPPort & AlbumURI
        If NextTrack Then
            If MyPreviousNextAlbumURI <> "" And MyPreviousNextAlbumURI = AlbumURI And MyPreviousAlbumArtPath <> "" Then
                GetAlbumArtPath = MyPreviousNextAlbumArtPath
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetAlbumArtPath returned for device - " & MyUPnPDeviceName & " with AlbumURI = " & AlbumURI & " and cached returned path= " & GetAlbumArtPath, LogType.LOG_TYPE_INFO)
                Exit Function
            End If
        Else
            If MyPreviousAlbumURI <> "" And MyPreviousAlbumURI = AlbumURI And MyPreviousAlbumArtPath <> "" Then
                GetAlbumArtPath = MyPreviousAlbumArtPath
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetAlbumArtPath returned for device - " & MyUPnPDeviceName & " with AlbumURI = " & AlbumURI & " and cached returned path= " & GetAlbumArtPath, LogType.LOG_TYPE_INFO)
                Exit Function
            End If
        End If
        AlbumArtImage = GetPicture(AlbumURI)
        GetAlbumArtPath = SaveArtwork(AlbumArtImage, AlbumURI, NextTrack)
    End Function

    Private Function GetPicture(ByVal url As String) As Image
        ' Get the picture at a given URL.
        Dim web_client As New WebClient()
        GetPicture = Nothing
        Try
            url = Trim(url)
            If url = "" Then
                Return Nothing
                Exit Function
            End If
            ' If Not (url.ToLower().StartsWith("http://") Or url.ToLower().StartsWith("file:")) Then url = "http://" & url
            Dim image_stream As New MemoryStream(web_client.DownloadData(url))
            GetPicture = Image.FromStream(image_stream, True, True)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPicture called for device - " & MyUPnPDeviceName & " url= " & url.ToString & " caused error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            GetPicture = Nothing
        Finally
            web_client.Dispose()
        End Try
    End Function


    Private Function SaveArtwork(AlbumArtImage As Image, ByVal AlbumURI As String, ByVal NextTrack As Boolean) As String
        If AlbumArtImage Is Nothing Then
            SaveArtwork = NoArtPath
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        If AlbumArtImage.Height = 0 Or AlbumArtImage.Width = 0 Then
            SaveArtwork = NoArtPath
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SaveArtWork encountered zero width/height picture for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
            AlbumArtImage.Dispose()
            AlbumArtImage = Nothing
            GC.Collect()
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        Dim ExtensionIndex As Integer = 0
        Dim ExtensionType As String = ".jpg"
        Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg
        Try
            ' get the extension file type
            ExtensionIndex = AlbumURI.LastIndexOf(".")
            Dim TempExtensiontype As String = ""
            If ExtensionIndex <> -1 Then
                TempExtensiontype = AlbumURI.Substring(ExtensionIndex, AlbumURI.Length - ExtensionIndex)
                If UCase(TempExtensiontype) = ".PNG" Then
                    ExtensionType = ".png"
                    'ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SaveArtWork for device = " & MyUPnPDeviceName & " has set image to .PNG")
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveArtWork for device = " & MyUPnPDeviceName & " when searching for the file type with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        'Dim FilePath As String = ""
        Dim TempFilePath As String = ""
        Dim MyNextArtFileIndex As Integer = 0
        Dim MyArtFileIndex As Integer = 0
        If NextTrack Then
            MyNextArtFileIndex = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtFileIndex.ToString, 0)
            MyNextArtFileIndex += 1
            WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diNextArtFileIndex.ToString, MyNextArtFileIndex)
            'FilePath = CurrentAppPath & "/html" & ArtWorkPath & "NextCover_" & MyUDN & "_" & MyNextArtFileIndex.ToString & ExtensionType
            TempFilePath = "NextCover_" & MyUDN & "*.*"
            MyNextArtFileIndex = MyNextArtFileIndex + 1
        Else
            MyArtFileIndex = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diArtFileIndex.ToString, 0)
            MyArtFileIndex += 1
            WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diArtFileIndex.ToString, MyArtFileIndex)
            'FilePath = CurrentAppPath & "/html" & ArtWorkPath & "Cover_" & MyUDN & "_" & MyArtFileIndex.ToString & ExtensionType
            TempFilePath = "Cover_" & MyUDN & "*.*"
            MyArtFileIndex = MyArtFileIndex + 1
        End If
        If ImRunningLocal Then
            Try
                ' let's try to delete the previous file
                Dim TempPath As String = ArtWorkPath.Remove(ArtWorkPath.Length - 1, 1) ' remove the "/' character
                For Each FileFound As String In Directory.GetFiles(CurrentAppPath & "/html" & TempPath, TempFilePath)
                    File.Delete(FileFound) ' tralala tobe fixed dcor
                Next
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in SaveArtWork when deleting the previous art work for device = " & MyUPnPDeviceName & " with Filename = " & TempFilePath & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else

        End If

        Try
            If NextTrack Then
                Dim SuccesfullSave As Boolean = False
                SuccesfullSave = hs.WriteHTMLImage(AlbumArtImage, FileArtWorkPath & "NextCover_" & MyUDN & "_" & MyNextArtFileIndex.ToString & ExtensionType, True)
                If Not SuccesfullSave Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveArtWork for Device = " & MyUPnPDeviceName & " had error storing NextCover at " & FileArtWorkPath & "NextCover_" & MyUDN & "_" & MyNextArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_ERROR)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveArtWork for Device = " & MyUPnPDeviceName & " stored NextCover at " & FileArtWorkPath & "NextCover_" & MyUDN & "_" & MyNextArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_INFO)
                End If
                SaveArtwork = ArtWorkPath & "NextCover_" & MyUDN & "_" & MyNextArtFileIndex.ToString & ExtensionType
                MyPreviousNextAlbumArtPath = SaveArtwork
                MyPreviousNextAlbumURI = AlbumURI
            Else
                Dim SuccesfullSave As Boolean = False
                SuccesfullSave = hs.WriteHTMLImage(AlbumArtImage, FileArtWorkPath & "Cover_" & MyUDN & "_" & MyArtFileIndex.ToString & ExtensionType, True)
                If Not SuccesfullSave Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveArtWork for Device = " & MyUPnPDeviceName & " had error storing Cover at " & FileArtWorkPath & "Cover_" & MyUDN & "_" & MyArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_ERROR)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveArtWork for Device = " & MyUPnPDeviceName & " stored Cover at " & FileArtWorkPath & "Cover_" & MyUDN & "_" & MyArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_INFO)
                End If
                SaveArtwork = ArtWorkPath & "Cover_" & MyUDN & "_" & MyArtFileIndex.ToString & ExtensionType
                MyPreviousAlbumArtPath = SaveArtwork
                MyPreviousAlbumURI = AlbumURI
            End If
        Catch ex As Exception
            If NextTrack Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveArtWork storing artwork for device - " & MyUPnPDeviceName & " and path = " & CurrentAppPath & "/html" & ArtWorkPath & "NextCover" & MyHSTMusicIndex.ToString & "_" & MyNextArtFileIndex.ToString & ExtensionType & " with error= " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveArtWork storing artwork for device - " & MyUPnPDeviceName & " and path = " & CurrentAppPath & "/html" & ArtWorkPath & "Cover" & MyHSTMusicIndex.ToString & "_" & MyArtFileIndex.ToString & ExtensionType & " with error= " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            SaveArtwork = NoArtPath
        End Try
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SaveArtWork returned for device - " & MyUPnPDeviceName & " with AlbumURI = " & AlbumURI & " and returned path= " & SaveArtwork, LogType.LOG_TYPE_INFO)
        Try
            AlbumArtImage.Dispose()
            AlbumArtImage = Nothing
            ImageFormat = Nothing
            GC.Collect()
        Catch ex As Exception
        End Try

    End Function


    Public Enum player_status_change
        SongChanged = 1         'raises whenever the current song changes
        PlayStatusChanged = 2   'raises when pause, stop, play, etc. pressed.
        PlayList = 3            'raises whenever the current playlist changes
        Library = 4             'raises when the library changes
        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        AlarmStart = 12          ' raised when the alarm goes off
        ConfigChange = 13        ' raised when the configuration of a device changes like alarm info being modified
        NextSong = 14            ' raised when the next song is about to start
        PartyOn = 15
        PartyOff = 16
    End Enum

    Public Enum player_state_values
        Playing = 1
        Stopped = 2
        Paused = 3
        Forwarding = 4
        Rewinding = 5
        Transitioning = 6
        UpdateHSServerOnly = 17
        Online = 18
        Offline = 19
    End Enum

    Public Enum repeat_modes
        repeat_off = 0
        repeat_one = 1
        repeat_all = 2
    End Enum

    Public Enum Shuffle_modes
        Shuffled = 1
        Ordered = 2
        Sorted = 3
    End Enum



    Public Function ToBase64(ByVal data As String) As String
        ToBase64 = ""
        If data Is Nothing Then Exit Function
        Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(data))
    End Function

    Public Function FromBase64(ByVal base64 As String) As Byte()
        FromBase64 = {}
        If base64 = "" Then Exit Function
        ' Dim DecodedString As String = System.Text.Encoding.UTF8.GetString(FromBase64(EncodedString)
        Return Convert.FromBase64String(base64)
    End Function

    Private Function UPnP_Error(ByVal ErrNumber) As String

        'http://msdn.microsoft.com/en-us/library/ms899567.aspx  has the errors posted here

        'If (ErrNumber < &H8004042B) And (ErrNumber >= &H80040300) Then ' UPNP_E_ACTION_SPECIFIC_BASE <= UPnP_Error <= UPNP_E_ACTION_SPECIFIC_MAX 
        Dim SaveError
        SaveError = ErrNumber
        If (ErrNumber < &H8004042B) And (ErrNumber >= &H80040200) Then
            ErrNumber = ErrNumber - &H80040300 + &H258
        ElseIf ErrNumber < 300 Then
            ErrNumber = ErrNumber + 700
        End If
        'UPNP_E_ROOT_ELEMENT_EXPECTED = 0x80040200, = 344 dec
        'UPNP_E_DEVICE_ELEMENT_EXPECTED = 0x80040201,
        'UPNP_E_SERVICE_ELEMENT_EXPECTED = 0x80040202,
        'UPNP_E_SERVICE_NODE_INCOMPLETE = 0x80040203,
        'UPNP_E_DEVICE_NODE_INCOMPLETE = 0x80040204,
        'UPNP_E_ICON_ELEMENT_EXPECTED = 0x80040205,
        'UPNP_E_ICON_NODE_INCOMPLETE = 0x80040206,
        'UPNP_E_INVALID_ACTION = 0x80040207,
        'UPNP_E_INVALID_ARGUMENTS = 0x80040208,
        'UPNP_E_OUT_OF_SYNC = 0x80040209,
        'UPNP_E_ACTION_REQUEST_FAILED = 0x80040210,
        'UPNP_E_TRANSPORT_ERROR = 0x80040211,
        'UPNP_E_VARIABLE_VALUE_UNKNOWN = 0x80040212,
        'UPNP_E_INVALID_VARIABLE = 0x80040213,
        'UPNP_E_DEVICE_ERROR = 0x80040214,
        'UPNP_E_PROTOCOL_ERROR = 0x80040215,
        'UPNP_E_ERROR_PROCESSING_RESPONSE = 0x80040216,
        'UPNP_E_DEVICE_TIMEOUT = 0x80040217,

        Select Case ErrNumber
            Case 0
                UPnP_Error = "Successfull Action"
            Case 344
                UPnP_Error = "UPNP_E_ROOT_ELEMENT_EXPECTED"
            Case 345
                UPnP_Error = "UPNP_E_DEVICE_ELEMENT_EXPECTED "
            Case 346
                UPnP_Error = "UPNP_E_SERVICE_ELEMENT_EXPECTED"
            Case 347
                UPnP_Error = "UPNP_E_SERVICE_NODE_INCOMPLETE"
            Case 348
                UPnP_Error = "UPNP_E_DEVICE_NODE_INCOMPLETE"
            Case 349
                UPnP_Error = "UPNP_E_ICON_ELEMENT_EXPECTED"
            Case 350
                UPnP_Error = "UPNP_E_ICON_NODE_INCOMPLETE"
            Case 351
                UPnP_Error = "UPNP_E_INVALID_ACTION"
            Case 352
                UPnP_Error = "UPNP_E_INVALID_ARGUMENTS "
            Case 353
                UPnP_Error = "UPNP_E_OUT_OF_SYNC"
            Case 360
                UPnP_Error = "UPNP_E_ACTION_REQUEST_FAILED "
            Case 361
                UPnP_Error = "UPNP_E_TRANSPORT_ERROR "
            Case 362
                UPnP_Error = "UPNP_E_VARIABLE_VALUE_UNKNOWN "
            Case 363
                UPnP_Error = "UPNP_E_INVALID_VARIABLE"
            Case 364
                UPnP_Error = "UPNP_E_DEVICE_ERROR "
            Case 365
                UPnP_Error = "UPNP_E_PROTOCOL_ERROR "
            Case 366
                UPnP_Error = "UPNP_E_ERROR_PROCESSING_RESPONSE "
            Case 367
                UPnP_Error = "UPNP_E_DEVICE_TIMEOUT"
            Case 351
                UPnP_Error = "UPNP_E_INVALID_DOCUMENT"
            Case 352
                UPnP_Error = "UPNP_E_EVENT_SUBSCRIPTION_FAILED"
            Case 401
                UPnP_Error = "Invalid Action"
            Case 402
                UPnP_Error = "Invalid args"
            Case 403
                UPnP_Error = "Out of Sync"
            Case 404
                UPnP_Error = "Invalid Var"
            Case 501
                UPnP_Error = "Action failed"
            Case 701
                UPnP_Error = "No such object / Transition no available / Incompatible protocol info"
            Case 702
                UPnP_Error = "Invalid CurrentTagValue / Invalid InstanceID / No contents / Incompatible directions"
            Case 703
                UPnP_Error = "Invalid NewTagValue / Read error / Insufficient network resources"
            Case 704
                UPnP_Error = "Required tag / Format not supported for playback / Local restrictions"
            Case 705
                UPnP_Error = "Read only tag / Transport is locked / Access denied"
            Case 706
                UPnP_Error = "Paramater Mismatch / Write Error / Invalid connection reference"
            Case 707
                UPnP_Error = "Media is protected / Not writable / Not in Network"
            Case 708
                UPnP_Error = "Unsupported or invalid search criteria / Format no supported for recording"
            Case 709
                UPnP_Error = "Unsupported or invalid sort criteria / Media is full"
            Case 710
                UPnP_Error = "No such container / Seek mode not supported"
            Case 711
                UPnP_Error = "Restricted object / Illegal seek target"
            Case 712
                UPnP_Error = "Bad metadata / Playmode not supported"
            Case 713
                UPnP_Error = "Restricted parent object"
            Case 714
                UPnP_Error = "No such source resource / Invalid MIME-type"
            Case 715
                UPnP_Error = "Source resource access denied / Content 'BUSY'"
            Case 716
                UPnP_Error = "Transfer busy / Resource not found"
            Case 717
                UPnP_Error = "No such file transfer / Play speed not supported"
            Case 718
                UPnP_Error = "No such destination source / Invalid InstanceID"
            Case 719
                UPnP_Error = "Destination resource access denied"
            Case 720
                UPnP_Error = "Cannot process the request"
            Case 801
                UPnP_Error = "Access denied"
            Case 802
                UPnP_Error = "Not Enough Room"
            Case Else
                If ErrNumber >= 600 And ErrNumber <= 699 Then
                    UPnP_Error = "Common action error - undefined. Error=" & SaveError.ToString
                ElseIf ErrNumber >= 800 And ErrNumber <= 899 Then
                    UPnP_Error = "DLNA error - undefined. Error=" & SaveError.ToString
                Else
                    UPnP_Error = ErrNumber & ": Unknown error type. OrgError=" & SaveError.ToString
                End If
        End Select
    End Function

    Private Function ConvertBase16(Input As Integer) As String
        ConvertBase16 = ""
        If Input >= 0 And Input <= 9 Then
            ConvertBase16 = Chr(Input + &H30)
        ElseIf Input = 10 Then
            ConvertBase16 = "A"
        ElseIf Input = 11 Then
            ConvertBase16 = "B"
        ElseIf Input = 12 Then
            ConvertBase16 = "C"
        ElseIf Input = 13 Then
            ConvertBase16 = "D"
        ElseIf Input = 14 Then
            ConvertBase16 = "E"
        ElseIf Input = 15 Then
            ConvertBase16 = "F"
        End If
    End Function

    Private Sub ProcessXMLDocument_(URLDoc As String) ' this procedure is becoming void
        Dim xmlDoc As New XmlDocument
        'Dim HttpAddr As String = DeriveIPAddress(URLDoc)
        Dim PageHTML As String = ""
        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim webRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            webRequest.Method = "GET"
            webRequest.KeepAlive = False
            Dim webResponse As WebResponse = webRequest.GetResponse
            Dim webStream As Stream = webResponse.GetResponseStream
            Dim strmRdr As New System.IO.StreamReader(webStream)
            PageHTML = strmRdr.ReadToEnd()
            strmRdr.Dispose()
            webResponse.Close()
            webStream.Close()
            webStream.Dispose()
        Catch ex As Exception
            Log("Error in ProcessXMLDocument for device = " & MyUPnPDeviceName & " while retieving document with URL = " & URLDoc & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        PageHTML = RemoveControlCharacters(PageHTML)
        Try
            xmlDoc.LoadXml(PageHTML)
            Log(" ProcessXMLDocument for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in ProcessXMLDocument for device = " & MyUPnPDeviceName & " while loading XML with URL = " & URLDoc & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim Index As Integer = 0
        Dim ServiceURL As String = ""
        Dim ServiceXmlDoc As New XmlDocument
        Do While Index < 20 ' can't be that there are 100 services
            Dim ServiceHTML As String = ""
            Try
                ServiceURL = xmlDoc.GetElementsByTagName("SCPDURL").Item(Index).InnerText
                Dim RequestUri = New Uri(DeriveIPAddress(URLDoc, ServiceURL) & ServiceURL)
                Dim p = ServicePointManager.FindServicePoint(RequestUri)
                p.Expect100Continue = False
                Dim ServicewebRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
                ServicewebRequest.Method = "GET"
                ServicewebRequest.KeepAlive = False
                Dim ServicewebResponse As WebResponse = ServicewebRequest.GetResponse
                Dim ServicewebStream As Stream = ServicewebResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(ServicewebStream)
                ServiceHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
                ServicewebResponse.Close()
                ServicewebStream.Close()
                ServicewebStream.Dispose()
                ServicewebResponse = Nothing
            Catch ex As Exception
                xmlDoc = Nothing
                ServiceXmlDoc = Nothing
                Exit Do
            End Try
            ServiceHTML = RemoveControlCharacters(ServiceHTML)
            Try
                ServiceXmlDoc.LoadXml(ServiceHTML)
                Log(" ProcessXMLDocument for device = " & MyUPnPDeviceName & " retrieved following  Service document = " & ServiceXmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in ProcessXMLDocument for device = " & MyUPnPDeviceName & " while loading Service XML with URL = " & DeriveIPAddress(URLDoc, ServiceURL) & " and ServiceURL = " & ServiceURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            Index = Index + 1
        Loop
        xmlDoc = Nothing
        ServiceXmlDoc = Nothing
    End Sub

    Private Function VoidExtractIPInfo(DocumentURL As String) As IPAddressInfo
        Dim HttpIndex As Integer = DocumentURL.ToUpper.IndexOf("HTTP://")
        VoidExtractIPInfo.IPAddress = ""
        VoidExtractIPInfo.IPPort = ""
        Try
            If HttpIndex = -1 Then
                Log("ERROR in ExtractIPInfo for UPnPDevice = " & MyUPnPDeviceName & ". Not HTTP:// found. URL = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SubStr As String
            SubStr = DocumentURL.Substring(HttpIndex + 7, DocumentURL.Length - HttpIndex - 7)
            ' substring should now be primed for an IP address in the form of 192.168.1.1
            ' The next forward slash marks the end of the IP address, this could include the Port!
            Dim SlashIndex As Integer = SubStr.IndexOf("/")
            If HttpIndex = -1 Then
                Log("ERROR in ExtractIPInfo for UPnPDevice = " & MyUPnPDeviceName & ". No delimiting end / found. URL = " & DocumentURL & " and Substring = " & SubStr, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            SubStr = SubStr.Substring(0, SlashIndex)
            SubStr = SubStr.Trim
            If SubStr = "" Then
                Log("ERROR in ExtractIPInfo for UPnPDevice = " & MyUPnPDeviceName & ". No IP address found = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SemiCollonIndex As Integer = SubStr.IndexOf(":")
            'log( "ExtractIPInfo for UPnPDevice = " & MyUPnPDeviceName & " has Substring = " & SubStr)
            If SemiCollonIndex <> -1 Then
                ' there is an IP address and a Port Number
                VoidExtractIPInfo.IPAddress = SubStr.Substring(0, SemiCollonIndex)
                VoidExtractIPInfo.IPPort = SubStr.Substring(SemiCollonIndex + 1, SubStr.Length - SemiCollonIndex - 1)
            Else
                ' only IP address
                VoidExtractIPInfo.IPAddress = SubStr
            End If
        Catch ex As Exception
            Log("ERROR in ExtractIPInfo for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function FindActionInServiceDocument(Service As String, Action As String) As Boolean
        FindActionInServiceDocument = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindActionInServiceDocument called for device = " & MyUPnPDeviceName & " with Service = " & Service.ToString & " and Action = " & Action.ToString & " and DocumentURL = " & MyDocumentURL.ToString, LogType.LOG_TYPE_INFO)
        If MyDocumentURL = "" Then Exit Function
        Dim xmlDoc As New XmlDocument
        xmlDoc.XmlResolver = Nothing
        'Dim HttpAddr As String = DeriveIPAddress(MyDocumentURL)
        Dim PageHTML As String = ""
        Try
            Dim RequestUri = New Uri(MyDocumentURL)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim webRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            webRequest.Method = "GET"
            webRequest.KeepAlive = False
            Dim webResponse As WebResponse = webRequest.GetResponse
            Dim webStream As Stream = webResponse.GetResponseStream
            Dim strmRdr As New System.IO.StreamReader(webStream)
            PageHTML = strmRdr.ReadToEnd()
            strmRdr.Dispose()
            webStream.Close()
            webResponse.Close()
            webStream.Dispose()
            webResponse = Nothing
        Catch ex As Exception
            Log("Error in FindActionInServiceDocument for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        PageHTML = RemoveControlCharacters(PageHTML)
        Try
            xmlDoc.LoadXml(PageHTML)
            Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in FindActionInServiceDocument for device = " & MyUPnPDeviceName & " while loading XML with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim Index As Integer = 0
        Dim ServiceURL As String = ""
        Dim ServiceXmlDoc As New XmlDocument
        Do While Index < 100 ' can't be that there are 100 services
            Try
                Dim ServiceId As String = xmlDoc.GetElementsByTagName("serviceId").Item(Index).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " found ServiceId = " & ServiceId.ToString, LogType.LOG_TYPE_INFO)
                If xmlDoc.GetElementsByTagName("serviceId").Item(Index).InnerText = Service Then
                    ServiceURL = xmlDoc.GetElementsByTagName("SCPDURL").Item(Index).InnerText
                    ServiceURL = Trim(ServiceURL)
                    If ServiceURL = "" Then Exit Do
                    'If ServiceURL(0) = "/" Or ServiceURL(0) = "\" Then
                    ' we need to remove this
                    'Mid(ServiceURL, 0, 1) = " "
                    'ServiceURL = Trim(ServiceURL)
                    'End If
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " is creating ServiceURL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL, LogType.LOG_TYPE_INFO)
                    Dim ServiceHTML As String = ""
                    Try
                        Dim RequestUri = New Uri(DeriveIPAddress(MyDocumentURL, ServiceURL) & ServiceURL)
                        Dim p = ServicePointManager.FindServicePoint(RequestUri)
                        p.Expect100Continue = False
                        Dim ServicewebRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
                        ServicewebRequest.Method = "GET"
                        ServicewebRequest.KeepAlive = False
                        Dim ServicewebResponse As WebResponse = ServicewebRequest.GetResponse
                        Dim ServicewebStream As Stream = ServicewebResponse.GetResponseStream
                        Dim strmRdr As New System.IO.StreamReader(ServicewebStream)
                        ServiceHTML = strmRdr.ReadToEnd()
                        strmRdr.Dispose()
                        ServicewebStream.Close()
                        ServicewebResponse.Close()
                        ServicewebStream.Dispose()
                        ServicewebRequest = Nothing
                        ServicewebResponse = Nothing
                    Catch ex As Exception
                        Log("Error in FindActionInServiceDocument for device = " & MyUPnPDeviceName & " while retrieving Service XML with URL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    ServiceHTML = RemoveControlCharacters(ServiceHTML)
                    Try
                        ServiceXmlDoc.LoadXml(ServiceHTML)
                    Catch ex As Exception
                        Log("Error in FindActionInServiceDocument for device = " & MyUPnPDeviceName & " while loading Service XML with URL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " retrieved following  Service document = " & ServiceXmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
                    Dim ActionIndex As Integer = 0
                    Do While ActionIndex < 100
                        Dim actiontype As String = ServiceXmlDoc.GetElementsByTagName("name").Item(ActionIndex).InnerText
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " found actiontype = " & actiontype.ToString, LogType.LOG_TYPE_INFO)
                        If ServiceXmlDoc.GetElementsByTagName("name").Item(ActionIndex).InnerText = Action Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(" FindActionInServiceDocument for device = " & MyUPnPDeviceName & " found service = " & Service.ToString, LogType.LOG_TYPE_INFO)
                            FindActionInServiceDocument = True
                            xmlDoc = Nothing
                            ServiceXmlDoc = Nothing
                            Exit Function
                        End If
                        ActionIndex = ActionIndex + 1
                    Loop
                End If
            Catch ex As Exception
                xmlDoc = Nothing
                ServiceXmlDoc = Nothing
                Exit Do
            End Try
            Index = Index + 1
        Loop
        xmlDoc = Nothing
        ServiceXmlDoc = Nothing
    End Function

    Private Sub ProcessServiceDocument(pDevice As MyUPnPDevice)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument called for device = " & MyUPnPDeviceName & " with URL = " & MyDocumentURL, LogType.LOG_TYPE_INFO)
        'If MyDocumentURL = "" Then Exit Sub
        If pDevice Is Nothing Then Exit Sub
        Dim xmlDoc As New XmlDocument
        xmlDoc.XmlResolver = Nothing
        Dim PageHTML As String = pDevice.DeviceUPnPDocument
        If PageHTML = "" Then Exit Sub
        'Try
        'Dim webRequest As HttpWebRequest = HttpWebRequest.Create(MyDocumentURL)
        'webRequest.Method = "GET"
        'webRequest.KeepAlive = False
        'Dim webResponse As WebResponse = webRequest.GetResponse
        'Dim webStream As Stream = webResponse.GetResponseStream
        'Dim strmRdr As New System.IO.StreamReader(webStream)
        'PageHTML = strmRdr.ReadToEnd()
        'strmRdr.Dispose()
        'webStream.Close()
        'webResponse.Close()
        'webStream.Dispose()
        'webResponse = Nothing
        'Catch ex As Exception
        'Log("Error in ProcessServiceDocument for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        'Exit Sub
        'End Try

        PageHTML = RemoveControlCharacters(PageHTML)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " is retrieving following DeviceUPnPDocument = " & PageHTML, LogType.LOG_TYPE_INFO) 

        Try
            xmlDoc.LoadXml(PageHTML)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO) 
        Catch ex As Exception
            Log("Error in ProcessServiceDocument for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Dim Index As Integer = 0
        Dim ServiceURL As String = ""
        Dim ServiceXmlDoc As New XmlDocument
        Try
            Do While Index < 100 ' can't be that there are 100 services
                Try
                    Dim ServiceId As String = xmlDoc.GetElementsByTagName("serviceId").Item(Index).InnerText
                    Try
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "ProcessServiceDocument for device = " & MyUPnPDeviceName & " found ServiceId = " & ServiceId.ToString)
                        ServiceURL = xmlDoc.GetElementsByTagName("SCPDURL").Item(Index).InnerText
                        ServiceURL = Trim(ServiceURL)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " found ServiceId = " & ServiceId.ToString & " and ServiceURL = " & ServiceURL.ToString, LogType.LOG_TYPE_INFO)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " found ServiceId = " & ServiceId.ToString & " and ServiceURL = " & ServiceURL.ToString, LogType.LOG_TYPE_INFO)
                        If ServiceURL <> "" Then
                            'If ServiceURL(0) = "/" Or ServiceURL(0) = "\" Then
                            ' we need to remove this
                            'Mid(ServiceURL, 1, 1) = " "
                            'ServiceURL = Trim(ServiceURL)
                            'If piDebuglevel > DebugLevel.dlEvents Then log( "ProcessServiceDocument for device = " & MyUPnPDeviceName & " adjusted ServiceURL = " & ServiceURL.ToString)
                            'End If
                            If ServiceId.ToString = "urn:schemas-sony-com:serviceId:IRCC" Then
                                Dim ActionListURL As String = ""
                                Try
                                    ActionListURL = xmlDoc.GetElementsByTagName("av:X_CERS_ActionList_URL").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                                If ActionListURL <> "" Then
                                    Try
                                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " found ActionListURL = " & ActionListURL, LogType.LOG_TYPE_INFO)
                                        RetrieveSonyActionList(ActionListURL)
                                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " found MySonySystemInformationURL = " & MySonySystemInformationURL, LogType.LOG_TYPE_INFO) 
                                        RetrieveSonySystemInformation(MySonySystemInformationURL)
                                        If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                                            ' check if still registered
                                            If Not SonyRegister(MySonyRegisterURL, True) Then
                                                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                                                SetHSRemoteState()
                                            End If
                                        End If
                                    Catch ex As Exception
                                        Log("Error in ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieving Actionlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    End Try
                                Else
                                    Try
                                        SonyProcessIRCCInfo(PageHTML)
                                    Catch ex As Exception
                                        Log("Error in ProcessServiceDocument3 for device = " & MyUPnPDeviceName & " and Index = " & Index.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    End Try

                                    If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                                        ' check if still registered
                                        '   SendJSONAuthentication()
                                    End If
                                End If
                                Try
                                    MySonyRDISEntryPort = xmlDoc.GetElementsByTagName("av:X_RDIS_ENTRY_PORT").Item(0).InnerText
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieved SonyRDISEntryPort = " & MySonyRDISEntryPort, LogType.LOG_TYPE_INFO)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    Catch ex As Exception
                        Log("Error in ProcessServiceDocument2 for device = " & MyUPnPDeviceName & " and Index = " & Index.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " is creating ServiceURL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL, LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessServiceDocument for device = " & MyUPnPDeviceName & " is creating ServiceURL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL, LogType.LOG_TYPE_INFO) 

                    If ServiceURL <> "" Then
                        Dim ServiceHTML As String = ""
                        Try
                            Dim RequestUri = New Uri(DeriveIPAddress(MyDocumentURL, ServiceURL) & ServiceURL)
                            Dim p = ServicePointManager.FindServicePoint(RequestUri)
                            p.Expect100Continue = False
                            Dim ServicewebRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
                            ServicewebRequest.Method = "GET"
                            ServicewebRequest.KeepAlive = False
                            Dim ServicewebResponse As WebResponse = ServicewebRequest.GetResponse
                            Dim ServicewebStream As Stream = ServicewebResponse.GetResponseStream
                            Dim strmRdr As New System.IO.StreamReader(ServicewebStream)
                            ServiceHTML = strmRdr.ReadToEnd()
                            strmRdr.Dispose()
                            ServicewebResponse.Close()
                            ServicewebStream.Close()
                            ServicewebStream.Dispose()
                            ServicewebResponse = Nothing
                            ServicewebRequest = Nothing
                        Catch ex As Exception
                            Log("Error in ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieving the ServiceXML with URL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            ServiceHTML = ""
                        End Try
                        If ServiceHTML <> "" Then
                            ServiceHTML = RemoveControlCharacters(ServiceHTML)
                            Try
                                ServiceXmlDoc.LoadXml(ServiceHTML)
                            Catch ex As Exception
                                Log("Error in ProcessServiceDocument for device = " & MyUPnPDeviceName & " loading the ServiceXML with URL = " & DeriveIPAddress(MyDocumentURL, ServiceURL) & " and ServiceURL = " & ServiceURL & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                            If PIDebuglevel > DebugLevel.dlEvents Then Log(" ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieved following Service document = " & ServiceXmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( " ProcessServiceDocument for device = " & MyUPnPDeviceName & " retrieved following Service document = " & ServiceXmlDoc.OuterXml.ToString)
                            Dim ActionIndex As Integer = 0
                            Do While ActionIndex < 100
                                Try
                                    Dim ActionType As String = ServiceXmlDoc.GetElementsByTagName("name").Item(ActionIndex).InnerText
                                    If PIDebuglevel > DebugLevel.dlEvents Then Log(" ProcessServiceDocument for device = " & MyUPnPDeviceName & " found actiontype = " & ActionType.ToString, LogType.LOG_TYPE_INFO)
                                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( " ProcessServiceDocument for device = " & MyUPnPDeviceName & " found actiontype = " & ActionType.ToString)
                                    SetServiceFlags(ServiceId, ActionType)
                                    ActionIndex = ActionIndex + 1
                                Catch ex As Exception
                                    Exit Do
                                End Try
                            Loop
                        End If
                    End If
                Catch ex As Exception
                    xmlDoc = Nothing
                    ServiceXmlDoc = Nothing
                    Exit Do
                End Try
                Index = Index + 1
            Loop
        Catch ex As Exception
            Log("Error in ProcessServiceDocument3 for device = " & MyUPnPDeviceName & "  with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        xmlDoc = Nothing
        ServiceXmlDoc = Nothing
    End Sub

    Private Function SearchServiceDocument(DocumentURL As String, ServiceID As String, Action As String, Argument As String) As String
        SearchServiceDocument = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchServiceDocument called for device = " & MyUPnPDeviceName & " with URL = " & DocumentURL & " and ServiceID = " & ServiceID & " and Action = " & Action & " and Argument = " & Argument, LogType.LOG_TYPE_INFO)
        If DocumentURL = "" Or ServiceID = "" Or Action = "" Or Argument = "" Then Exit Function
        If MyUPnPDevice.Services Is Nothing Then Exit Function
        Dim MyService As MyUPnPService = MyUPnPDevice.Services.Item(ServiceID)
        If MyService Is Nothing Then Exit Function
        Dim MyActionList As MyactionList = MyService.ActionList
        If MyActionList Is Nothing Then Exit Function
        Dim MyAction As MyUPnPAction = MyActionList.Item(Action)
        If MyAction Is Nothing Then Exit Function
        '<action>
        '  <name>SelectPreset</name>
        '  <argumentList>
        '    <argument>
        '      <name>InstanceID</name>
        '      <direction>in</direction>
        '      <relatedStateVariable>A_ARG_TYPE_InstanceID</relatedStateVariable>
        '    </argument>
        '    <argument>
        '      <name>PresetName</name>
        '      <direction>in</direction>
        '      <relatedStateVariable>A_ARG_TYPE_PresetName</relatedStateVariable>
        '    </argument>
        '  </argumentList>
        '</action>
        Dim RelatedStateVariable As String = ""
        Dim FoundIt As Boolean = False
        Try
            Dim ArgumentListXMLDoc As New XmlDocument
            ArgumentListXMLDoc.LoadXml(MyAction.argumentList)
            Dim ArgumentNodeList As XmlNodeList = ArgumentListXMLDoc.GetElementsByTagName("argument")
            For Each Argument_ As System.Xml.XmlNode In ArgumentNodeList
                For Each child As XmlElement In Argument_
                    If child.Name = "name" Then
                        If child.InnerText = Argument Then
                            ' we got it
                            FoundIt = True
                        End If
                    ElseIf child.Name = "relatedStateVariable" Then
                        RelatedStateVariable = child.InnerText
                    End If
                Next
                If FoundIt Then Exit For
            Next

            If RelatedStateVariable = "" Then Exit Function

            Dim MyServiceStateTable As MyServiceStateTable = MyService.ServiceStateTable
            If MyServiceStateTable Is Nothing Then Exit Function

            Dim MyServiceStateVar As MyStateVariable = MyServiceStateTable.Item(RelatedStateVariable)
            If MyServiceStateVar Is Nothing Then Exit Function

            '<stateVariable sendEvents="no">
            '  <name>CurrentPlayMode</name>
            '  <dataType>string</dataType>
            '  <defaultValue>NORMAL</defaultValue>
            '  <allowedValueList>
            '    <allowedValue>NORMAL</allowedValue>
            '    <allowedValue>RANDOM</allowedValue>
            '    <allowedValue>REPEAT_ONE</allowedValue>
            '    <allowedValue>REPEAT_ALL</allowedValue>
            '  </allowedValueList>
            '</stateVariable>
            If MyServiceStateVar.allowedValueList IsNot Nothing Then
                Try
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchServiceDocument found AllowedValueList for device = " & MyUPnPDeviceName & " = " & MyServiceStateVar.allowedValueList, LogType.LOG_TYPE_INFO)
                    SearchServiceDocument = MyServiceStateVar.allowedValueList
                    Exit Function
                Catch ex As Exception
                End Try
            End If
            If MyServiceStateVar.allowedValueRange IsNot Nothing Then
                Try
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchServiceDocument found allowedValueRange for device = " & MyUPnPDeviceName & " = " & MyServiceStateVar.allowedValueRange, LogType.LOG_TYPE_INFO)
                    SearchServiceDocument = MyServiceStateVar.allowedValueRange
                    Exit Function
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            Log("Error in SearchServiceDocument for device = " & MyUPnPDeviceName & " while retieving document with URL = " & DocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Private Function GetXMLValues(inXML As String, Tag As String) As String()
        GetXMLValues = Nothing
        Dim TempResult As String() = Nothing
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetXMLValues called for device = " & MyUPnPDeviceName & " with inXML = " & inXML & " and Tag = " & Tag, LogType.LOG_TYPE_INFO)
        If inXML = "" Or Tag = "" Then Exit Function
        Dim XMLDoc As New XmlDocument
        Try
            XMLDoc.LoadXml(inXML)
        Catch ex As Exception
            Log("Error in GetXMLValues for device = " & MyUPnPDeviceName & " loading the xml with inXML = " & inXML & " and Tag = " & Tag & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            XMLDoc = Nothing
            Exit Function
        End Try

        Dim Index As Integer = 0

        Try
            Do While Index < 100 ' can't be that there are 100 
                Try
                    Dim FoundTagValue As String = XMLDoc.GetElementsByTagName(Tag).Item(Index).InnerText
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetXMLValues for device = " & MyUPnPDeviceName & " found matching Tag = " & Tag.ToString & " and TagValue = " & FoundTagValue.ToString, LogType.LOG_TYPE_INFO)
                    If TempResult Is Nothing Then
                        ReDim TempResult(0)
                    Else
                        ReDim Preserve TempResult(UBound(TempResult, 1) + 1)
                    End If
                    TempResult(UBound(TempResult, 1)) = FoundTagValue
                Catch ex As Exception
                    'log( "Error in GetXMLValues 1 for device = " & MyUPnPDeviceName & "  with error = " &ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try
                Index = Index + 1
            Loop
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning in GetXMLValues for device = " & MyUPnPDeviceName & "  with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
        XMLDoc = Nothing
        If PIDebuglevel > DebugLevel.dlEvents Then
            If TempResult Is Nothing Then
                Log("GetXMLValues for device = " & MyUPnPDeviceName & " found no values for Tag = " & Tag.ToString, LogType.LOG_TYPE_INFO)
            Else
                Log("GetXMLValues for device = " & MyUPnPDeviceName & " found " & (UBound(TempResult, 1) + 1).ToString & " values for Tag = " & Tag.ToString, LogType.LOG_TYPE_INFO)
            End If
        End If

        Return TempResult

    End Function

    Private Sub SetServiceFlags(ServiceId As String, ActionType As String)
        Select Case ServiceId
            Case "urn:upnp-org:serviceId:RenderingControl"
                Select Case ActionType
                    Case "GetBrightness"
                        BrightnessIsConfigurable = True
                    Case "GetColorTemperature"
                        ColorTemperatureIsConfigurable = True
                    Case "GetContrast"
                        ContrastIsConfigurable = True
                    Case "GetLoudness"
                        LoudnessIsConfigurable = True
                    Case "GetMute"
                        MuteIsConfigurable = True
                    Case "GetSharpness"
                        SharpnessIsConfigurable = True
                    Case "GetVolume"
                        VolumeIsConfigurable = True
                    Case "GetVolumeDB"
                        VolumeDBIsConfigurable = True
                    Case "X_GetImageRotation"
                        GetImageRotationIsConfigurable = True
                    Case "X_GetImageScale"
                        GetImageScaleIsConfigurable = True
                    Case "X_GetSlideShowEffect"
                        GetSlideShowEffectIsConfigurable = True
                    Case Else
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("SetServiceFlags for device - " & MyUPnPDeviceName & " found unkown ActionType = " & ActionType & " for Service = " & ServiceId, LogType.LOG_TYPE_INFO)
                End Select
            Case "urn:upnp-org:serviceId:AVTransport"
                Select Case ActionType
                    Case "SetNextAVTransportURI"
                        NextAvTransportIsAvailable = True
                        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diNextAV.ToString, True)
                        If GetStringIniFile(MyUDN, DeviceInfoIndex.diUseNextAV.ToString, "") = "" Then
                            WriteBooleanIniFile(DeviceUDN, DeviceInfoIndex.diUseNextAV.ToString, True)
                        End If
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetServiceFlags for device - " & MyUPnPDeviceName & " found ActionType = " & ActionType, LogType.LOG_TYPE_INFO)
                    Case "SetPlayMode"
                        PlayModeisConfigurable = True
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetServiceFlags for device - " & MyUPnPDeviceName & " found ActionType = " & ActionType, LogType.LOG_TYPE_INFO)
                End Select
            Case "urn:upnp-org:serviceId:ContentDirectory"
                Select Case ActionType
                    Case "Search"
                        SupportSearch = True
                End Select
            Case Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SetServiceFlags for device - " & MyUPnPDeviceName & " found unkown Service = " & ServiceId & " and ActionType = " & ActionType, LogType.LOG_TYPE_INFO)
        End Select
    End Sub

    Public Function ConvertSecondsToTimeFormat(ByVal Seconds As Integer) As String
        ConvertSecondsToTimeFormat = "00:00:00"
        If Seconds < 0 Then Exit Function
        Dim StartTime As Date = CDate("00:00:00")
        ConvertSecondsToTimeFormat = Format(DateAdd("s", CType(Seconds, Double), StartTime), "HH:mm:ss")
    End Function

    Public Function ConvertTimeFormatToSeconds(inTime As String) As Integer
        ConvertTimeFormatToSeconds = 0
        If inTime = "" Then Exit Function
        If inTime.ToUpper = "NOT_IMPLEMENTED" Then Exit Function
        Dim TimeParts As String() = Split(inTime, ":")
        If TimeParts IsNot Nothing Then
            If UBound(TimeParts) = 0 Then
                inTime = "00:00:" & inTime
            ElseIf UBound(TimeParts) = 1 Then
                inTime = "00:" & inTime
            End If
            Try
                Dim ts As TimeSpan = TimeSpan.Parse(inTime)
                ConvertTimeFormatToSeconds = CType(ts.TotalSeconds, Integer)
            Catch ex As Exception
            End Try
        End If
    End Function

    Public Sub ReadDeviceIniSettings()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadDeviceIniSettings called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim AnythingStored As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diTimeBetweenPictures.ToString, "")
            If AnythingStored = "" Then WriteIntegerIniFile(MyUDN, DeviceInfoIndex.diTimeBetweenPictures.ToString, 60) ' set to 1 minute as a default
            MyQueueDelayUserDefined = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diTimeBetweenPictures.ToString, 0)
            If MyQueueDelayUserDefined = 0 Then
                ' this should not be
                MyQueueDelayUserDefined = 60
            End If
        Catch ex As Exception
            Log("Error in ReadDeviceIniSettings reading diTimeBetweenPictures for device " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyQueueDelayUserDefined = 60
        End Try
        Try
            MyQueueRepeatState = GetBooleanIniFile(MyUDN, DeviceInfoIndex.diQueueRepeat.ToString, False)
        Catch ex As Exception
            Log("Error in ReadDeviceIniSettings reading diQueueRepeat for device " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyQueueRepeatState = False
        End Try
        Try
            MyQueueShuffleState = GetBooleanIniFile(MyUDN, DeviceInfoIndex.diQueueShuffle.ToString, False)
        Catch ex As Exception
            Log("Error in ReadDeviceIniSettings reading diQueueShuffle for device " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyQueueShuffleState = False
        End Try
        Try
            MyServerUDN = GetStringIniFile(MyUDN, DeviceInfoIndex.diServerUDN.ToString, "")
        Catch ex As Exception
            Log("Error in ReadDeviceIniSettings reading diServerUDN for device " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MyPictureSize = GetIntegerIniFile(MyUDN, DeviceInfoIndex.diPictureSize.ToString, 2)
        Catch ex As Exception
            MyPictureSize = PictureSize.psMedium
        End Try
        Try
            UseNextAvTransport = GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diUseNextAV.ToString, False)
        Catch ex As Exception
            Log("Error in ReadDeviceIniSettings reading diUseNextAV for device " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diPollTransportChanges.ToString, False) Then
            StartAVPollTimer()     ' changed 12/2/2018 v.039
        Else
            StopAVPollTimer()
        End If
        If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diPollVolumeChanges.ToString, False) Then
            StartRenderPollTimer()    ' changed 12/2/2018 v0.39
        Else

        End If
    End Sub

    Public Sub SetTrackLength(ByVal TrackLength As Integer)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SetTrackLength called for device - " & MyUPnPDeviceName & " with TrackLength = " & TrackLength, LogType.LOG_TYPE_INFO)
        Try
            If MyTrackLength <> TrackLength Then
                If HSRefTrackLength <> -1 Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetTrackLength is setting HS Device TrackLength for device - " & MyUPnPDeviceName & " with Value = " & TrackLength.ToString & " and HSRef = " & HSRefTrackLength.ToString, LogType.LOG_TYPE_INFO)
                    Select Case MyHSTrackLengthFormat
                        Case HSSTrackLengthSettings.TLSSeconds
                            hs.SetDeviceString(HSRefTrackLength, TrackLength.ToString, True)
                        Case HSSTrackLengthSettings.TLSHoursMinutesSeconds
                            hs.SetDeviceString(HSRefTrackLength, ConvertSecondsToTimeFormat(TrackLength), True)
                    End Select
                    hs.SetDeviceValueByRef(HSRefTrackLength, CType(TrackLength, Double), True)
                    If HSRefTrackPos <> -1 Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetTrackLength is setting HS Status Max Position TrackPosition for device - " & MyUPnPDeviceName & " and HSRef = " & HSRefTrackPos.ToString, LogType.LOG_TYPE_INFO)
                        ' update the slider control pair
                        Dim VSVGPair As VSPair
                        VSVGPair = hs.DeviceVSP_Get(HSRefTrackPos, 0, ePairStatusControl.Both) ' use value = 0 to be within the range
                        VSVGPair.Render_Location.Column = 2
                        VSVGPair.Render_Location.Row = 1
                        If VSVGPair.Render = HomeSeerAPI.Enums.CAPIControlType.ValuesRangeSlider Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetTrackLength set Pair for device - " & MyUPnPDeviceName & " Old Max Range = " & VSVGPair.RangeEnd.ToString & " New Max Range = " & TrackLength.ToString, LogType.LOG_TYPE_INFO)
                            If MyHSTrackPositionFormat <> HSSTrackPositionSettings.TPSPercentage Then
                                VSVGPair.RangeEnd = CType(TrackLength, Double)
                            Else
                                VSVGPair.RangeEnd = 100
                            End If
                            hs.DeviceVSP_ClearAny(HSRefTrackPos, 0)
                            hs.DeviceVSP_AddPair(HSRefTrackPos, VSVGPair)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Log("Error in SetTrackLength set Pair for device - " & MyUPnPDeviceName & " with TrackLength = " & TrackLength & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyTrackLength = TrackLength
        CurrentLibEntry.LengthSeconds = TrackLength

    End Sub

    Private Function GetSeconds(ByVal Time As String) As Integer
        GetSeconds = 0
        If Time = "" Then Exit Function
        Try
            Dim RemovePartialsBehindDot As String()
            RemovePartialsBehindDot = Split(Time, ".")
            Dim Conversion As String()
            Conversion = Split(RemovePartialsBehindDot(0), ":")
            GetSeconds = (CInt(Conversion(0)) * 60 * 60) + (CInt(Conversion(1)) * 60) + CInt(Conversion(2))
        Catch ex As Exception
            GetSeconds = 0
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSeconds for device = " & MyUPnPDeviceName & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function RemoveControlCharacters(inString As String) As String
        RemoveControlCharacters = inString
        Dim OutString As String = ""
        Try
            ' remove any control characters
            Dim strIndex As Integer = inString.Length
            If PIDebuglevel > DebugLevel.dlEvents Then Log("RemoveControlCharacters for device = " & MyUPnPDeviceName & " retrieved document with length = " & strIndex.ToString, LogType.LOG_TYPE_INFO)
            Dim SomethingGotRemoved As Boolean = False
            While strIndex > 0
                strIndex = strIndex - 1
                If inString(strIndex) < " " Then
                    inString = inString.Remove(strIndex, 1)
                    SomethingGotRemoved = True
                End If
            End While
            inString = Trim(inString)
            If PIDebuglevel > DebugLevel.dlEvents And SomethingGotRemoved Then Log("RemoveControlCharacters for device = " & MyUPnPDeviceName & " updated document to = " & inString.ToString, LogType.LOG_TYPE_INFO)
            RemoveControlCharacters = inString
        Catch ex As Exception
            Log("Error in RemoveControlCharacters for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function DeriveIPAddress(inString As String, NextChar As String) As String
        If PIDebuglevel > DebugLevel.dlEvents Then Log("DeriveIPAddress called for Device = " & MyUPnPDeviceName & " and inString = " & inString & " and NextChar = " & NextChar, LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeriveIPAddress called for Device = " & MyUPnPDeviceName & " and inString = " & inString & " and NextChar = " & NextChar, LogType.LOG_TYPE_INFO) 
        DeriveIPAddress = inString
        Dim NewURLDoc As String = Trim(inString)
        NextChar = Trim(NextChar)
        If NextChar.ToUpper.IndexOf("HTTP://") <> -1 Or NextChar.ToUpper.IndexOf("HTTP:\\") <> -1 Then
            ' The NextChar already has a full URL definition
            DeriveIPAddress = ""
        Else
            Dim httpIndex As Integer = 0
            If NextChar <> "" And (NextChar(0) <> "/" And NextChar(0) <> "\") Then
                httpIndex = NewURLDoc.LastIndexOf("/")
            ElseIf NewURLDoc.ToUpper.IndexOf("HTTP://") <> -1 Then
                NewURLDoc = NewURLDoc.Remove(0, 7)
                httpIndex = NewURLDoc.IndexOf("/") + 6
            ElseIf NewURLDoc.ToUpper.IndexOf("HTTP:\\") <> -1 Then
                NewURLDoc = NewURLDoc.Remove(0, 7)
                httpIndex = NewURLDoc.IndexOf("/") + 6
            Else
                httpIndex = NewURLDoc.LastIndexOf("/")
            End If
            If httpIndex <> 0 Then
                DeriveIPAddress = inString.Substring(0, httpIndex + 1)
            End If
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then Log("DeriveIPAddress for Device = " & MyUPnPDeviceName & " with inString = " & inString & " and NextChar = " & NextChar & " returned = " & DeriveIPAddress, LogType.LOG_TYPE_INFO)
    End Function

#Region "Music API methods"





    Public Property APIInstance As Integer 'Implements MediaCommon.MusicAPI.APIInstance
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Get APIInstance called for device - " & MyUPnPDeviceName & ". Index = " & MyHSTMusicIndex.ToString)
            APIInstance = MyHSTMusicIndex
        End Get
        Set(value As Integer)
            MyHSTMusicIndex = value
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set APIInstance called for device - " & MyUPnPDeviceName & ". Index = " & MyHSTMusicIndex.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property APIName As String 'Implements MediaCommon.MusicAPI.APIName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( " Get APIName called for device - " & MyUPnPDeviceName)
            APIName = MyUPnPDeviceName
        End Get
    End Property

    Public Function PlayerState() As player_state_values 'Implements MediaCommon.MusicAPI.PlayerState
        'Returns the current state of the player using the following Enum values:
        'Public Enum player_state_values
        '   playing = 1
        '   stopped = 2
        '   paused = 3
        '   forwarding = 4
        '   rewinding = 5
        'End Enum
        If PIDebuglevel > DebugLevel.dlEvents Then Log("PlayerState called for UPnPDevice - " & MyUPnPDeviceName & ". State = " & MyCurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        PlayerState = MyCurrentPlayerState
    End Function

    Public Function GetItems(ByVal ObjectID As String, Optional ByVal DownToItems As Boolean = False) As System.Array
        'Returns a list of all playlist names in the system.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetItems called for device - " & MyUPnPDeviceName & " with ObjectId = " & ObjectID & " and DownToItems = " & DownToItems.ToString, LogType.LOG_TYPE_INFO)
        GetItems = {""}
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetItems called for device - " & MyUPnPDeviceName & " has no ContentDirectory Handle", LogType.LOG_TYPE_WARNING)
            GetItems = Nothing
            Exit Function
        End If
        Dim MyItems As String()
        MyItems = {""}
        Dim Index As Integer = 0


        Dim InArg(5)
        Dim OutArg(3)

        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4 -- should be 1 but that gives errors for wireless doc Genres
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("Error in GetItems for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("Error in GetItems for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID, LogType.LOG_TYPE_ERROR)
            GetItems = Nothing
            Exit Function
        End Try

        Dim BrowseResult As String = OutArg(0)
        Dim BrowseNumberReturned As Integer = OutArg(1)
        Dim BrowseTotalMatches As Integer = OutArg(2)
        If BrowseNumberReturned < 1 Then
            Log("Error in GetItems called for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned = " & BrowseNumberReturned.ToString & " records", LogType.LOG_TYPE_ERROR)
            GetItems = Nothing
            Exit Function
        Else
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetItems called for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & " and found = " & BrowseTotalMatches.ToString & " records", LogType.LOG_TYPE_INFO)
        End If

        ReDim MyItems(CInt(BrowseTotalMatches) - 1)
        Dim StartIndex As Integer = 0

        Dim xmlData As XmlDocument = New XmlDocument
        Dim OuterXMLData As String
        Dim OuterXML As XmlDocument = New XmlDocument

        Do

            InArg(3) = StartIndex
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("Error in GetItems for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                BrowseNumberReturned = OutArg(1)
                For LoopIndex = 0 To BrowseNumberReturned - 1
                    OuterXMLData = ""
                    Try
                        OuterXMLData = xmlData.GetElementsByTagName("item").Item(LoopIndex).OuterXml
                        OuterXML.LoadXml(OuterXMLData)
                    Catch ex As Exception
                    End Try
                    If OuterXMLData = "" Then
                        ' probably a containter
                        Try
                            OuterXMLData = xmlData.GetElementsByTagName("container").Item(LoopIndex).OuterXml
                            OuterXML.LoadXml(OuterXMLData)
                        Catch ex As Exception
                        End Try
                    End If
                    If OuterXMLData <> "" Then
                        Try
                            Dim MyItemName As String = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                            MyItems(StartIndex + LoopIndex) = MyItemName
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetItems found for device - " & MyUPnPDeviceName & " Title = " & MyItemName.ToString)
                        Catch ex As Exception
                            MyItems(StartIndex + LoopIndex) = ""
                        End Try
                        'Try
                        'Dim MyItemClass As String = OuterXML.GetElementsByTagName("upnp:class").Item(0).InnerText
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetItems found for device - " & MyUPnPDeviceName & " Class = " & MyItemClass.ToString)
                        'Catch ex As Exception
                        'End Try
                    End If
                Next
            Catch ex As Exception
                Log("Error in GetItems for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + BrowseNumberReturned
            If StartIndex >= BrowseTotalMatches Then
                Exit Do
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetItems for device = " & MyUPnPDeviceName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            'hs.WaitEvents()
        Loop

        InArg = Nothing
        OutArg = Nothing
        xmlData = Nothing
        OuterXML = Nothing
        GetItems = MyItems
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetItems called for device - " & MyUPnPDeviceName & " returned " & BrowseTotalMatches.ToString & " playlists ", LogType.LOG_TYPE_INFO)
    End Function



    Public Function PrepareForQuery(ByVal inString As String) As String
        ' this function deals with ' in query names
        PrepareForQuery = inString.Replace("'", "''")
    End Function

#End Region

    Public Function CurrentlyPlaying() As HomeSeerAPI.Lib_Entry_Key Implements HomeSeerAPI.IMediaAPI.CurrentlyPlaying
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CurrentlyPlaying called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO, LogColorPink)
        Return Nothing
    End Function

    Public Function CurrentPlayList() As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.CurrentPlayList
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CurrentPlayList called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO, LogColorPink)
        Return Nothing
    End Function

    Public Function CurrentPlayListAdd(TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.CurrentPlayListAdd
        Return Nothing
    End Function

    Public Sub CurrentPlayListClear() Implements HomeSeerAPI.IMediaAPI.CurrentPlayListClear

    End Sub

    Public Function CurrentPlayListCount() As Integer Implements HomeSeerAPI.IMediaAPI.CurrentPlayListCount
        Return Nothing
    End Function

    Public Function CurrentPlayListRange(Start As Integer, Count As Integer) As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.CurrentPlayListRange
        Return Nothing
    End Function

    Public Function CurrentPlayListSet(TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.CurrentPlayListSet
        Return False
    End Function

    Public Function LibGetAlbums1(artist As String, genre As String, Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetAlbums
        Return Nothing
    End Function

    Public Function LibGetArtists1(album As String, genre As String, Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetArtists
        Return Nothing
    End Function

    Public Function LibGetEntry(Key As HomeSeerAPI.Lib_Entry_Key) As HomeSeerAPI.Lib_Entry Implements HomeSeerAPI.IMediaAPI.LibGetEntry
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LibGetEntry called for Player - " & MyUPnPDeviceName & " with iKey = " & Key.iKey & ", SKey = " & Key.sKey & ", title = " & Key.Title & ", Libaray = " & Key.Library & " whichKey = " & Key.WhichKey.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return Nothing
    End Function

    Public Function LibGetGenres1(Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetGenres
        Return Nothing
    End Function

    Public Function LibGetLibrary() As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibrary
        Return Nothing
    End Function

    Public Function LibGetLibrarybyLibType(Lib_Type As UShort) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibrarybyLibType
        Return Nothing
    End Function

    Public Function LibGetLibraryCount() As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCount
        Return 0
    End Function

    Public Function LibGetLibraryCountbyEntryType(EntryType As HomeSeerAPI.eLib_Media_Type) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyEntryType
        Return Nothing
    End Function

    Public Function LibGetLibraryCountbyLibType(Lib_Type As UShort) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyLibType
        Return 0
    End Function

    Public Function LibGetLibraryCountbyType(Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyType
        Return 0
    End Function

    Public Function LibGetLibraryRange(Start As Integer, Count As Integer) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRange
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyEntryType(Start As Integer, Count As Integer, EntryType As HomeSeerAPI.eLib_Media_Type) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyEntryType
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyLibType(Start As Integer, Count As Integer, LibType As UShort) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyLibType
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyType(Start As Integer, Count As Integer, LibType As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyType
        Return Nothing
    End Function

    Public Function LibGetLibraryTypes() As HomeSeerAPI.Lib_Type() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryTypes
        Return Nothing
    End Function

    Public Function LibGetPlaylists1(Optional Lib_Type As UShort = 0) As HomeSeerAPI.Playlist_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetPlaylists
        Return Nothing
    End Function

    Public Function LibGetPlaylistTracks1(Playlist As HomeSeerAPI.Playlist_Entry) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetPlaylistTracks
        Return Nothing
    End Function

    Public Function LibGetTracks1(artist As String, album As String, genre As String, Lib_Type As UShort) As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.LibGetTracks
        Return Nothing
    End Function

    Public ReadOnly Property LibLoading1 As Boolean Implements HomeSeerAPI.IMediaAPI.LibLoading
        Get
            Return False
        End Get
    End Property

    Public Sub Play1(Key As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.Play

    End Sub

    Public Sub PlayGenre(GenreName As String, Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) Implements HomeSeerAPI.IMediaAPI.PlayGenre

    End Sub

    Public Sub PlayGenreAt(GenreName As String, Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type, Start_Track As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.PlayGenreAt

    End Sub

    Public Function Playlist_Add(Playlist As HomeSeerAPI.Playlist_Entry) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add
        Return False
    End Function

    Public Function Playlist_Add_Track(Playlist As HomeSeerAPI.Playlist_Entry, TrackKey As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add_Track
        Return False
    End Function

    Public Function Playlist_Add_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add_Tracks
        Return False
    End Function

    Public Function Playlist_Delete(Playlist As HomeSeerAPI.Playlist_Entry) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete
        Return False
    End Function

    Public Function Playlist_Delete_Track(Playlist As HomeSeerAPI.Playlist_Entry, TrackKey As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete_Track
        Return False
    End Function

    Public Function Playlist_Delete_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete_Tracks
        Return False
    End Function

    Public Sub PlayPlaylist(Playlist As HomeSeerAPI.Playlist_Entry) Implements HomeSeerAPI.IMediaAPI.PlayPlaylist

    End Sub

    Public Sub PlayPlaylistAt(Playlist As HomeSeerAPI.Playlist_Entry, Start_Key As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.PlayPlaylistAt

    End Sub

    Public Sub PlayMatch(MatchInfo As HomeSeerAPI.Play_Match_Info) Implements HomeSeerAPI.IMediaAPI_2.PlayMatch

    End Sub

    Public Sub AdjustVolume(Amount As Integer, Optional Direction As HomeSeerAPI.eVolumeDirection = HomeSeerAPI.eVolumeDirection.Absolute) Implements HomeSeerAPI.IMediaAPI_3.AdjustVolume

    End Sub

    Public Function GetMatch(QData As HomeSeerAPI.Query_Object) As HomeSeerAPI.Response_Object Implements HomeSeerAPI.IMediaAPI_3.GetMatch
        Return Nothing
    End Function

    Public Sub Halt() Implements HomeSeerAPI.IMediaAPI_3.Halt

    End Sub

    Public Sub Mute(Mode As HomeSeerAPI.mute_modes) Implements HomeSeerAPI.IMediaAPI_3.Mute

    End Sub

    Public ReadOnly Property Muted As HomeSeerAPI.mute_modes Implements HomeSeerAPI.IMediaAPI_3.Muted
        Get
            Return False
        End Get
    End Property

    Public Sub Pause1() Implements HomeSeerAPI.IMediaAPI_3.Pause

    End Sub

    Public ReadOnly Property PlayerPosition1 As Integer Implements HomeSeerAPI.IMediaAPI_3.PlayerPosition
        Get
            Return 0
        End Get
    End Property

    Public Function Playlist_Add_Matched_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, MatchInfo As HomeSeerAPI.Play_Match_Info) As Boolean Implements HomeSeerAPI.IMediaAPI_3.Playlist_Add_Matched_Tracks
        Return False
    End Function

    Public Sub Repeat1(Mode As HomeSeerAPI.repeat_modes) Implements HomeSeerAPI.IMediaAPI_3.Repeat

    End Sub

    Public ReadOnly Property Repeating As HomeSeerAPI.repeat_modes Implements HomeSeerAPI.IMediaAPI_3.Repeating
        Get
            Return Nothing
        End Get
    End Property

    Public Sub SelectTrack(TrackKey As Integer, Optional PlaylistKey As Integer = -1) Implements HomeSeerAPI.IMediaAPI_3.SelectTrack

    End Sub

    Public Sub Shuffle1(Mode As HomeSeerAPI.shuffle_modes) Implements HomeSeerAPI.IMediaAPI_3.Shuffle

    End Sub

    Public ReadOnly Property Shuffled As HomeSeerAPI.shuffle_modes Implements HomeSeerAPI.IMediaAPI_3.Shuffled
        Get
            Return Nothing
        End Get
    End Property

    Public Sub SkipToTrack1(TrackName As String) Implements HomeSeerAPI.IMediaAPI_3.SkipToTrack

    End Sub

    Public Sub SkipTracks(SkipValue As Integer) Implements HomeSeerAPI.IMediaAPI_3.SkipTracks

    End Sub

    Public ReadOnly Property State As HomeSeerAPI.player_state_values Implements HomeSeerAPI.IMediaAPI_3.State
        Get
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("State called for device - " & MyUPnPDeviceName & " and State = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO, LogColorPink)
            State = HomeSeerAPI.player_state_values.playing
        End Get
    End Property

    Public ReadOnly Property Volume1 As Integer Implements HomeSeerAPI.IMediaAPI_3.Volume
        Get
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Volume1 called for device - " & MyUPnPDeviceName & " and Volume = " & Volume.ToString, LogType.LOG_TYPE_INFO, LogColorPink)
            Volume1 = Volume
        End Get
    End Property
End Class

Public Enum player_status_change
    SongChanged = 1         'raises whenever the current song changes
    PlayStatusChanged = 2   'raises when pause, stop, play, etc. pressed.
    PlayList = 3            'raises whenever the current playlist changes
    Library = 4             'raises when the library changes
End Enum


<Serializable()> _
Public Class LastMusic
    Public Sub New()

    End Sub
    Public Album As String
    Public Artist As String
    Public Genre As String
    Public iMode As Integer
    Public Playlist As String
    Public Track As String
    Public WasPlaying As Boolean
End Class

Public Class track_desc
    Public name As String
    Public artist As String
    Public album As String
    Public length As String
End Class


<Serializable()> _
Public Class myUPnPControlCallback
    Public Event ControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event ControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent ControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent ControlDied()
        Return 0
    End Function
End Class
