Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Xml
Imports System.Drawing
Imports System.IO
Imports System.Web.Script.Serialization
'Imports System.Object

Partial Public Class HSPI

    Dim LGWebSocket As WebSocketClient
    Dim LGWebPointerSocket As WebSocketClient
    Dim LGCommandCount As Integer = 0
    Const MaxButtonCollumns = 6

    Private Sub CreateLGRemoteIniFileInfo()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGRemoteIniFileInfo called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = gRemoteControlPath
        Dim SamsungRemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")

        WriteStringIniFile(MyUDN, "20", "PowerOff" & ":;:-:" & "turnoff" & ":;:-:1:;:-:1:;:-:" & "", objRemoteFile)

        WriteStringIniFile(MyUDN, "21", "Volume Up" & ":;:-:" & "inputvolumeup" & ":;:-:2:;:-:1:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "22", "Volume Down" & ":;:-:" & "inputvolumedown" & ":;:-:2:;:-:2:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "23", "Setmute On" & ":;:-:" & "setmute" & ":;:-:2:;:-:3:;:-:" & "true", objRemoteFile)
        WriteStringIniFile(MyUDN, "24", "Setmute Off" & ":;:-:" & "setmute" & ":;:-:2:;:-:4:;:-:" & "false", objRemoteFile)
        WriteStringIniFile(MyUDN, "25", "Setmute Togle" & ":;:-:" & "togglemute" & ":;:-:2:;:-:5:;:-:" & "", objRemoteFile)

        WriteStringIniFile(MyUDN, "26", "Play" & ":;:-:" & "inputmediaplay" & ":;:-:3:;:-:1:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "27", "Stop" & ":;:-:" & "inputmediastop" & ":;:-:3:;:-:2:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "28", "Pause" & ":;:-:" & "inputmediapause" & ":;:-:3:;:-:3:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "29", "Rewind" & ":;:-:" & "inputmediarewind" & ":;:-:3:;:-:4:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "30", "Forward" & ":;:-:" & "inputmediaforward" & ":;:-:3:;:-:5:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "31", "Enter" & ":;:-:" & "inputenter" & ":;:-:3:;:-:6:;:-:" & "", objRemoteFile)


        WriteStringIniFile(MyUDN, "32", "Channel Up" & ":;:-:" & "inputchannelup" & ":;:-:4:;:-:1:;:-:" & "", objRemoteFile)
        WriteStringIniFile(MyUDN, "33", "Channel Down" & ":;:-:" & "inputchanneldown" & ":;:-:4:;:-:2:;:-:" & "", objRemoteFile)

        ' This section is the SocketPointer commands
        '         ' REWIND GOTOPREV GOTONEXT FASTFORWARD STOP PLAY PAUSE RED GREEN YELLOW BLUE HOME UP 3D_MODE LEFT ENTER RIGHT EXIT DOWN BACK CHANNELUP CHANNELDOWN (and I think just numbers)


        WriteStringIniFile(MyUDN, "40", "1" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:1:;:-:" & "1", objRemoteFile)
        WriteStringIniFile(MyUDN, "41", "2" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:2:;:-:" & "2", objRemoteFile)
        WriteStringIniFile(MyUDN, "42", "3" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:3:;:-:" & "3", objRemoteFile)
        WriteStringIniFile(MyUDN, "43", "4" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:4:;:-:" & "4", objRemoteFile)
        WriteStringIniFile(MyUDN, "44", "5" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:5:;:-:" & "5", objRemoteFile)
        WriteStringIniFile(MyUDN, "45", "6" & ":;:-:" & "pointerbutton" & ":;:-:5:;:-:6:;:-:" & "6", objRemoteFile)
        WriteStringIniFile(MyUDN, "46", "7" & ":;:-:" & "pointerbutton" & ":;:-:6:;:-:1:;:-:" & "7", objRemoteFile)
        WriteStringIniFile(MyUDN, "47", "8" & ":;:-:" & "pointerbutton" & ":;:-:6:;:-:2:;:-:" & "8", objRemoteFile)
        WriteStringIniFile(MyUDN, "48", "9" & ":;:-:" & "pointerbutton" & ":;:-:6:;:-:3:;:-:" & "9", objRemoteFile)
        WriteStringIniFile(MyUDN, "49", "0" & ":;:-:" & "pointerbutton" & ":;:-:6:;:-:4:;:-:" & "0", objRemoteFile)

        WriteStringIniFile(MyUDN, "50", "RW" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:1:;:-:" & "REWIND", objRemoteFile)
        WriteStringIniFile(MyUDN, "51", "Prev" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:2:;:-:" & "GOTOPREV", objRemoteFile)
        WriteStringIniFile(MyUDN, "52", "Next" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:3:;:-:" & "GOTONEXT", objRemoteFile)
        WriteStringIniFile(MyUDN, "53", "FF" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:4:;:-:" & "FASTFORWARD", objRemoteFile)
        WriteStringIniFile(MyUDN, "54", "Stop" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:5:;:-:" & "STOP", objRemoteFile)
        WriteStringIniFile(MyUDN, "55", "Play" & ":;:-:" & "pointerbutton" & ":;:-:7:;:-:6:;:-:" & "PLAY", objRemoteFile)
        WriteStringIniFile(MyUDN, "56", "Pause" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:1:;:-:" & "PAUSE", objRemoteFile)
        WriteStringIniFile(MyUDN, "57", "Red" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:2:;:-:" & "RED", objRemoteFile)
        WriteStringIniFile(MyUDN, "58", "Green" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:3:;:-:" & "GREEN", objRemoteFile)
        WriteStringIniFile(MyUDN, "59", "Yellow" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:4:;:-:" & "YELLOW", objRemoteFile)
        WriteStringIniFile(MyUDN, "60", "Blue" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:5:;:-:" & "BLUE", objRemoteFile)
        WriteStringIniFile(MyUDN, "61", "Home" & ":;:-:" & "pointerbutton" & ":;:-:8:;:-:6:;:-:" & "HOME", objRemoteFile)
        WriteStringIniFile(MyUDN, "62", "Up" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:1:;:-:" & "UP", objRemoteFile)
        WriteStringIniFile(MyUDN, "63", "Down" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:2:;:-:" & "DOWN", objRemoteFile)
        WriteStringIniFile(MyUDN, "64", "Left" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:3:;:-:" & "LEFT", objRemoteFile)
        WriteStringIniFile(MyUDN, "65", "Right" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:4:;:-:" & "RIGHT", objRemoteFile)
        WriteStringIniFile(MyUDN, "66", "Exit" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:5:;:-:" & "EXIT", objRemoteFile)
        WriteStringIniFile(MyUDN, "67", "Enter" & ":;:-:" & "pointerbutton" & ":;:-:9:;:-:6:;:-:" & "ENTER", objRemoteFile)
        WriteStringIniFile(MyUDN, "68", "3D_Mode" & ":;:-:" & "pointerbutton" & ":;:-:10:;:-:1:;:-:" & "3D_MODE", objRemoteFile)
        WriteStringIniFile(MyUDN, "69", "CH_Up" & ":;:-:" & "pointerbutton" & ":;:-:10:;:-:2:;:-:" & "CHANNELUP", objRemoteFile)
        WriteStringIniFile(MyUDN, "70", "CH_Down" & ":;:-:" & "pointerbutton" & ":;:-:10:;:-:3:;:-:" & "CHANNELDOWN", objRemoteFile)
        WriteStringIniFile(MyUDN, "71", "Ok" & ":;:-:" & "pointerclick" & ":;:-:10:;:-:4:;:-:" & "", objRemoteFile)

        'WriteStringIniFile(MyUDN, "72", "Test1" & ":;:-:" & "test1" & ":;:-:10:;:-:5:;:-:" & "", objRemoteFile)
        'WriteStringIniFile(MyUDN, "73", "Test2" & ":;:-:" & "test2" & ":;:-:10:;:-:6:;:-:" & "", objRemoteFile)
        'WriteStringIniFile(MyUDN, "74", "Test3" & ":;:-:" & "test3" & ":;:-:10:;:-:7:;:-:" & "", objRemoteFile)


    End Sub

    Private Sub CreateHSLGRemoteButtons(ReCreate As Boolean)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSLGRemoteButtons called for UPnPDevice = " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
        HSRefRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Remote.ToString & "HSCode", -1)
        If HSRefRemote = -1 Then
            HSRefRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Remote.ToString)
        Else
            If Not ReCreate Then Exit Sub ' already exist
        End If

        If HSRefRemote = -1 Then Exit Sub
        If ReCreate Then
            hs.DeviceVSP_ClearAll(HSRefRemote, True)
            hs.DeviceVGP_ClearAll(HSRefRemote, True)
        End If

        Dim Pair As VSPair
        Dim Column As Integer = 1

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRemoteOn
        Pair.Status = "Remote On"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = Column
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRemoteOff
        Pair.Status = "Remote Off"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = Column + 1
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psCreateRemoteButtons
        Pair.Status = "Create Remote Buttons"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = Column + 2
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psCreateRemoteAppButtons
        Pair.Status = "Retrieve App, Channel & Input Info"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = Column + 3
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)

        Dim GraphicsPair As New VGPair()

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsDeactivated
        Pair.Status = "Deactivated"
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "NOKBtn.png"
        GraphicsPair.Set_Value = dsDeactivated
        hs.DeviceVGP_AddPair(HSRefRemote, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOffLine
        Pair.Status = "Activated Offline"
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "PartialOKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOffLine
        hs.DeviceVGP_AddPair(HSRefRemote, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivateOnLine
        Pair.Status = "Activated Online"
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivateOnLine
        hs.DeviceVGP_AddPair(HSRefRemote, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsActivatedOnLineUnregistered
        Pair.Status = "Activated Online Unregistered"
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "PartialOKBtn.png"
        GraphicsPair.Set_Value = dsActivatedOnLineUnregistered
        hs.DeviceVGP_AddPair(HSRefRemote, GraphicsPair)

    End Sub

    Private Sub CreateHSLGRemoteServices()

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSLGRemoteServices called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        ' First Create the Status Device
        HSRefServiceRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Status.ToString & "HSCode", -1)
        If HSRefServiceRemote = -1 Then
            HSRefServiceRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Status.ToString)
        End If
        If HSRefServiceRemote = -1 Then Exit Sub
        'CreateOnOffTogglePairs((HSRefMute)

        ' First Create the Volume Device
        HSRefLGVolume = GetIntegerIniFile(MyUDN, "di" & HSDevices.Volume.ToString & "HSCode", -1)
        If HSRefLGVolume = -1 Then
            HSRefLGVolume = CreateHSServiceDevice(HSRefRemote, HSDevices.Volume.ToString)
        End If
        If HSRefLGVolume = -1 Then Exit Sub
        HSRefVolume = HSRefLGVolume
        CreateLGVolumePairs(HSRefLGVolume)

        ' First Create the Mute Device
        HSRefMute = GetIntegerIniFile(MyUDN, "di" & HSDevices.Mute.ToString & "HSCode", -1)
        If HSRefMute = -1 Then
            HSRefMute = CreateHSServiceDevice(HSRefRemote, HSDevices.Mute.ToString)
        End If
        If HSRefMute = -1 Then Exit Sub
        CreateLGOnOffTogglePairs(HSRefMute)

    End Sub

    Private Sub CreateLGRemoteServiceButtons(HSRef As Integer)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGRemoteServiceButtons called for device - " & MyUPnPDeviceName & " and HSRef = " & HSRef.ToString, LogType.LOG_TYPE_INFO)

        Dim objRemoteFile As String = gRemoteControlPath
        Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
        Dim AppOffset As Integer = (GetIntegerIniFile("LG Offset Info" & MyUDN, "AddAppButtons", 20, objRemoteFile) / MaxButtonCollumns) + 1
        Dim InputOffset As Integer = (GetIntegerIniFile("LG Offset Info" & MyUDN, "AddInputButtons", 40, objRemoteFile) / MaxButtonCollumns) + 1 + AppOffset
        ' Dim ChannelOffset As Integer = (GetIntegerIniFile("LG Offset Info" & myudn, "AddChannelButtons", 60, objRemoteFile) / MaxButtonCollumns) + 1 + InputOffset

        Try
            RemoteButtons = GetIniSection(MyUDN, objRemoteFile) '  As Dictionary(Of String, String)
            If RemoteButtons Is Nothing Then
                If PIDebuglevel > DebugLevel.dlOff Then Log("Error in CreateLGRemoteServiceButtons for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
                Exit Try
            Else
                Dim RemoteButtonString As String = ""
                Dim RemoteButtonName As String = ""
                Dim RemoteButtonValue As Integer = 0
                For Each RemoteButton In RemoteButtons
                    If RemoteButton.Key <> "" Then
                        RemoteButtonString = RemoteButton.Value
                        Dim RemoteButtonInfos As String()
                        RemoteButtonInfos = Split(RemoteButtonString, ":;:-:")
                        ' Active ; Given Name ; Key Code ; Row Index ; Column Index -- RowIndex and Column Index start with 1
                        If UBound(RemoteButtonInfos, 1) > 2 Then
                            If RemoteButtonInfos(1).IndexOf("LGlaunch") = 0 Then
                                RemoteButtonName = RemoteButtonInfos(0)
                                RemoteButtonValue = Val(RemoteButton.Key)
                                Dim Pair As VSPair
                                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                                Pair.PairType = VSVGPairType.SingleValue
                                Pair.Value = RemoteButtonValue
                                Pair.Status = RemoteButtonName
                                Pair.Render = Enums.CAPIControlType.Button
                                Pair.Render_Location.Row = 1 + Val(RemoteButtonInfos(2))
                                Pair.Render_Location.Column = Val(RemoteButtonInfos(3))
                                If File.Exists(CurrentAppPath & "/html" & ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png") Then ' dcor tralala tobe fixed needs to be remoted!
                                    ' maybe I should use the HS.WriteHTMLImage(Image As Image, Dest As String, OverWrite As Boolean) As Boolean as a workaround; set overwrite to false
                                    Dim vg As New VGPair
                                    vg.PairType = VSVGPairType.SingleValue
                                    vg.Graphic = ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png"
                                    vg.Set_Value = RemoteButtonValue
                                    hs.DeviceVGP_AddPair(HSRef, vg)
                                    Pair.PairButtonImageType = Enums.CAPIControlButtonImage.Use_Custom
                                    Pair.PairButtonImage = ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png"
                                End If
                                hs.DeviceVSP_AddPair(HSRef, Pair)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGRemoteServiceButtons added launch for device - " & MyUPnPDeviceName & " with name = " & RemoteButtonName, LogType.LOG_TYPE_INFO)
                            ElseIf RemoteButtonInfos(1).IndexOf("LGsetinput") = 0 Then
                                RemoteButtonName = RemoteButtonInfos(0)
                                RemoteButtonValue = Val(RemoteButton.Key)
                                Dim Pair As VSPair
                                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                                Pair.PairType = VSVGPairType.SingleValue
                                Pair.Value = RemoteButtonValue
                                Pair.Status = RemoteButtonName
                                Pair.Render = Enums.CAPIControlType.Button
                                Pair.Render_Location.Row = 1 + Val(RemoteButtonInfos(2)) + AppOffset
                                Pair.Render_Location.Column = Val(RemoteButtonInfos(3))
                                If File.Exists(CurrentAppPath & "/html" & ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png") Then ' dcor tralala tobe fixed needs to be remoted!
                                    Dim vg As New VGPair
                                    vg.PairType = VSVGPairType.SingleValue
                                    vg.Graphic = ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png"
                                    vg.Set_Value = RemoteButtonValue
                                    hs.DeviceVGP_AddPair(HSRef, vg)
                                    Pair.PairButtonImageType = Enums.CAPIControlButtonImage.Use_Custom
                                    Pair.PairButtonImage = ImagesPath & "Artwork/" & RemoteButtonInfos(6) & ".png"
                                End If
                                hs.DeviceVSP_AddPair(HSRef, Pair)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGRemoteServiceButtons added input for device - " & MyUPnPDeviceName & " with name = " & RemoteButtonName, LogType.LOG_TYPE_INFO)
                            ElseIf RemoteButtonInfos(1).IndexOf("LGsetchannel") = 0 Then
                                RemoteButtonName = RemoteButtonInfos(0)
                                RemoteButtonValue = Val(RemoteButton.Key)
                                Dim Pair As VSPair
                                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                                Pair.PairType = VSVGPairType.SingleValue
                                Pair.Value = RemoteButtonValue
                                Pair.Status = RemoteButtonName
                                Pair.Render = Enums.CAPIControlType.Button
                                Pair.Render_Location.Row = 1 + Val(RemoteButtonInfos(2)) + InputOffset
                                Pair.Render_Location.Column = Val(RemoteButtonInfos(3))
                                If File.Exists(CurrentAppPath & "/html" & ImagesPath & "Artwork/" & RemoteButtonInfos(5) & ".png") Then ' dcor tralala tobe fixed needs to be remoted!
                                    Dim vg As New VGPair
                                    vg.PairType = VSVGPairType.SingleValue
                                    vg.Graphic = ImagesPath & "Artwork/" & RemoteButtonInfos(5) & ".png"
                                    vg.Set_Value = RemoteButtonValue
                                    hs.DeviceVGP_AddPair(HSRef, vg)
                                    Pair.PairButtonImageType = Enums.CAPIControlButtonImage.Use_Custom
                                    Pair.PairButtonImage = ImagesPath & "Artwork/" & RemoteButtonInfos(5) & ".png"
                                End If
                                hs.DeviceVSP_AddPair(HSRef, Pair)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGRemoteServiceButtons added channel for device - " & MyUPnPDeviceName & " with name = " & RemoteButtonName, LogType.LOG_TYPE_INFO)
                            End If
                        End If
                    End If
                Next
                RemoteButtons = Nothing
            End If
        Catch ex As Exception
            Log("Error in CreateLGRemoteServiceButtons parsing the .ini file for remote button info for device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub CreateLGOnOffTogglePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpLGOff
        Pair.Status = "Off"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpLGOn
        Pair.Status = "On"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpLGToggle
        Pair.Status = "Toggle"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)


        Dim GraphicsPair As New VGPair()

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = msUnmuted
        Pair.Status = "Mute Off"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "UnMuted.png"
        GraphicsPair.Set_Value = msUnmuted
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = msMuted
        Pair.Status = "Mute On"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "Muted.png"
        GraphicsPair.Set_Value = msMuted
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

    End Sub

    Private Sub CreateLGVolumePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        ' add a Down button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpLGDown
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
        Pair.Value = vpLGUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)


        Dim GraphicsPair As New VGPair()

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = 1000
        Pair.Status = "Disabled"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "Muted.png"
        GraphicsPair.Set_Value = 1000
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)


    End Sub


    Private Sub CreateLGAppandInputButtons(HSRefRemote As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateLGAppandInputButtons called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        LGSendCommand("listLaunchPoints_0", "request", "ssap://com.webos.applicationManager/listLaunchPoints", "", "", True)
        LGSendCommand("getexternalinputlist_0", "request", "ssap://tv/getExternalInputList", "", "", True)
        LGSendCommand("channellist_0", "request", "ssap://tv/getChannelList", "", "", True)
    End Sub

    Private Sub LGAddAppButtons(Payload As Object)
        ' itterate through LaunchPoints
        '"type""response",
        '"id":"launcher_1",
        '"payload":{
        '   "subscribed"false,
        '   "launchPoints":[
        '       {
        '       "systemApp":true,
        '       "removable":false,
        '       "relaunch":false,
        '       "largeIcon":"",
        '       "bgImages":[
        '       ],
        '       "userData":"",
        '       "id":"com.webos.app.browser",
        '       "title":"Web Browser",
        '       "bgColor":"",
        '       "iconColor":"#009999",
        '       "appDescription":"",
        '       "lptype":"default",
        '       "params":{
        '       },
        '       "bgImage":"",
        '       "unmovable":false,
        '       "icon":"http://192.168.1.205:3000/resources/44c7d15a7645f85432188a5229a591435d06b9b5/webbrowser_icon.png",
        '       "launchPointId":"com.webos.app.browser_default",
        '       "favicon":"",
        '       "imageForRecents":"",
        '        "tileSize":"normal"
        '       },
        '        ......
        '   }
        '],
        '   "caseDetail":{
        '       "change"{
        '        }
        '   },
        '   "returnValue":true
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGAddAppButtons called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim SeqNbr As Integer = 0
            Dim ButtonIndex As Integer = 100
            Dim RowIndex As Integer = 1
            Dim ColumnIndex As Integer = 1
            Dim objRemoteFile As String = gRemoteControlPath
            Dim LaunchPoints As Object = FindPairInJSONString(Payload, "launchPoints")
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(LaunchPoints)
            WriteStringIniFile("LG Offset Info" & MyUDN, "AddAppButtons", JSONdataLevel1.length, objRemoteFile)

            For Each LaunchPoint As Object In JSONdataLevel1
                Dim id As String = ""
                Dim Title As String = ""
                Dim LaunchpointID As String = ""
                Dim Icon As String = ""
                Dim AppImage As Image = Nothing
                'Dim JSONDataLevel2 As Object In LaunchPoint
                For Each Entry As Object In LaunchPoint
                    If Entry.Key.ToString = "id" Then
                        id = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "title" Then
                        Title = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "launchPointId" Then
                        LaunchpointID = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "icon" Then
                        Icon = Entry.value.ToString
                    End If
                Next
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGAddAppButtons for Device = " & MyUPnPDeviceName & " found app = " & Title & " with id=" & id & ", LaunchpointID=" & LaunchpointID & " and ICON=" & Icon, LogType.LOG_TYPE_INFO)
                WriteStringIniFile(MyUDN, ButtonIndex.ToString, Title & ":;:-:" & "LGlaunch" & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString & ":;:-:" & id & ":;:-:" & LaunchpointID & ":;:-:" & "LGAppImage_" & MyUDN & "_" & SeqNbr.ToString.ToString, objRemoteFile)
                If Icon <> "" And id <> "" Then
                    Try
                        AppImage = GetPicture(Icon.ToString)
                        If Not AppImage Is Nothing Then
                            Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                            Dim SuccesfullSave As Boolean = False
                            SuccesfullSave = hs.WriteHTMLImage(AppImage, FileArtWorkPath & "LGAppImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", True)
                            If Not SuccesfullSave Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in LGAddAppButtons for Device = " & MyUPnPDeviceName & " had error storing Image at " & FileArtWorkPath & "LGAppImage_" & MyUDN & "_" & SeqNbr.ToString.ToString & ".png", LogType.LOG_TYPE_ERROR)
                            Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGAddAppButtons for Device = " & MyUPnPDeviceName & " stored LG App Image at " & FileArtWorkPath & "LGAppImage_" & MyUDN & "_" & SeqNbr.ToString.ToString & ".png", LogType.LOG_TYPE_INFO)
                            End If
                            AppImage.Dispose()
                            SeqNbr += 1
                        End If
                    Catch ex As Exception
                        Log("Error in LGAddAppButtons retrieving the LG App Image with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
                ButtonIndex = ButtonIndex + 1
                ColumnIndex = ColumnIndex + 1
                If ColumnIndex > MaxButtonCollumns Then
                    ColumnIndex = 1
                    RowIndex += 1
                End If
            Next
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in LGAddAppButtons for Device = " & MyUPnPDeviceName & " while processing response with error = " & ex.Message & " with Payload = " & Payload.ToString, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub AddInputButtons(Payload As Object)

        '    "type""response",
        '    "id":"channels_2",
        '    "payload":{
        '        "returnValue":true,
        '        "valueList":"",
        '        "dataSource":0,
        '        "dataType":1,
        '        "cableAnalogSkipped":false,
        '        "scannedChannelCount":{
        '            "terrestrialAnalogCount":2,
        '            "terrestrialDigitalCount":0,
        '            "cableAnalogCount":0,
        '            "cableDigitalCount":0,
        '            "satelliteDigitalCount":0
        '        },
        '        "deviceSourceIndex":1,
        '        "channelListCount":2,
        '        "channelLogoServerUrl":"",
        '        "ipChanInteractiveUrl":"",
        '        "channelList":[
        '            {
        '                "channelId":"0_2_2_0_0_0_0",
        '                "programId":"0_2_2_0_0_0_0",
        '                "signalChannelId":"0_2_2_0_0_0_0",
        '                "chanCode":"UNKNOWN",
        '                "channelMode":"Terrestrial",
        '                "channelModeId":0,
        '                "channelType":"Terrestrial Analog TV",
        '                "channelTypeId":0,
        '                "channelNumber":"2-0",
        '                "majorNumber":2,
        '                "minorNumber":0,
        '                "channelName":"",
        '                "skipped":false,
        '                "locked":false,
        '                "descrambled":true,
        '                "scrambled":false,
        '                "serviceType":0,
        '                "favoriteGroup":[],
        '                "imgUrl":"",
        '                "display":1,
        '                "satelliteName":" ",
        '                "fineTuned":false,
        '                "Frequency":0,
        '                "shortCut":0,
        '                "Bandwidth":0,
        '                "HDTV":false,
        '                "Invisible":false,
        '                "TV":true,
        '                "DTV":false,
        '                "ATV":true,
        '                "Data":false,
        '                "Radio":false,
        '                "Numeric":false,
        '                "PrimaryCh":true,
        '                "specialService":false,
        '                "CASystemIDList":{
        '                },
        '                "CASystemIDListCount":0,
        '                "groupIdList":[0],
        '                "channelGenreCode":"UNKNOWN",
        '                "favoriteIdxA":250,
        '                "favoriteIdxB":250,
        '                "favoriteIdxC":250,
        '                "favoriteIdxD":250,
        '                "imgUrl2":"",
        '                "channelLogoSize":"UNKNOWN",
        '                "ipChanServerUrl":"",
        '                "payChan":false,
        '                "IPChannelCode":"UNKNOWN",
        '                "ipCallNumber":"UNKNOWN",
        '                "otuFlag":false,
        '                "favoriteIdxE":250,
        '                "favoriteIdxF":250,
        '                "favoriteIdxG":250,
        '                "favoriteIdxH":250,
        '                "satelliteLcn":false,
        '                "waterMarkUrl":"",
        '                "channelNameSortKey":"",
        '                "ipChanType":"UNKNOWN",
        '                "adultFlag":0,
        '                "ipChanCategory":"UNKNOWN",
        '                "ipChanInteractive":false,
        '                "callSign":"UNKNOWN",
        '                "adFlag":0,
        '                "configured":false,
        '                "lastUpdated":"",
        '                "ipChanCpId":"UNKNOWN",
        '                "isFreeviewPlay":1,
        '                "playerService":"com.webos.service.tv",
        '                "TSID":0,
        '                "SVCID":0
        '            },
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddInputButtons called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Try
            Dim SeqNbr As Integer = 0
            Dim ButtonIndex As Integer = 300
            Dim RowIndex As Integer = 1
            Dim ColumnIndex As Integer = 1
            Dim objRemoteFile As String = gRemoteControlPath
            Dim Devices As Object = FindPairInJSONString(Payload, "devices")
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(Devices)
            WriteStringIniFile("LG Offset Info" & MyUDN, "AddInputButtons", JSONdataLevel1.length, objRemoteFile)

            For Each Device As Object In JSONdataLevel1
                Dim id As String = ""
                Dim Label As String = ""
                Dim AppID As String = ""
                Dim Icon As String = ""
                Dim Port As String = ""
                Dim AppImage As Image = Nothing
                For Each Entry As Object In Device
                    If Entry.Key.ToString = "id" Then
                        id = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "label" Then
                        Label = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "appId" Then
                        AppID = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "icon" Then
                        Icon = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "port" Then
                        Port = Entry.value.ToString
                    End If
                Next
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddInputButtons for Device = " & MyUPnPDeviceName & " found device = " & Label & " with id=" & id & ", AppID=" & AppID & ", Port=" & Port & " and ICON=" & Icon, LogType.LOG_TYPE_INFO)
                WriteStringIniFile(MyUDN, ButtonIndex.ToString, Label & ":;:-:" & "LGsetinput" & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString & ":;:-:" & id & ":;:-:" & Port & ":;:-:" & "LGDeviceImage_" & MyUDN & "_" & SeqNbr.ToString.ToString, objRemoteFile)
                If Icon <> "" And id <> "" Then
                    Try
                        AppImage = GetPicture(Icon.ToString)
                        If Not AppImage Is Nothing Then
                            Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                            Dim SuccesfullSave As Boolean = False
                            SuccesfullSave = hs.WriteHTMLImage(AppImage, FileArtWorkPath & "LGDeviceImage_" & SeqNbr.ToString & ".png", True)
                            If Not SuccesfullSave Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddInputButtons for Device = " & MyUPnPDeviceName & " had error storing Image at " & FileArtWorkPath & "LGDeviceImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", LogType.LOG_TYPE_ERROR)
                            Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddInputButtons for Device = " & MyUPnPDeviceName & " stored LG InputDevice Image at " & FileArtWorkPath & "LGDeviceImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", LogType.LOG_TYPE_INFO)
                            End If
                            AppImage.Dispose()
                            SeqNbr += 1
                        End If
                    Catch ex As Exception
                        Log("Error in AddInputButtons retrieving the LG InputDevice Image with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
                ButtonIndex = ButtonIndex + 1
                ColumnIndex = ColumnIndex + 1
                If ColumnIndex > MaxButtonCollumns Then
                    ColumnIndex = 1
                    RowIndex += 1
                End If
            Next
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddInputButtons for Device = " & MyUPnPDeviceName & " while processing response with error = " & ex.Message & " with Payload = " & Payload.ToString, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub AddChannelButtons(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddChannelButtons called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Try
            Dim SeqNbr As Integer = 0
            Dim ButtonIndex As Integer = 400
            Dim RowIndex As Integer = 1
            Dim ColumnIndex As Integer = 1
            Dim objRemoteFile As String = gRemoteControlPath
            Dim Devices As Object = FindPairInJSONString(Payload, "channelList")
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(Devices)
            WriteStringIniFile("LG Offset Info" & MyUDN, "AddChannelButtons", JSONdataLevel1.length, objRemoteFile)

            For Each Channel As Object In JSONdataLevel1
                Dim channelId As String = ""
                Dim channelNumber As String = ""
                Dim imgUrl As String = ""
                Dim AppImage As Image = Nothing
                For Each Entry As Object In Channel
                    If Entry.Key.ToString = "channelId" Then
                        channelId = Entry.value.ToString
                    ElseIf Entry.Key = "channelNumber" Then
                        channelNumber = Entry.value.ToString
                    ElseIf Entry.Key.ToString = "imgUrl" Then
                        imgUrl = Entry.value.ToString
                    End If
                Next
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddChannelButtons for Device = " & MyUPnPDeviceName & " found ChannelNumber = " & channelNumber & " with id=" & channelId & " and imgUrl=" & imgUrl, LogType.LOG_TYPE_INFO)
                WriteStringIniFile(MyUDN, ButtonIndex.ToString, channelNumber & ":;:-:" & "LGsetchannel" & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString & ":;:-:" & channelId & ":;:-:" & "LGChannelImage_" & MyUDN & "_" & SeqNbr.ToString, objRemoteFile)
                If imgUrl <> "" And channelId <> "" Then
                    Try
                        AppImage = GetPicture(imgUrl.ToString)
                        If Not AppImage Is Nothing Then
                            Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                            Dim SuccesfullSave As Boolean = False
                            SuccesfullSave = hs.WriteHTMLImage(AppImage, FileArtWorkPath & "LGChannelImage_" & SeqNbr.ToString & ".png", True)
                            If Not SuccesfullSave Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddChannelButtons for Device = " & MyUPnPDeviceName & " had error storing Image at " & FileArtWorkPath & "LGChannelImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", LogType.LOG_TYPE_ERROR)
                            Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddChannelButtons for Device = " & MyUPnPDeviceName & " stored LG channel Image at " & FileArtWorkPath & "LGChannelImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", LogType.LOG_TYPE_INFO)
                            End If
                            AppImage.Dispose()
                            SeqNbr += 1
                        End If
                    Catch ex As Exception
                        Log("Error in AddChannelButtons retrieving the LG Channel Image with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
                ButtonIndex = ButtonIndex + 1
                ColumnIndex = ColumnIndex + 1
                If ColumnIndex > MaxButtonCollumns Then
                    ColumnIndex = 1
                    RowIndex += 1
                End If
            Next
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddChannelButtons for Device = " & MyUPnPDeviceName & " while processing response with error = " & ex.Message & " with Payload = " & Payload.ToString, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub TreatRegistered(inBytes As Byte())
        ' great, we registered successfully, now retrieve the client-key in the payload
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatRegistered called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Dim Payload As Object = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "payload")
        If Payload Is Nothing Then Exit Sub
        Dim ClientKey As String = FindPairInJSONString(Payload, "client-key").ToString.Trim("""")
        If ClientKey <> "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGDataReceived received a ClientKey for Device = " & MyUPnPDeviceName & " with ClientKey = " & ClientKey, LogType.LOG_TYPE_INFO)
            ' save it!
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diLGClientKey.ToString, ClientKey)
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)
            MyRemoteServiceActive = True
            SetHSRemoteState()
            ' check some capabilities
            LGSendCommand("services_0", "request", "ssap://api/getServiceList", "", "", True)
            ' get the websocket info to do rest of "old style" remote
            LGSendCommand("getpointerinputsocket_0", "request", "ssap://com.webos.service.networkinput/getPointerInputSocket", """{}""", "", True)
            ' Subscribe to interesting events
            'LGSendCommand("channels_0", "subscribe", "ssap://tv/getChannelList", "", "", True)
            LGSendCommand("externalinput_0", "subscribe", "ssap://tv/getExternalInputList", "", "", True)
            LGSendCommand("audiostatus_0", "subscribe", "ssap://audio/getStatus", "", "", True)
            LGSendCommand("volumestatus_0", "subscribe", "ssap://audio/getVolume", "", "", True)
            ' additional stuff
            LGSendCommand("keyboardstatus_0", "subscribe", "ssap://com.webos.service.ime/registerRemoteKeyboard", "", "", True)
            ' {"type":"response","id":"keyboardstatus_0","payload":{"currentWidget":{"focus":false},"subscribed":true}}
            'LGSendCommand("interestingevents_0", "subscribe", "ssap://com.webos.service.tv.keymanager/listInterestingEvents", "", "", True)
            '{"type":"error","id":"interestingevents_0","error":"401 insufficient permissions","payload":{}} 
            LGSendCommand("foregroundapp_0", "subscribe", "ssap://com.webos.applicationManager/getForegroundAppInfo", "", "", True)
            ' {"type":"error","id":"foregroundapp_010","error":"400 unknown message type","payload":{}} 
            ' nok LGSendCommand("subscribeappstatus_0", "subscribe", "ssap://com.webos.service.appstatus/getAppStatus", "", "", True)
            'LGSendCommand("subscribeappstatus_0", "subscribe", "ssap://system.launcher/getAppState", "", "", True) ' return denied status
            'LGSendCommand("subscribeappstatus_2", "request", "ssap://system.launcher/getAppState", "", "", True)
            ' nok LGSendCommand("subscribeappstatus_3", "request", "ssap://system.launcher/getAppStatus", "", "", True)
            'LGSendCommand("currentchannel_0", "subscribe", "ssap://tv/getCurrentChannel", "", "", True)
        End If
    End Sub

    Private Sub TreatPointerInputSocket(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatPointerInputSocket called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        '{"type":"response","id":"1","payload":{"socketPath":"ws://192.168.1.69:3000/resources/9b7d24a4862c409499865f0acfc2edd12aaee2db/netinput.pointer.sock","returnValue":true}}
        ' or over SSL
        ' {"type":"response","id":"getpointerinputsocket_0","payload":{"socketPath":"wss://192.168.1.183:3001/resources/d33c9903271be293c3576fbcbe90cc195184d720/netinput.pointer.sock","returnValue":true}} 
        Try
            Dim SocketPath As String = FindPairInJSONString(Payload, "socketPath").ToString
            If SocketPath = "" Then Exit Sub
            Dim SecWebSocketKey As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "")
            If SecWebSocketKey = "" Then Exit Sub

            Dim URL = New Uri(SocketPath)

            If LGWebPointerSocket Is Nothing Then
                Try
                    LGWebPointerSocket = New WebSocketClient(SocketPath.IndexOf("wss://") <> -1) ' no SSL
                Catch ex As Exception
                    Log("Error in TreatPointerInputSocket for UPnPDevice = " & MyUPnPDeviceName & " unable to open WebPointerSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
            End If

            AddHandler LGWebPointerSocket.DataReceived, AddressOf HandleLGPointerDataReceived
            AddHandler LGWebPointerSocket.WebSocketClosed, AddressOf HandleLGPointerSocketClosed

            If Not LGWebPointerSocket.ConnectSocket(URL.Host, URL.Port) Then
                Exit Sub
            End If

            ' wait until connected. Important for SSL as it takes longer
            Dim WaitLoopCounter As Integer = 0
            While LGWebPointerSocket.MySocketIsClosed
                wait(1)
                WaitLoopCounter += 1
                If WaitLoopCounter > 10 Then Exit While
            End While

            LGWebPointerSocket.Receive()

            Dim WebURL = URL.PathAndQuery
            If WebURL.IndexOf("/") = 0 Then WebURL = WebURL.Remove(0, 1)

            If Not LGWebPointerSocket.UpgradeWebSocket(WebURL, SecWebSocketKey, 0, True) Then ' timevalue set to 0 no pings
                LGWebPointerSocket.CloseSocket()
                Exit Sub
            End If

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatPointerInputSocket called for Device = " & MyUPnPDeviceName & " and Payload = " & Payload.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TreatHelloReceived(Payload As Byte())
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatHelloReceived called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"protocolVersion":1,"deviceType":"tv","deviceOS":"webOS","deviceOSVersion":"4.1.0","deviceOSReleaseVersion":"3.0.0","deviceUUID":"05ee840a-cf71-15a1-12d8-11e0eb21437c","pairingTypes":["PIN","PROMPT","COMBINED"]}}
        Try
            Dim PairingTypes As Object = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(Payload), "pairingTypes")
            If PairingTypes Is Nothing Then Exit Sub
            If PairingTypes.ToString.IndexOf("PROMPT") = -1 Then
                'Exit Sub ' we need to use PIN or COMBINED
            End If

            Dim PairingTypeArray As String() = Split(PairingTypes.ToString, ",")

            Dim ClientKey As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diLGClientKey.ToString, "")

            'Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS""],""manifestVersion"":1}}}"
            'Dim JSONAlreadyRegisteredString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""client-key"":""" & ClientKey & """,""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS""],""manifestVersion"":1}}}"

            Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""TEST_OPEN"",""TEST_PROTECTED"",""TEST_SECURE"",""READ_APP_STATUS"",""READ_POWER_STATE"",""READ_COUNTRY_INFO"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME"",""CLOSE""],""manifestVersion"":1}}}"
            Dim JSONAlreadyRegisteredString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""client-key"":""" & ClientKey & """,""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""TEST_OPEN"",""TEST_PROTECTED"",""TEST_SECURE"",""READ_APP_STATUS"",""READ_POWER_STATE"",""READ_COUNTRY_INFO"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME"",""CLOSE""],""manifestVersion"":1}}}"

            If ClientKey <> "" Then
                JSONRegisterString = JSONAlreadyRegisteredString
            End If

            Dim SocketData As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(JSONRegisterString)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatHelloReceived send Registration String for device - " & MyUPnPDeviceName & " with String = " & JSONRegisterString.ToString, LogType.LOG_TYPE_INFO)

            If Not LGWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatHelloReceived for device - " & MyUPnPDeviceName & " unsuccessful sending registration_0", LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
            ' expected response inHandleLGDataREceived
            ' {"type":"response","id":"register_0","payload":{"pairingType":"PROMPT","returnValue":true}}
            ' {"type": "registered","id":"register_0","payload":{"client-key":"5e738846a5e1d5ca08df28c4d955e8b8"}}

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatHelloReceived called for Device = " & MyUPnPDeviceName & " and Payload = " & ASCIIEncoding.ASCII.GetChars(Payload) & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub TreatSetIOExLG(ButtonValue As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExLG called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            Case psRemoteOff  ' remote off
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, False)
                Try
                    LGCloseWebSocket()
                Catch ex As Exception
                    Log("Error in TreatSetIOExLG for UPnPDevice = " & MyUPnPDeviceName & "  setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                SetAdministrativeStateRemote(False)
            Case psRemoteOn  ' remote on
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                If LGActivateRemote() Then
                    SetAdministrativeStateRemote(True)
                End If
            Case psCreateRemoteButtons
                CreateLGRemoteIniFileInfo()
                CreateRemoteButtons(HSRefRemote)
                CreateLGRemoteServiceButtons(HSRefServiceRemote)
            Case psWOL
                SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), PlugInIPAddress)
                SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, ""), PlugInIPAddress)
                Dim deviceIpAddress As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diIPAddress.ToString, "")
                If deviceIpAddress <> "" Then
                    SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), deviceIpAddress)
                    SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, ""), deviceIpAddress)
                End If
            Case psCreateRemoteAppButtons
                CreateLGAppandInputButtons(HSRefRemote)
            Case Else
                If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) And UCase(DeviceStatus) = "ONLINE" Then
                    Dim objRemoteFile As String = gRemoteControlPath
                    Dim ButtonInfoString As String = GetStringIniFile(MyUDN, ButtonValue.ToString, "", objRemoteFile)
                    Dim ButtonInfos As String()
                    ButtonInfos = Split(ButtonInfoString, ":;:-:")
                    If UBound(ButtonInfos, 1) > 3 Then
                        LGSendKeyCode(ButtonInfos(1), ButtonInfos(4))
                    End If
                End If
        End Select

    End Sub

    Private Sub TreatServicesInfo(payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatServicesInfo called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"services_0","payload":{"returnValue":true,"services":[{"name":"api","version":1},{"name":"audio","version":1},{"name":"config","version":1},{"name":"media.controls","version":1},{"name":"media.viewer","version":1},{"name":"pairing","version":1},{"name":"settings","version":1},{"name":"system","version":1},{"name":"system.launcher","version":1},{"name":"system.notifications","version":1},{"name":"timer","version":1},{"name":"tv","version":1},{"name":"user","version":1},{"name":"webapp","version":2}]}} 
    End Sub

    Private Sub TreatChannelInfoEvent(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatChannelInfoEvent called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"channels_0","payload":{"returnValue":true,"valueList":"","dataSource":0,"dataType":1,"cableAnalogSkipped":false,"scannedChannelCount":{"terrestrialAnalogCount":2,"terrestrialDigitalCount":0,"cableAnalogCount":0,"cableDigitalCount":0,"satelliteDigitalCount":0},"deviceSourceIndex":1,"channelListCount":2,"channelLogoServerUrl":"","ipChanInteractiveUrl":"","channelList":[{"channelId":"0_2_2_0_0_0_0","programId":"0_2_2_0_0_0_0","signalChannelId":"0_2_2_0_0_0_0","chanCode":"UNKNOWN","channelMode":"Terrestrial","channelModeId":0,"channelType":"Terrestrial Analog TV","channelTypeId":0,"channelNumber":"2-0","majorNumber":2,"minorNumber":0,"channelName":"","skipped":false,"locked":false,"descrambled":true,"scrambled":false,"serviceType":0,"favoriteGroup":[],"imgUrl":"","display":1,"satelliteName":" ","fineTuned":false,"Frequency":0,"shortCut":0,"Bandwidth":0,"HDTV":false,"Invisible":false,"TV":true,"DTV":false,"ATV":true,"Data":false,"Radio":false,"Numeric":false,"PrimaryCh":true,"specialService":false,"CASystemIDList":{},"CASystemIDListCount":0,"groupIdList":[0],"channelGenreCode":"UNKNOWN","favoriteIdxA":250,"favoriteIdxB":250,"favoriteIdxC":250,"favoriteIdxD":250,"imgUrl2":"","channelLogoSize":"UNKNOWN","ipChanServerUrl":"","payChan":false,"IPChannelCode":"UNKNOWN","ipCallNumber":"UNKNOWN","otuFlag":false,"favoriteIdxE":250,"favoriteIdxF":250,"favoriteIdxG":250,"favoriteIdxH":250,"satelliteLcn":false,"waterMarkUrl":"","channelNameSortKey":"","ipChanType":"UNKNOWN","adultFlag":0,"ipChanCategory":"UNKNOWN","ipChanInteractive":false,"callSign":"UNKNOWN","adFlag":0,"configured":false,"lastUpdated":"","ipChanCpId":"UNKNOWN","isFreeviewPlay":1,"playerService":"com.webos.service.tv","TSID":0,"SVCID":0},{"channelId":"0_3_3_0_0_0_0","programId":"0_3_3_0_0_0_0","signalChannelId":"0_3_3_0_0_0_0","chanCode":"UNKNOWN","channelMode":"Terrestrial","channelModeId":0,"channelType":"Terrestrial Analog TV","channelTypeId":0,"channelNumber":"3-0","majorNumber":3,"minorNumber":0,"channelName":"","skipped":false,"locked":false,"descrambled":true,"scrambled":false,"serviceType":0,"favoriteGroup":[],"imgUrl":"","display":1,"satelliteName":" ","fineTuned":false,"Frequency":0,"shortCut":0,"Bandwidth":0,"HDTV":false,"Invisible":false,"TV":true,"DTV":false,"ATV":true,"Data":false,"Radio":false,"Numeric":false,"PrimaryCh":true,"specialService":false,"CASystemIDList":{},"CASystemIDListCount":0,"groupIdList":[0],"channelGenreCode":"UNKNOWN","favoriteIdxA":250,"favoriteIdxB":250,"favoriteIdxC":250,"favoriteIdxD":250,"imgUrl2":"","channelLogoSize":"UNKNOWN","ipChanServerUrl":"","payChan":false,"IPChannelCode":"UNKNOWN","ipCallNumber":"UNKNOWN","otuFlag":false,"favoriteIdxE":250,"favoriteIdxF":250,"favoriteIdxG":250,"favoriteIdxH":250,"satelliteLcn":false,"waterMarkUrl":"","channelNameSortKey":"","ipChanType":"UNKNOWN","adultFlag":0,"ipChanCategory":"UNKNOWN","ipChanInteractive":false,"callSign":"UNKNOWN","adFlag":0,"configured":false,"lastUpdated":"","ipChanCpId":"UNKNOWN","isFreeviewPlay":1,"playerService":"com.webos.service.tv","TSID":0,"SVCID":0}],"subscribed":true}}
    End Sub

    Private Sub TreatExternalInputInfoEvent(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatExternalInputInfoEvent called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"externalinput_0","payload":{"devices":[{"id":"AV_1","label":"AV","port":1,"appId":"com.webos.app.externalinput.av1","icon":"http://192.168.1.69:3000/resources/ab9b4e9c4bffbdd51f410c2221b8c2a3102f6420/av.png","modified":false,"subList":[],"subCount":0,"connected":false,"favorite":false},{"id":"HDMI_1","label":"Set-Top Box","port":1,"appId":"com.webos.app.hdmi1","icon":"http://192.168.1.69:3000/resources/ed950809c417e52569bc0c8aeed67d50163ac551/settopbox.png","modified":true,"spdProductDescription":"TiVo","spdVendorName":"Broadcom","spdSourceDeviceInfo":"Digital STB","lastUniqueId":255,"subList":[{"id":"URCU","serviceType":"settop","connectedInput":"HDMI_1","serviceName":"Comcast(Saratoga)","serviceId":"10","serviceArea":"","manufacturerName":"Samsung","manufacturerId":"10210002","settopCode":"C-CA67102","settopOption":"","irType":"C"}],"subCount":1,"connected":true,"favorite":true},{"id":"HDMI_2","label":"HDMI2","port":2,"appId":"com.webos.app.hdmi2","icon":"http://192.168.1.69:3000/resources/0604a74a61bc0f940cf1372169add19b2f7b3a70/HDMI_2.png","modified":false,"lastUniqueId":255,"subList":[],"subCount":0,"connected":false,"favorite":true},{"id":"HDMI_3","label":"HDMI3","port":3,"appId":"com.webos.app.hdmi3","icon":"http://192.168.1.69:3000/resources/33f30cb550c0ad1a09d2188030803ce8c99b96c7/HDMI_3.png","modified":false,"lastUniqueId":255,"subList":[],"subCount":0,"connected":false,"favorite":false},{"id":"HDMI_4","label":"HDMI4","port":4,"appId":"com.webos.app.hdmi4","icon":"http://192.168.1.69:3000/resources/46e19792c140226a116692add4677cd9c1c74efe/HDMI_4.png","modified":false,"lastUniqueId":255,"subList":[],"subCount":0,"connected":false,"favorite":false}],"subscribed":true}}
    End Sub

    Private Sub TreatAudioInfoEvent(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatAudioInfoEvent called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"audiostatus_0","payload":{"returnValue":true,"volumeMax":100,"scenario":"mastervolume_ext_speaker_optical","subscribed":true,"volume":-1,"action":"requested","active":false,"mute":false}}
        Try
            Dim MuteInfo As Object = FindPairInJSONString(Payload, "mute")
            'If MuteInfo Is Nothing Then Exit Sub
            SetMuteState = MuteInfo
            Dim VolumeInfo As Object = FindPairInJSONString(Payload, "volume")
            If VolumeInfo = -1 Then
                SetVolume = 1000
            Else
                SetVolume = VolumeInfo
            End If
            Dim ScenarioInfo As Object = FindPairInJSONString(Payload, "scenario")
            Dim VolMax As Object = FindPairInJSONString(Payload, "volumeMax")
            If VolMax IsNot Nothing Then
                If Val(VolMax) <> MyMaximumVolume Then ' added 11/28/2018 in v .38
                    ' I should now update the slider urggh

                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatAudioInfoEvent called for Device = " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TreatVolumeInfoEvent(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatVolumeInfoEvent called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"volumestatus_0","payload":{"returnValue":true,"volumeMax":100,"muted":false,"scenario":"mastervolume_ext_speaker_optical","subscribed":true,"volume":-1,"action":"requested","active":false}} 
    End Sub

    Private Sub TreatSetAppResponse(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetAppResponse called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"14","payload":{"returnValue":true,"id":"com.showtime.app.showtimeanytime","sessionId":"Y29tLnNob3d0aW1lLmFwcC5zaG93dGltZWFueXRpbWU="}}
    End Sub

    Private Sub TreatForegroundAppResponse(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetTreatForegroundAppResponseAppResponse called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"foregroundapp_0","payload":{"subscribed":true,"appId":"netflix","returnValue":true,"windowId":"","processId":""}} 
        Try
            Dim AppId As Object = FindPairInJSONString(Payload, "appId")
            ' this might be not right to hard code but I'll do it anyway.
            If AppId.ToString.ToLower = "com.webos.app.livetv" Then
                ' we can only subscribe to channel info when the TV is in TV mode
                LGSendCommand("currentchannel_0", "subscribe", "ssap://tv/getCurrentChannel", "", "", True)
            End If
            ' now find this in the remoteinifile
            Dim objRemoteFile As String = gRemoteControlPath
            Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
            Try
                RemoteButtons = GetIniSection(MyUDN, objRemoteFile) '  As Dictionary(Of String, String)
                If RemoteButtons Is Nothing Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatSetAppResponse for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
                    Exit Try
                Else
                    Dim RemoteButtonString As String = ""
                    Dim RemoteButtonName As String = ""
                    Dim RemoteButtonValue As Integer = 0
                    For Each RemoteButton In RemoteButtons
                        If RemoteButton.Key <> "" Then
                            RemoteButtonString = RemoteButton.Value
                            Dim RemoteButtonInfos As String()
                            RemoteButtonInfos = Split(RemoteButtonString, ":;:-:")
                            ' Active ; Given Name ; Key Code ; Row Index ; Column Index -- RowIndex and Column Index start with 1
                            If UBound(RemoteButtonInfos, 1) > 2 Then
                                RemoteButtonName = RemoteButtonInfos(0)
                                RemoteButtonValue = Val(RemoteButton.Key)
                                Dim FoundAppID As String = ""
                                If (RemoteButtonInfos(1).IndexOf("LGlaunch") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetinput") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetchannel") = 0) Then
                                    FoundAppID = RemoteButtonInfos(4)
                                    If FoundAppID = AppId.ToString Then
                                        UpdateLGREmoteState(RemoteButton.Key, RemoteButtonInfos(0))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    Next
                    RemoteButtons = Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatSetAppResponse parsing the .ini file for remote button info for device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatForegroundAppResponse called for Device = " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)

        End Try
    End Sub

    Private Sub TreatCurrentChannelResponse(Payload As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatCurrentChannelResponse called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' {"type":"response","id":"currentchannel_2","payload":{"channelId":"0_2_2_0_0_0_0","dualChannel":{"dualChannelId":null,"dualChannelTypeName":null,"dualChannelTypeId":null,"dualChannelNumber":null},"isScrambled":false,"channelTypeName":"Terrestrial Analog TV","isLocked":false,"isChannelChanged":false,"channelModeName":"Terrestrial","channelNumber":"2-0","isFineTuned":false,"channelTypeId":0,"isDescrambled":false,"isSkipped":false,"isHEVCChannel":false,"hybridtvType":null,"isInvisible":false,"favoriteGroup":null,"channelName":"","channelModeId":0,"signalChannelId":"0_2_2_0_0_0_0"}} 
        Try
            Dim channelId As Object = FindPairInJSONString(Payload, "channelId")
            ' this might be not right to hard code but I'll do it anyway.

            ' now find this in the remoteinifile
            Dim objRemoteFile As String = gRemoteControlPath
            Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
            Try
                RemoteButtons = GetIniSection(MyUDN, objRemoteFile) '  As Dictionary(Of String, String)
                If RemoteButtons Is Nothing Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatCurrentChannelResponse for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
                    Exit Try
                Else
                    Dim RemoteButtonString As String = ""
                    Dim RemoteButtonName As String = ""
                    Dim RemoteButtonValue As Integer = 0
                    For Each RemoteButton In RemoteButtons
                        If RemoteButton.Key <> "" Then
                            RemoteButtonString = RemoteButton.Value
                            Dim RemoteButtonInfos As String()
                            RemoteButtonInfos = Split(RemoteButtonString, ":;:-:")
                            ' Active ; Given Name ; Key Code ; Row Index ; Column Index -- RowIndex and Column Index start with 1
                            If UBound(RemoteButtonInfos, 1) > 2 Then
                                RemoteButtonName = RemoteButtonInfos(0)
                                RemoteButtonValue = Val(RemoteButton.Key)
                                Dim FoundChannelId As String = ""
                                If (RemoteButtonInfos(1).IndexOf("LGlaunch") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetinput") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetchannel") = 0) Then
                                    FoundChannelId = RemoteButtonInfos(4)
                                    If FoundChannelId = channelId.ToString Then
                                        UpdateLGREmoteState(RemoteButton.Key, RemoteButtonInfos(0))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    Next
                    RemoteButtons = Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatCurrentChannelResponse parsing the .ini file for remote button info for device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatCurrentChannelResponse called for Device = " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)

        End Try

    End Sub

    Private Sub SetLGMute(Controlvalue As Integer)
        Select Case Controlvalue
            Case tpLGOn
                LGSendKeyCode("setmute", "true")
            Case tpLGOff
                LGSendKeyCode("setmute", "false")
            Case tpLGToggle
                If MyCurrentMuteState Then
                    LGSendKeyCode("setmute", "false")
                Else
                    LGSendKeyCode("setmute", "true")
                End If
        End Select
    End Sub

    Private Sub setLGVolume(Controlvalue As Integer)
        Select Case Controlvalue
            Case vpLGUp
                LGSendKeyCode("inputvolumeup")
            Case vpLGDown
                LGSendKeyCode("inputvolumedown")
            Case Else
                LGSendKeyCode("setvolume", Controlvalue)
        End Select
    End Sub


    Private Function LGActivateRemote() As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGActivateRemote called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        LGActivateRemote = False
        If Not MyRemoteServiceActive Then
            LGActivateRemote = SendLGRegistration()
        End If
        SetHSRemoteState()
    End Function

    Private Function SendLGRegistration() As Boolean

        SendLGRegistration = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendLGRegistration called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim ReturnJSON As String = ""
        Dim SecWebSocketKey As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "")
        Dim Port As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "")
        If SecWebSocketKey = "" Then Return False

        If LGWebSocket Is Nothing Then
            Try
                LGWebSocket = New WebSocketClient(False)    'dcorssl set to true if you want to test SSL
            Catch ex As Exception
                Log("Error in SendLGRegistration for UPnPDevice = " & MyUPnPDeviceName & " unable to open WebSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyRemoteServiceActive = False
                Return False
            End Try
        End If
        'Port = "3001"  'dcorssl

        'GET / HTTP/1.1
        'Sec-WebSocket-Key: PhlKjlr5qP6gw/T+VHzCZg==
        'Connection: Upgrade
        'Upgrade: websocket
        'Sec-WebSocket - Version:  13
        'Host: 192.168.1.69:3000
        'Cache-Control: no-Cache

        'Http/ 1.1 101 Switching Protocols
        'Upgrade: websocket
        'Connection: Upgrade
        'Sec-WebSocket - Accept: oR5ZpF68Y2sHNlDMM4jQZSHMeKg =

        ' this could be the JSON way of registering
        '  {"type":"register","payload":{"manifest":{"permissions":["LAUNCH","LAUNCH_WEBAPP","APP_TO_APP","CONTROL_AUDIO","CONTROL_INPUT_MEDIA_PLAYBACK","CONTROL_POWER","READ_INSTALLED_APPS","CONTROL_DISPLAY","CONTROL_INPUT_JOYSTICK","CONTROL_INPUT_MEDIA_RECORDING","CONTROL_INPUT_TV","READ_INPUT_DEVICE_LIST","READ_NETWORK_STATE","READ_TV_CHANNEL_LIST","WRITE_NOTIFICATION_TOAST","CONTROL_INPUT_TEXT","CONTROL_MOUSE_AND_KEYBOARD","READ_CURRENT_CHANNEL","READ_RUNNING_APPS"]}}}

        AddHandler LGWebSocket.DataReceived, AddressOf HandleLGDataReceived
        AddHandler LGWebSocket.WebSocketClosed, AddressOf HandleLGSocketClosed

        If Not LGWebSocket.ConnectSocket(MyIPAddress, Port) Then
            Return False
        End If

        ' wait until connected. Important for SSL as it takes longer
        Dim WaitLoopCounter As Integer = 0
        While LGWebSocket.MySocketIsClosed
            wait(1)
            WaitLoopCounter += 1
            If WaitLoopCounter > 10 Then Exit While
        End While

        LGWebSocket.Receive()

        If Not LGWebSocket.UpgradeWebSocket("", SecWebSocketKey, 20, True) Then
            LGWebSocket.CloseSocket()
            Return False
        End If

        Dim ClientKey As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diLGClientKey.ToString, "")

        ' https://github.com/aurodionov/lgtv2/blob/master/pairing.json
        ' https://github.com/msloth/lgtv.js

        'Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""forcePairing"":false,""pairingType"":""PROMPT"",""manifest"":{""manifestVersion"":1,""appVersion"":""1.1"",""signed"":{""created"":""20140509"",""appId"":""com.lge.test"",""vendorId"":""com.lge"",""localizedAppNames"":{"""":""LG Remote App"",""ko-KR"":""리모컨 앱"",""zxx-XX"":""ЛГ Rэмotэ AПП""},""localizedVendorNames"":{"""":""LG Electronics""},""permissions"":[""TEST_SECURE"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_INSTALLED_APPS"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""CONTROL_POWER"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME""],""serial"":""2f930e2d2cfe083771f68e4fe7bb07""},""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CLOSE"",""TEST_OPEN"",""TEST_PROTECTED"",""CONTROL_AUDIO"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_INPUT_TV"",""CONTROL_POWER"",""READ_APP_STATUS"",""READ_CURRENT_CHANNEL"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_RUNNING_APPS"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""READ_POWER_STATE"",""READ_COUNTRY_INFO""],""signatures"":[{""signatureVersion"":1,""signature"":""eyJhbGdvcml0aG0iOiJSU0EtU0hBMjU2Iiwia2V5SWQiOiJ0ZXN0LXNpZ25pbmctY2VydCIsInNpZ25hdHVyZVZlcnNpb24iOjF9.hrVRgjCwXVvE2OOSpDZ58hR+59aFNwYDyjQgKk3auukd7pcegmE2CzPCa0bJ0ZsRAcKkCTJrWo5iDzNhMBWRyaMOv5zWSrthlf7G128qvIlpMT0YNY+n/FaOHE73uLrS/g7swl3/qH/BGFG2Hu4RlL48eb3lLKqTt2xKHdCs6Cd4RMfJPYnzgvI4BNrFUKsjkcu+WD4OO2A27Pq1n50cMchmcaXadJhGrOqH5YmHdOCj5NSHzJYrsW0HPlpuAx/ECMeIZYDh6RMqaFM2DXzdKX9NmmyqzJ3o/0lkk/N97gfVRLW5hA29yeAwaCViZNCP8iC9aO0q9fQojoa7NQnAtw==""}]}}}"
        'Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""forcePairing"":false,""pairingType"":""PROMPT"",""manifest"":{""manifestVersion"":1,""appVersion"":""1.1"",""signed"":{""created"":""20140509"",""appId"":""com.lge.test"",""vendorId"":""com.lge"",""localizedAppNames"":{"""":""LG Remote App"",""ko-KR"":""ë¦¬ëª¨ì»¨ ì±"",""zxx-XX"":""ÐÐ RÑÐ¼otÑ AÐÐ""},""localizedVendorNames"":{"""":""LG Electronics""},""permissions"":[""TEST_SECURE"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_INSTALLED_APPS"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""CONTROL_POWER"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME""],""serial"":""2f930e2d2cfe083771f68e4fe7bb07""},""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CLOSE"",""TEST_OPEN"",""TEST_PROTECTED"",""CONTROL_AUDIO"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_INPUT_TV"",""CONTROL_POWER"",""READ_APP_STATUS"",""READ_CURRENT_CHANNEL"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_RUNNING_APPS"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""READ_POWER_STATE"",""READ_COUNTRY_INFO""],""signatures"":[{""signatureVersion"":1,""signature"":""eyJhbGdvcml0aG0iOiJSU0EtU0hBMjU2Iiwia2V5SWQiOiJ0ZXN0LXNpZ25pbmctY2VydCIsInNpZ25hdHVyZVZlcnNpb24iOjF9.hrVRgjCwXVvE2OOSpDZ58hR+59aFNwYDyjQgKk3auukd7pcegmE2CzPCa0bJ0ZsRAcKkCTJrWo5iDzNhMBWRyaMOv5zWSrthlf7G128qvIlpMT0YNY+n/FaOHE73uLrS/g7swl3/qH/BGFG2Hu4RlL48eb3lLKqTt2xKHdCs6Cd4RMfJPYnzgvI4BNrFUKsjkcu+WD4OO2A27Pq1n50cMchmcaXadJhGrOqH5YmHdOCj5NSHzJYrsW0HPlpuAx/ECMeIZYDh6RMqaFM2DXzdKX9NmmyqzJ3o/0lkk/N97gfVRLW5hA29yeAwaCViZNCP8iC9aO0q9fQojoa7NQnAtw==""}]}}}"
        'Dim JSONAlreadyRegisteredString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""forcePairing"":false,""pairingType"":""PROMPT"",""client-key"":""" & ClientKey & """,""manifest"":{""manifestVersion"":1,""appVersion"":""1.1"",""signed"":{""created"":""20140509"",""appId"":""com.lge.test"",""vendorId"":""com.lge"",""localizedAppNames"":{"""":""LG Remote App"",""ko-KR"":""ë¦¬ëª¨ì»¨ ì±"",""zxx-XX"":""ÐÐ RÑÐ¼otÑ AÐÐ""},""localizedVendorNames"":{"""":""LG Electronics""},""permissions"":[""TEST_SECURE"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_INSTALLED_APPS"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""CONTROL_POWER"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME""],""serial"":""2f930e2d2cfe083771f68e4fe7bb07""},""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CLOSE"",""TEST_OPEN"",""TEST_PROTECTED"",""CONTROL_AUDIO"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_INPUT_TV"",""CONTROL_POWER"",""READ_APP_STATUS"",""READ_CURRENT_CHANNEL"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_RUNNING_APPS"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""READ_POWER_STATE"",""READ_COUNTRY_INFO""],""signatures"":[{""signatureVersion"":1,""signature"":""eyJhbGdvcml0aG0iOiJSU0EtU0hBMjU2Iiwia2V5SWQiOiJ0ZXN0LXNpZ25pbmctY2VydCIsInNpZ25hdHVyZVZlcnNpb24iOjF9.hrVRgjCwXVvE2OOSpDZ58hR+59aFNwYDyjQgKk3auukd7pcegmE2CzPCa0bJ0ZsRAcKkCTJrWo5iDzNhMBWRyaMOv5zWSrthlf7G128qvIlpMT0YNY+n/FaOHE73uLrS/g7swl3/qH/BGFG2Hu4RlL48eb3lLKqTt2xKHdCs6Cd4RMfJPYnzgvI4BNrFUKsjkcu+WD4OO2A27Pq1n50cMchmcaXadJhGrOqH5YmHdOCj5NSHzJYrsW0HPlpuAx/ECMeIZYDh6RMqaFM2DXzdKX9NmmyqzJ3o/0lkk/N97gfVRLW5hA29yeAwaCViZNCP8iC9aO0q9fQojoa7NQnAtw==""}]}}}"
        'Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""forcePairing"":false,""pairingType"":""PROMPT"",""manifest"":{""manifestVersion"":1,""appVersion"":""1.1"",""signed"":{""created"":""20140509"",""appId"":""com.lge.test"",""vendorId"":""com.lge"",""localizedAppNames"":{"""":""LG Remote App"",""ko-KR"":""????????? ???"",""zxx-XX"":""???? R????ot?? A????""},""localizedVendorNames"":{"""":""LG Electronics""},""permissions"":[""TEST_SECURE"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_INSTALLED_APPS"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""CONTROL_POWER"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME""],""serial"":""2f930e2d2cfe083771f68e4fe7bb07""},""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CLOSE"",""TEST_OPEN"",""TEST_PROTECTED"",""CONTROL_AUDIO"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_INPUT_TV"",""CONTROL_POWER"",""READ_APP_STATUS"",""READ_CURRENT_CHANNEL"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_RUNNING_APPS"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""READ_POWER_STATE"",""READ_COUNTRY_INFO""],""signatures"":[{""signatureVersion"":1,""signature"":""eyJhbGdvcml0aG0iOiJSU0EtU0hBMjU2Iiwia2V5SWQiOiJ0ZXN0LXNpZ25pbmctY2VydCIsInNpZ25hdHVyZVZlcnNpb24iOjF9.hrVRgjCwXVvE2OOSpDZ58hR+59aFNwYDyjQgKk3auukd7pcegmE2CzPCa0bJ0ZsRAcKkCTJrWo5iDzNhMBWRyaMOv5zWSrthlf7G128qvIlpMT0YNY+n/FaOHE73uLrS/g7swl3/qH/BGFG2Hu4RlL48eb3lLKqTt2xKHdCs6Cd4RMfJPYnzgvI4BNrFUKsjkcu+WD4OO2A27Pq1n50cMchmcaXadJhGrOqH5YmHdOCj5NSHzJYrsW0HPlpuAx/ECMeIZYDh6RMqaFM2DXzdKX9NmmyqzJ3o/0lkk/N97gfVRLW5hA29yeAwaCViZNCP8iC9aO0q9fQojoa7NQnAtw==""}]}}}"
        'Dim JSONAlreadyRegisteredString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""forcePairing"":false,""pairingType"":""PROMPT"",""client-key"":""" & ClientKey & """,""manifest"":{""manifestVersion"":1,""appVersion"":""1.1"",""signed"":{""created"":""20140509"",""appId"":""com.lge.test"",""vendorId"":""com.lge"",""localizedAppNames"":{"""":""LG Remote App"",""ko-KR"":""????????? ???"",""zxx-XX"":""???? R????ot?? A????""},""localizedVendorNames"":{"""":""LG Electronics""},""permissions"":[""TEST_SECURE"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_INSTALLED_APPS"",""READ_LGE_SDX"",""READ_NOTIFICATIONS"",""SEARCH"",""WRITE_SETTINGS"",""WRITE_NOTIFICATION_ALERT"",""CONTROL_POWER"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS"",""READ_UPDATE_INFO"",""UPDATE_FROM_REMOTE_APP"",""READ_LGE_TV_INPUT_EVENTS"",""READ_TV_CURRENT_TIME""],""serial"":""2f930e2d2cfe083771f68e4fe7bb07""},""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CLOSE"",""TEST_OPEN"",""TEST_PROTECTED"",""CONTROL_AUDIO"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_INPUT_TV"",""CONTROL_POWER"",""READ_APP_STATUS"",""READ_CURRENT_CHANNEL"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_RUNNING_APPS"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""READ_POWER_STATE"",""READ_COUNTRY_INFO""],""signatures"":[{""signatureVersion"":1,""signature"":""eyJhbGdvcml0aG0iOiJSU0EtU0hBMjU2Iiwia2V5SWQiOiJ0ZXN0LXNpZ25pbmctY2VydCIsInNpZ25hdHVyZVZlcnNpb24iOjF9.hrVRgjCwXVvE2OOSpDZ58hR+59aFNwYDyjQgKk3auukd7pcegmE2CzPCa0bJ0ZsRAcKkCTJrWo5iDzNhMBWRyaMOv5zWSrthlf7G128qvIlpMT0YNY+n/FaOHE73uLrS/g7swl3/qH/BGFG2Hu4RlL48eb3lLKqTt2xKHdCs6Cd4RMfJPYnzgvI4BNrFUKsjkcu+WD4OO2A27Pq1n50cMchmcaXadJhGrOqH5YmHdOCj5NSHzJYrsW0HPlpuAx/ECMeIZYDh6RMqaFM2DXzdKX9NmmyqzJ3o/0lkk/N97gfVRLW5hA29yeAwaCViZNCP8iC9aO0q9fQojoa7NQnAtw==""}]}}}"
        Dim JSONRegisterString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS""],""manifestVersion"":1}}}"
        Dim JSONAlreadyRegisteredString As String = "{""type"":""register"",""id"":""register_0"",""payload"":{""pairingType"":""PROMPT"",""client-key"":""" & ClientKey & """,""manifest"":{""permissions"":[""LAUNCH"",""LAUNCH_WEBAPP"",""APP_TO_APP"",""CONTROL_AUDIO"",""CONTROL_INPUT_MEDIA_PLAYBACK"",""CONTROL_POWER"",""READ_INSTALLED_APPS"",""CONTROL_DISPLAY"",""CONTROL_INPUT_JOYSTICK"",""CONTROL_INPUT_MEDIA_RECORDING"",""CONTROL_INPUT_TV"",""READ_INPUT_DEVICE_LIST"",""READ_NETWORK_STATE"",""READ_TV_CHANNEL_LIST"",""WRITE_NOTIFICATION_TOAST"",""CONTROL_INPUT_TEXT"",""CONTROL_MOUSE_AND_KEYBOARD"",""READ_CURRENT_CHANNEL"",""READ_RUNNING_APPS""],""manifestVersion"":1}}}"

        If ClientKey <> "" Then
            JSONRegisterString = JSONAlreadyRegisteredString
        End If

        ' this should happen first because there might be scenari where pairingType = PROMPT is not supported
        LGSendCommand("hello_0", "hello", "", "", "", True) ' check some capabilities

        Return True

    End Function

    Public Sub HandleLGDataReceived(send As Object, inBytes As Byte())
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGDataReceived called for Device = " & MyUPnPDeviceName & " and Data = " & Encoding.UTF8.GetString(inBytes, 0, inBytes.Length), LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGDataReceived called for Device = " & MyUPnPDeviceName & " Datasize = " & inBytes.Length.ToString, LogType.LOG_TYPE_INFO)
        Dim Type As String = ""
        Dim Id As String = ""
        Try
            Type = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "type").ToString.Trim("""")
            If Type = "" Then Exit Sub
        Catch ex As Exception
            Log("Error in HandleLGDataReceived for UPnPDevice = " & MyUPnPDeviceName & " retrieving json info with Type = " & Type & " and ID = " & Id & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Id = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "id").ToString.Trim("""")
        Try
            Select Case Type
                Case "registered"
                    If Id = "register_0" Then
                        TreatRegistered(inBytes)
                    End If
                Case "response"
                    ' If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGDataReceived received a response     for Device = " & MyUPnPDeviceName & " with Error = " & ReturnError, LogType.LOG_TYPE_INFO)
                    Dim Payload As Object = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "payload")
                    If Payload Is Nothing Then Exit Sub
                    Dim returnValue As Object = Nothing
                    Try
                        returnValue = FindPairInJSONString(Payload, "returnValue")
                    Catch ex As Exception
                        returnValue = ""
                    End Try
                    If returnValue.ToString.ToLower = "false" Then Exit Sub
                    If Id = "listLaunchPoints_0" Then
                        ' this is the response to us getting the list of apps (Launchpoints)
                        LGAddAppButtons(Payload)
                    ElseIf Id = "getexternalinputlist_0" Then
                        ' this is the response to us getting the list of input to choose from
                        AddInputButtons(Payload)
                    ElseIf Id = "services_0" Then
                        TreatServicesInfo(Payload)
                    ElseIf Id = "channellist_0" Then
                        AddChannelButtons(Payload)
                    ElseIf Id = "getpointerinputsocket_0" Then
                        ' this should be used to open a second websocket to allow entry of numbers etc. ie the old remote control stuff
                        TreatPointerInputSocket(Payload)
                    ElseIf Id = "channels_0" Then
                        TreatChannelInfoEvent(Payload)
                    ElseIf Id = "externalinput_0" Then
                        TreatExternalInputInfoEvent(Payload)
                    ElseIf Id = "audiostatus_0" Then
                        TreatAudioInfoEvent(Payload)
                    ElseIf Id = "volumestatus_0" Then
                        TreatVolumeInfoEvent(Payload)
                    ElseIf Id = "keyboardstatus_0" Then
                    ElseIf Id = "interestingevents_0" Then
                    ElseIf Id = "foregroundapp_0" Then
                        TreatForegroundAppResponse(Payload)
                    ElseIf Id = "setapp_0" Then
                        TreatSetAppResponse(Payload)
                    ElseIf Id = "currentchannel_0" Then
                        TreatCurrentChannelResponse(Payload)
                    Else

                    End If
                Case "hello"
                    TreatHelloReceived(inBytes)
                Case "error"
                    Dim ReturnError As String = ""
                    Try
                        ReturnError = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "error").ToString.Trim("""")
                    Catch ex As Exception
                        ReturnError = ""
                    End Try
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGDataReceived received a response error for Device = " & MyUPnPDeviceName & " with Error = " & ReturnError, LogType.LOG_TYPE_INFO)
                    If Id = "register_0" Then
                        ' we registered unsuccessfully
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diLGClientKey.ToString, "")
                        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                        LGCloseWebSocket()
                        MyRemoteServiceActive = False
                        SetHSRemoteState()
                    Else
                        Dim Payload As Object = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "payload")
                        If Payload Is Nothing Then Exit Sub
                        ' returnValue typically true , if false it comes with type error ..... or so I think
                        ' errorCode
                        ' errorText
                    End If
            End Select
        Catch ex As Exception
            Log("Error in HandleLGDataReceived for UPnPDevice = " & MyUPnPDeviceName & " with Type = " & Type & " and ID = " & Id & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub HandleLGSocketClosed(sender As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGSocketClosed called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        MyRemoteServiceActive = False
        Try
            If LGWebSocket IsNot Nothing Then
                RemoveHandler LGWebSocket.DataReceived, AddressOf HandleLGDataReceived
                RemoveHandler LGWebSocket.WebSocketClosed, AddressOf HandleLGSocketClosed
            End If
        Catch ex As Exception
        End Try
        LGWebSocket = Nothing
        ' maybe some more actions are needed here, like updating the HS status of the remote
        Try
            If HSRefRemote <> -1 Then hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
            MyRemoteServiceActive = False
        Catch ex As Exception
            Log("Error in HandleLGSocketClosed for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub HandleLGPointerDataReceived(send As Object, inBytes As Byte())
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGPointerDataReceived called for Device = " & MyUPnPDeviceName & " and Data = " & Encoding.UTF8.GetString(inBytes, 0, inBytes.Length), LogType.LOG_TYPE_INFO)
        Dim Type As String = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "type").ToString.Trim("""")
        If Type = "" Then Exit Sub
        Dim Id As String = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "id").ToString.Trim("""")
        Select Case Type
            Case "registered"
            Case "response"
                Dim Payload As Object = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "payload")
                If Payload Is Nothing Then Exit Sub
                Dim returnValue As Object = FindPairInJSONString(Payload, "returnValue")
                If returnValue.ToString.ToLower <> "true" Then Exit Sub
            Case "error"
                Dim ReturnError As String = FindPairInJSONString(ASCIIEncoding.ASCII.GetChars(inBytes), "error").ToString.Trim("""")
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGPointerDataReceived received a response error for Device = " & MyUPnPDeviceName & " with Error = " & ReturnError, LogType.LOG_TYPE_INFO)
        End Select

    End Sub

    Public Sub HandleLGPointerSocketClosed(sender As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLGPointerSocketClosed called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            If LGWebPointerSocket IsNot Nothing Then
                RemoveHandler LGWebPointerSocket.DataReceived, AddressOf HandleLGPointerDataReceived
                RemoveHandler LGWebPointerSocket.WebSocketClosed, AddressOf HandleLGPointerSocketClosed
            End If
        Catch ex As Exception
        End Try
        LGWebPointerSocket = Nothing
    End Sub

    Private Sub LGCloseWebSocket()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGCloseWebSocket called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        'If Not MyRemoteServiceActive Then Exit Sub

        If LGWebSocket IsNot Nothing Then
            Try
                LGWebSocket.CloseSocket()
            Catch ex As Exception
                Log("Error in LGCloseWebSocket for UPnPDevice = " & MyUPnPDeviceName & " closing the websocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                RemoveHandler LGWebSocket.DataReceived, AddressOf HandleLGDataReceived
                RemoveHandler LGWebSocket.WebSocketClosed, AddressOf HandleLGSocketClosed
            Catch ex As Exception
            End Try
            LGWebSocket = Nothing
        End If
        If LGWebPointerSocket IsNot Nothing Then
            Try
                LGWebPointerSocket.CloseSocket()
            Catch ex As Exception
                Log("Error in LGCloseWebSocket for UPnPDevice = " & MyUPnPDeviceName & " closing the websocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                RemoveHandler LGWebPointerSocket.DataReceived, AddressOf HandleLGPointerDataReceived
                RemoveHandler LGWebPointerSocket.WebSocketClosed, AddressOf HandleLGPointerSocketClosed
            Catch ex As Exception
            End Try
            LGWebPointerSocket = Nothing
        End If

        Try
            If HSRefRemote <> -1 Then hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
            MyRemoteServiceActive = False
        Catch ex As Exception
            Log("Error in CloseTCPConnection 2 for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        SetHSRemoteState()
    End Sub

    Private Function LGSendCommand(Prefix As String, Msgtype As String, Uri As String, Payload As Object, fn As String, Optional SkipCommandCount As Boolean = False) As Boolean
        'The `Msgtype` Is either (at least, there may be more),

        '* `request` - a single request, eg get volume
        '* `response` - response to a request, Or subscription event
        '* `subscribe` - subscribe to a topic ie get notifications when something happens, eg channel Is changed
        '* `unsubscribe` - unsubscribe a subscribed topic

        'The `id` Is a concatenation Of the command And a message counter, Like so:

        'Request:
        '{"type""request","id":"status_3", ...}

        'Response:
        '{"type""response","id":"status_3", ...}

        'This Is used so that a request can be matched with a response.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendCommand called for device - " & MyUPnPDeviceName & " with Prefix = " & Prefix & ", MsgType = " & Msgtype & ", URI = " & Uri & ", Payload = " & Payload, LogType.LOG_TYPE_INFO)

        LGCommandCount += 1
        Dim SendCommandString As String = ""

        If SkipCommandCount Then
            SendCommandString = "{""id"":""" & Prefix & """,""type"":""" & Msgtype & """"
        Else
            SendCommandString = "{""id"":""" & Prefix & LGCommandCount.ToString & """,""type"":""" & Msgtype & """"
        End If
        If Uri <> "" Then
            SendCommandString &= ",""uri"":""" & Uri & """"
        End If
        If TypeOf (Payload) Is String Then
            If Payload <> "" Then SendCommandString &= ",""payload"":" & Payload
        Else

        End If
        SendCommandString &= "}"
        Try
            If Not LGWebSocket.SendDataOverWebSocket(OpcodeText, System.Text.ASCIIEncoding.ASCII.GetBytes(SendCommandString), True) Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendCommand for device - " & MyUPnPDeviceName & " unsuccessful sending command", LogType.LOG_TYPE_INFO)
                Return False
            End If
        Catch ex As Exception
            Log("Error in LGSendCommand for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return True
    End Function

    Private Function LGSendPointerCommand(bType As String, Message As String) As Boolean
        ' pointerSocket
        ' =============
        ' Send Log Message
        ' pointerSocket.sendTextMessage(message)
        ' Send move (dx,dy)
        ' pointerSocket.sendLogMessage('type:move\ndx:' + dx + '\ndy:' + dy + '\ndown:0\n\n')
        ' Send scroll (dy)
        ' pointerSocket.sendLogMessage('type:scroll\ndx:0\ndy:' + dy + '\ndown:0\n\n')
        ' SendClick ()
        ' pointerSocket.sendLogMessage('type:click\n\n')
        ' SendInput (btype, bname) btype = "button" 
        ' pointerSocket.sendLogMessage('type:' + btype + '\nname:' + bname + '\n\n')
        ' input buttons here https://github.com/CODeRUS/harbour-lgremote-webos/blob/6034950cf7bb22a369b6e8e6d56c8dbcfc0bb533/qml/pages/ActionsPanel.qml
        ' REWIND GOTOPREV GOTONEXT FASTFORWARD STOP PLAY PAUSE RED GREEN YELLOW BLUE HOME UP 3D_MODE LEFT ENTER RIGHT EXIT DOWN BACK CHANNELUP CHANNELDOWN (and I think just numbers)

        ' msg := "type:" + btype + "\n" + "name:" + bname + "\n\n"
        ' Return ps.writeMessage(websocket.TextMessage, []byte(msg))
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendPointerCommand called for device - " & MyUPnPDeviceName & " with bType = " & bType & " and Message = " & Message, LogType.LOG_TYPE_INFO)
        Dim SendCommandString As String = ""

        Select Case bType
            Case "button"
                SendCommandString = "type:button" & vbLf & "name:" & Message & vbLf & vbLf
            Case "click"
                SendCommandString = "type:click" & vbLf & vbLf
        End Select

        Try
            If Not LGWebPointerSocket.SendDataOverWebSocket(OpcodeText, System.Text.ASCIIEncoding.ASCII.GetBytes(SendCommandString), True) Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendPointerCommand for device - " & MyUPnPDeviceName & " unsuccessful sending command", LogType.LOG_TYPE_INFO)
                Return False
            End If
        Catch ex As Exception
            Log("Error in LGSendPointerCommand for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return True
    End Function

    Private Sub UpdateLGREmoteState(Value As Integer, ValueString As String)
        Try
            Try
                If HSRefServiceRemote <> -1 Then hs.SetDeviceValueByRef(HSRefServiceRemote, Value, True)
                If HSRefServiceRemote <> -1 Then hs.SetDeviceString(HSRefServiceRemote, ValueString, True)
                '   If piDebuglevel > DebugLevel.dlEvents Then Log("HS updated in UpdateTransportState. HSRef = " & HSRefPlayer & " and MyTransportStateHasChanged = " & MyTransportStateHasChanged.ToString & ", MyTrackInfoHasChanged = " & MyTrackInfoHasChanged.ToString & ". Info = " & TransportInfo, LogType.LOG_TYPE_INFO)
                'End If
            Catch ex As Exception
                Log("Error in UpdateLGREmoteState updating HS with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("ERROR in UpdateLGREmoteState 6 for UPnPDevice = " & MyUPnPDeviceName & " with error=" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Public Function LGSendKeyCode(KeyCode As String, Optional Param1 As String = "", Optional Param2 As String = "") As Boolean
        LGSendKeyCode = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendKeyCode called for UPnPDevice = " & MyUPnPDeviceName & " with KeyCode = " & KeyCode, LogType.LOG_TYPE_INFO)

        ' a key
        ' {"type":"response","id":"status_0","payload":{"scenario":"mastervolume_tv_speaker","active":false,"action":"requested","volume":0,"returnValue":true,"subscribed":true,"mute":false}}

        ' webbrowser
        '  {"type":"response","id":"0","payload":{"returnValue":true,"id":"com.webos.app.browser","sessionId":"Y29tLndlYm9zLmFwcC5icm93c2VyOnVuZGVmaW5lZA=="}}

        ' var send_command = function(prefix, msgtype, uri, payload, fn) {        command_count++;
        '       var msg = '{"id":"' + prefix + command_count + '","type":"' + msgtype + '","uri":"' + uri + '"';
        'If (TypeOf payload === 'string' && payload.length > 0) {
        'msg += ',"payload":' + payload + "}";
        '  send_command("", "request", "ssap://system.notifications/createToast", '{"message": "MSG"}'.replace('MSG', text), fn);
        ' send_command("", "request", "ssap://system.launcher/open", JSON.stringify({target: url}), function(err, resp){
        ' send_command("", "request", "ssap://system/turnOff", null, fn);
        '  send_command("", "request", "ssap://system.notifications/createToast", '{"message": "MSG"}'.replace('MSG', text), fn);

        If LGWebSocket Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning LGSendKeyCode called for UPnPDevice = " & MyUPnPDeviceName & " with KeyCode = " & KeyCode & " but no Open Socket", LogType.LOG_TYPE_WARNING)
            Return False
        End If

        ' https://github.com/hobbyquaker/lgtv2mqtt/blob/master/index.js
        ' https://github.com/CODeRUS/harbour-lgremote-webos/blob/6034950cf7bb22a369b6e8e6d56c8dbcfc0bb533/qml/pages/MainSocket.qml
        ' https://www.javatips.net/api/popcorn-android-master/connectsdk/core/src/com/connectsdk/service/WebOSTVService.java
        ' https://www.javatips.net/api/popcorn-android-master/connectsdk/core/src/com/connectsdk/service/webos/WebOSTVKeyboardInput.java
        ' https://forum.fhem.de/index.php?topic=66671.435;wap2
        ' Sending command: {"id":"request_1505496770.15871","client-key":"20bfc93ff6574d66bc70712d548d4e2d","uri":["ssap://com.webos.applicationManager/getForegroundAppInfo"],"type":"request"}
        ' Sending command: {"uri":["ssap://com.webos.service.tv.display/get3DStatus"],"type":"request","client-key":"20bfc93ff6574d66bc70712d548d4e2d","id":"request_1505496772.15841"}
        ' https://github.com/mhop/fhem-mirror/blob/master/fhem/FHEM/82_LGTV_WebOS.pm

        Select Case KeyCode
            Case "unsubscribe"
                LGSendCommand("", "unsubscribe", "", "", "")
            Case "showfloat"
                LGSendCommand("", "request", "ssap://system.notifications/createToast", "{""message"": """ & Param1 & """}", "")
                'LGSendCommand("", "request", "ssap://system.notifications/createAlert", "{""message"": """ & Param1 & """}", "")
                ' https://github.com/msloth/lgtv.js/issues/15
                'http://webosose.org/develop/ls2-api/ls2-api-reference/com-webos-notification/
                ' you can also send messages with a custom icon with a base64 encoded string of an image
                ' {"id""message","type":"request","uri":"ssap://system.notifications/createToast","payload":{"message": "%@", "iconData": "%@", "iconExtension": "png"}}
            Case "openbrowserat"
                LGSendCommand("", "request", "ssap://system.launcher/open", "{""target"": """ & Param1 & """}", "")  ' launchWithPayload({"id": "com.webos.app.browser", "target": url})
            Case "turnoff"
                LGSendCommand("", "request", "ssap://system/turnOff", "", "")
            Case "channellist"
                LGSendCommand("channels_", "request", "ssap://tv/getChannelList", "", "")
            Case "subscribechannellist"
                LGSendCommand("channels_", "subscribe", "ssap://tv/getChannelList", "", "")
            Case "channel"
                LGSendCommand("channels_", "request", "ssap://tv/getCurrentChannel", "", "")
            Case "subscribechannel"
                LGSendCommand("channels_", "subscribe", "ssap://tv/getCurrentChannel", "", "")
            Case "LGsetchannel"
                LGSendCommand("", "request", "ssap://tv/openChannel", "{""channelId"": """ & Param1 & """}", "")
            Case "inputlist"
                LGSendCommand("", "request", "ssap://tv/getExternalInputList", "", "")
            Case "subscribeinputlist"
                LGSendCommand("", "subscribe", "ssap://tv/getExternalInputList", "", "")
            Case "LGsetinput"   ' "setinput"
                LGSendCommand("", "request", "ssap://tv/switchInput", "{""inputId"": """ & Param1 & """}", "")
            Case "setmute"
                If Param1.ToLower = "false" Then
                    LGSendCommand("", "request", "ssap://audio/setMute", "{""mute"": false}", "")
                Else
                    LGSendCommand("", "request", "ssap://audio/setMute", "{""mute"": true}", "")
                End If
            Case "togglemute"
                If Not MyCurrentMuteState Then
                    LGSendCommand("", "request", "ssap://audio/setMute", "{""mute"": false}", "")
                Else
                    LGSendCommand("", "request", "ssap://audio/setMute", "{""mute"": true}", "")
                End If
            Case "getmute"
                LGSendCommand("status_", "request", "ssap://audio/getStatus", "", "")
            Case "subscribestatus"
                LGSendCommand("status_", "subscribe", "ssap://audio/getStatus", "", "")
            Case "getvolume"
                LGSendCommand("status_", "request", "ssap://audio/getVolume", "", "")
            Case "subscribevolume"
                LGSendCommand("status_", "subscribe", "ssap://audio/getVolume", "", "")
            Case "setvolume"
                LGSendCommand("", "request", "ssap://audio/setVolume", "{""volume"": " & Param1.ToString & "}", "")
            Case "inputmediaplay"
                LGSendCommand("", "request", "ssap://media.controls/play", "", "")
            Case "inputmediastop"
                LGSendCommand("", "request", "ssap://media.controls/stop", "", "")
            Case "inputmediapause"
                LGSendCommand("", "request", "ssap://media.controls/pause", "", "")
            Case "inputmediarewind"
                LGSendCommand("", "request", "ssap://media.controls/rewind", "", "")
            Case "inputmediaforward"
                LGSendCommand("", "request", "ssap://media.controls/fastForward", "", "")
            Case "inputchannelup"
                LGSendCommand("", "request", "ssap://tv/channelUp", "", "")
            Case "inputchanneldown"
                LGSendCommand("", "request", "ssap://tv/channelDown", "", "")
            Case "getstatus"
                LGSendCommand("status_", "request", "ssap://audio/getStatus", "", "")
            Case "swinfo"
                LGSendCommand("sw_info_", "request", "ssap://com.webos.service.update/getCurrentSWInformation", "", "")
            Case "services"
                LGSendCommand("services_", "request", "ssap://api/getServiceList", "", "")
            Case "apps"
                LGSendCommand("launcher_", "request", "ssap://com.webos.applicationManager/listLaunchPoints", "", "")
                'LGSendCommand("", "request", "ssap://com.webos.applicationManager/listApps", "", "")
                'LGSendCommand("launcher_", "request", "ssap://com.webos.applicationManager/listApps", "", "")
            Case "openappwithpayload"
                LGSendCommand("", "request", "ssap://com.webos.applicationManager/launch", Param1, "")
            Case "LGlaunch" '"startapp"
                LGSendCommand("setapp_0", "request", "ssap://system.launcher/launch", "{""id"": """ & Param1 & """}", "", True)
            Case "closeapp"
                LGSendCommand("", "request", "ssap://system.launcher/close", "{""id"": """ & Param1 & """}", "")
                ' mainSocket.sendCommand("", "request", "ssap://system.launcher/close", {"id": appId, "sessionId": Qt.btoa(appId + ":undefined")})
            Case "inputenter"
                LGSendCommand("", "request", "ssap://com.webos.service.ime/sendEnterKey", "", "")
            Case "inputvolumeup"
                LGSendCommand("volumeup", "request", "ssap://audio/volumeUp", "", "")
            Case "inputvolumedown"
                LGSendCommand("volumedown_", "request", "ssap://audio/volumeDown", "", "")
            Case "inputbackspace"
                LGSendCommand("pause_", "request", "ssap://com.webos.service.ime/deleteCharacters", "{""{count: """ & Param1 & """}", "")
            Case "openyoutubeaturl"
                LGSendCommand("", "request", "ssap://system.launcher/launch", "{""id"": ""youtube.leanback.v4"", ""params"": {""contentTarget"": ""http://www.youtube.com/tv?v=" & Param1 & """}}", "")
            Case "getappstatus"
                'LGSendCommand("", "request", "ssap://com.webos.service.appstatus/getAppStatus", "", "")
                LGSendCommand("", "request", "ssap://system.launcher/getAppState", "", "")
            Case "hello"
                LGSendCommand("", "hello", "", "", "")
            Case "setpin"
                LGSendCommand("pin_", "request", "ssap://pairing/setPin", "{""pin"":""" & Param1 & """}", "")
            Case "getpointer"
                LGSendCommand("", "request", "ssap://com.webos.service.networkinput/getPointerInputSocket", """{}""", "")
            Case "sendtext"
                ' mainSocket.sendCommand("", "request", "ssap://com.webos.service.ime/insertText", {"text": text, "replace": replace == true})
            Case "pointerbutton"
                LGSendPointerCommand("button", Param1)
            Case "pointerclick"
                LGSendPointerCommand("click", Param1)
            Case "subscribeappstatus"
                LGSendCommand("", "subscribe", "ssap://com.webos.service.appstatus/getAppStatus", "", "")
                LGSendCommand("", "subscribe", "ssap://system.launcher/getAppState", "", "")
            Case "test1"
                LGSendCommand("currentchannel_1", "request", "ssap://tv/getCurrentChannel", "", "", True)
            Case "test2"
                LGSendCommand("currentchannel_2", "subscribe", "ssap://tv/getCurrentChannel", "", "", True)
                LGSendCommand("interestingevents_2", "subscribe", "ssap://com.webos.service.tv.keymanager/listInterestingEvents", "", "", True)
            Case "test3"
                LGSendCommand("interestingevents_1", "request", "ssap://com.webos.service.tv.keymanager/listInterestingEvents", "", "", True)
                LGSendCommand("", "request", "ssap://com.webos.service.tv.systemproperty", "", "", True)
                LGSendCommand("", "request", "ssap://com.webos.service.update/getCurrentSWInformation", "", "", True)
                'ssap://com.webos.service.update/getCurrentSWInformation
        End Select
        '   mainSocket.sendCommand("keyboard_", "subscribe", "ssap://com.webos.service.ime/registerRemoteKeyboard")
        '   mainSocket.sendCommand("events_", "subscribe", "ssap://com.webos.service.tv.keymanager/listInterestingEvents",
        '   mainSocket.sendCommand("foreground_app_", "subscribe", "ssap://com.webos.applicationManager/getForegroundAppInfo")



        Return True

    End Function

    Public Sub LGSendMessage(Message As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LGSendMessage called with Message = " & Message.ToString & " while service active = " & MyMessageServiceActive.ToString, LogType.LOG_TYPE_INFO)
        LGSendKeyCode("showfloat", Message)
    End Sub
End Class
