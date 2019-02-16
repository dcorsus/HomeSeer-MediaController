Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.IO
Imports System.Drawing

Partial Public Class HSPI

    Dim MyOnkyoAsyncSocket As AsynchronousClient
    Dim MyOnkyoClient As Socket
    Dim MyOnkyoPortNbr As Integer = 60128
    Dim ArtWork As MemoryStream = Nothing
    Dim ArtworkType As String = ""
    Dim OnkyoArt As Image = Nothing
    'Dim ArtFileIndex As Integer = 0

    Private Sub WriteOnkyoKeyInfoToInfoFile()
        If g_bDebug Then Log("WriteOnkyoKeyInfoToInfoFile called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = gRemoteControlPath
        'WriteStringIniFile(MyUDN & " - Default Codes", "PowerOn", "PWR01" & ":;:-:20", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "PowerStandby", "PWR00" & ":;:-:21", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "PowerQ", "PWRQSTN" & ":;:-:22", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Input Music Server", "SLI27" & ":;:-:22", objRemoteFile)
        ' added v24
        WriteStringIniFile(MyUDN, "20", "PowerOn" & ":;:-:" & "PWR01" & ":;:-:1:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "21", "PowerStandby" & ":;:-:" & "PWR00" & ":;:-:1:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "22", "Mute Off" & ":;:-:" & "AMT00" & ":;:-:1:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "23", "Mute On" & ":;:-:" & "AMT01" & ":;:-:1:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "24", "Volume Up" & ":;:-:" & "MVLUP" & ":;:-:1:;:-:5", objRemoteFile)
        WriteStringIniFile(MyUDN, "25", "Volume Down" & ":;:-:" & "MVLDOWN" & ":;:-:2:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "26", "Video1" & ":;:-:" & "SLI00" & ":;:-:2:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "27", "Video2" & ":;:-:" & "SLI01" & ":;:-:2:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "28", "Video3" & ":;:-:" & "SLI02" & ":;:-:2:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "29", "Video4" & ":;:-:" & "SLI03" & ":;:-:2:;:-:5", objRemoteFile)
        WriteStringIniFile(MyUDN, "30", "Video5" & ":;:-:" & "SLI04" & ":;:-:3:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "31", "Video6" & ":;:-:" & "SLI05" & ":;:-:3:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "32", "Video7" & ":;:-:" & "SLI06" & ":;:-:3:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "33", "DVD" & ":;:-:" & "SLI10" & ":;:-:3:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "34", "Tape1" & ":;:-:" & "SLI20" & ":;:-:3:;:-:5", objRemoteFile)
        WriteStringIniFile(MyUDN, "35", "Tape2" & ":;:-:" & "SLI21" & ":;:-:4:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "36", "Phono" & ":;:-:" & "SLI22" & ":;:-:4:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "37", "CD" & ":;:-:" & "SLI23" & ":;:-:4:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "38", "FM" & ":;:-:" & "SLI24" & ":;:-:4:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "39", "AM" & ":;:-:" & "SLI25" & ":;:-:5:;:-:5", objRemoteFile)
        WriteStringIniFile(MyUDN, "40", "Tuner" & ":;:-:" & "SLI26" & ":;:-:5:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "41", "Input Music Server" & ":;:-:" & "SLI27" & ":;:-:5:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "42", "Internet Radio" & ":;:-:" & "SLI28" & ":;:-:5:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "43", "USB/USB Front" & ":;:-:" & "SLI29" & ":;:-:5:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "44", "USB Rear" & ":;:-:" & "SLI2A" & ":;:-:5:;:-:5", objRemoteFile)
        WriteStringIniFile(MyUDN, "45", "Universal Port" & ":;:-:" & "SLI40" & ":;:-:6:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "46", "MultiCH" & ":;:-:" & "SLI30" & ":;:-:6:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "47", "XM" & ":;:-:" & "SLI31" & ":;:-:6:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "48", "Sirius" & ":;:-:" & "SLI32" & ":;:-:6:;:-:4", objRemoteFile)
    End Sub

    Private Sub CreateHSOnkyoRemoteButtons(ReCreate As Boolean)
        If g_bDebug Then Log("CreateHSOnkyoRemoteButtons called for device - " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
        HSRefRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Remote.ToString & "HSCode", -1)
        If HSRefRemote = -1 Then
            HSRefRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Remote.ToString)
        Else
            If Not ReCreate Then Exit Sub ' already exist
        End If

        If HSRefRemote = -1 Then Exit Sub ' problem!!
        If ReCreate Then
            hs.DeviceVSP_ClearAll(HSRefRemote, True)  ' added v24
            hs.DeviceVGP_ClearAll(HSRefRemote, True)  ' added v24
        End If
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRemoteOn
        Pair.Status = "Remote On"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psRemoteOff
        Pair.Status = "Remote Off"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 2
        hs.DeviceVSP_AddPair(HSRefRemote, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psCreateRemoteButtons
        Pair.Status = "Create Remote Buttons"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 3
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

        'CreateRemoteButtons(HSRefRemote)

    End Sub

    Private Sub TreatSetIOExOnkyo(ButtonValue As Integer)
        If g_bDebug Then Log("TreatSetIOExOnkyo called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            Case psRemoteOff  ' Remote Off
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, False)
                Try
                    OnkyoCloseTCPConnection()
                Catch ex As Exception
                    log("Error in TreatSetIOExOnkyo for UPnPDevice = " & MyUPnPDeviceName & "  setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                SetAdministrativeStateRemote(False)
            Case psRemoteOn  ' Remote On
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                If Not MyRemoteServiceActive Then
                    Try
                        OnkyoEstablishTCPConnection()
                        OnkyoGetBasicInfo()
                    Catch ex As Exception
                        log("Error in TreatSetIOExOnkyo for UPnPDevice = " & MyUPnPDeviceName & "  setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
                SetAdministrativeStateRemote(True)
            Case psCreateRemoteButtons
                CreateHSOnkyoRemoteButtons(True)
                CreateRemoteButtons(HSRefRemote)
            Case Else
                If g_bDebug Then Log("TreatSetIOExOnkyo called for UPnPDevice = " & MyUPnPDeviceName & " and Buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
                If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) And UCase(DeviceStatus) = "ONLINE" Then
                    Dim objRemoteFile As String = gRemoteControlPath
                    Dim ButtonInfoString As String = GetStringIniFile(MyUDN, ButtonValue.ToString, "", objRemoteFile)
                    Dim ButtonInfos As String()
                    ButtonInfos = Split(ButtonInfoString, ":;:-:")
                    If UBound(ButtonInfos, 1) > 2 Then
                        OnkyoSendKeyCode(ButtonInfos(1))
                    End If
                End If
        End Select


    End Sub

    Private Sub HandleOnkyoDataReceived(sender As Object, inData As Byte())
        If SuperDebug Then Log("HandleOnkyoDataReceived called for Device = " & MyUPnPDeviceName & " and Data = " & Encoding.UTF8.GetString(inData, 0, UBound(inData)), LogType.LOG_TYPE_INFO)
        '            					+0		    +1		    +2		    +3			
        '	eISCP Header				I		    S		    C		    P			
        '                                                                   Header(Size)
        '                                                                   Data(Size)
        '                               Version     (Reserved)
        '	eISCP Data				    1stChar		2ndChar		3rdChar		4thChar			
        '					            5thChar		ISCP Message							
        '                                                                   EndChar()
        If inData Is Nothing Then Exit Sub
        Dim Continue_ As Boolean = True
        Dim OnkyoData As String = ""
        Dim HeaderSize As Integer = 0
        Dim DataSize As Integer = 0
        Dim Version As Integer = 0
        Dim DataIndex As Integer = 0

        Do Until Not Continue_
            Try
                'Dim Headerstring As String = Encoding.ASCII.GetString(inData, 0, 4)
                If Encoding.UTF8.GetString(inData, DataIndex, 4) <> "ISCP" Then Exit Sub
                HeaderSize = inData(DataIndex + 7) ' BitConverter.ToInt32(inData, 4)
                DataSize = inData(DataIndex + 11) + (256 * inData(DataIndex + 10)) 'BitConverter.ToInt32(inData, 8)
                Version = inData(DataIndex + 12) ' Val(BitConverter.ToChar(inData, 12))
                If DataSize = 0 Then Exit Sub
                OnkyoData = Encoding.UTF8.GetString(inData, DataIndex + 16, DataSize - 1)
                If Trim(OnkyoData) = "" Then Exit Sub
                'If inData(DataIndex + HeaderSize + DataSize) <> 0 Then Exit Sub
            Catch ex As Exception
                If g_bDebug Then Log("Error in HandleOnkyoDataReceived called for Device = " & MyUPnPDeviceName & " with Data = " & Encoding.UTF8.GetString(inData, 0, UBound(inData)) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            ' OK all should be fine here!!

            ' !		1		P		W		R		0		1		[EOF]		[CR]		[LF]													
            '---   ---     -----------------------------------      -----------------------------  
            ' +		+						+							+->	End Character															
            ' +		+						+								"[EOF]" or "[EOF][CR]" or "[EOF][CR][LF]" depend on model															
            ' +     +                       +-> ISCP(Message(Command() + Parameter))
            ' +		+->		Destination Unit type Character ("1" for Receiver)																											
            ' +->           Start(Character)
            ProcessOnkyoResponse(OnkyoData)
            DataIndex = DataIndex + HeaderSize + DataSize
            If DataIndex >= UBound(inData) Then Exit Sub
        Loop

    End Sub

    Private Sub ProcessOnkyoResponse(ResponseData As String)
        'If g_bDebug Then Log("ProcessOnkyoResponse called for Device = " & MyUPnPDeviceName & " and ResponseData = " & ResponseData, LogType.LOG_TYPE_INFO)
        If ResponseData(0) <> "!" Then Exit Sub
        ResponseData = ResponseData.Remove(0, 2)
        ResponseData = ResponseData.Replace(vbCr, "")
        ResponseData = ResponseData.Replace(vbLf, "")
        ResponseData = ResponseData.Replace(vbNullChar, "")
        Dim Code As String = ResponseData.Substring(0, 3)
        Dim CodeData As String = ResponseData.Substring(3, ResponseData.Length - 3)
        Try
            Select Case Code
                Case "SLI"
                Case "AMT"
                Case "SPA"
                Case "SPB"
                Case "MVL"
                Case "SLA"
                Case "RES"
                Case "LMD"
                Case "TUN"
                Case "NST" ' Network/USB Play Status format 'NET/USB Play Status (3 letters) p -> Play Status: "S": STOP, "P": Play, "p": Pause, "F": FF, "R": FR r -> Repeat Status: "-": Off, "R": All, "F": Folder, "1": Repeat 1,s -> Shuffle Status: "-": Off, "S": All , "A": Album, "F": Folder
                    If CodeData.Length > 0 Then
                        Select Case CodeData(0)
                            Case "S"    ' Stopped
                                CurrentPlayerState = HomeSeerAPI.player_state_values.stopped
                            Case "P"    ' Playing
                                CurrentPlayerState = HomeSeerAPI.player_state_values.playing
                            Case "p"    ' Paused
                                CurrentPlayerState = HomeSeerAPI.player_state_values.paused
                            Case "F"    ' FF
                                CurrentPlayerState = HomeSeerAPI.player_state_values.forwarding
                            Case "R"    ' RR
                                CurrentPlayerState = HomeSeerAPI.player_state_values.rewinding
                        End Select
                    End If
                    If CodeData.Length > 1 Then
                        Select Case CodeData(1)
                            Case "-"    ' off
                                'Repeat = repeat_modes.repeat_off
                            Case "R"    ' Repeat ALL
                            Case "F"    ' Folder
                            Case "1"    ' Repeat 1
                        End Select
                    End If
                    If CodeData.Length > 2 Then
                        Select Case CodeData(2)
                            Case "-"    ' Shuffle off
                            Case "S"    ' All
                            Case "A"    ' Album 
                            Case "F"    ' Folder
                        End Select
                    End If
                Case "NTM" ' 00:16/04:10 this is position and length
                    Dim TimeInfo As String() = Split(CodeData.Substring(0, CodeData.Length - 1), "/")
                    If TimeInfo IsNot Nothing Then
                        If UBound(TimeInfo) > 0 Then
                            SetPlayerPosition(ConvertTimeFormatToSeconds(TimeInfo(0)))
                            SetTrackLength(ConvertTimeFormatToSeconds(TimeInfo(1)))
                        End If
                    End If
                Case "NJA" ' NET/USB Jacket Art/Album Art Data t-> Image type 0:BMP,1:JPEG p-> Packet flag 0:Start, 1:Next, 2:End xxxxxxxxxxxxxx -> Jacket/Album Art Data (variable length, 1024 ASCII HEX letters max)
                    If CodeData.Length < 2 Then Exit Select

                    If CodeData(1) = "0" Then
                        If CodeData(0) = "0" Then
                            ArtworkType = "bmp"
                        ElseIf CodeData(0) = "1" Then
                            ArtworkType = "jpg"
                        End If
                        If ArtWork IsNot Nothing Then
                            Try
                                ArtWork.Close()
                            Catch ex As Exception
                            End Try
                            Try
                                ArtWork.Dispose()
                            Catch ex As Exception
                            End Try
                            ArtWork = Nothing
                        End If
                        ArtWork = New MemoryStream
                        If ArtWork Is Nothing Then Exit Select
                        'ArtWork.Capacity = 1000000
                        ' calculate the length of the byte array and dim an array to that
                        Dim nBytes = (CodeData.Length - 3) \ 2
                        If nBytes > 0 Then
                            Dim ArtBytes(nBytes - 1) As Byte
                            ' pick out every two bytes and convert them from hex representation
                            For i = 0 To nBytes - 1
                                ArtBytes(i) = Convert.ToByte(CodeData.Substring((i * 2) + 2, 2), 16)
                            Next
                            ArtWork.Write(ArtBytes, 0, nBytes)
                        End If
                    ElseIf CodeData(1) = "1" Then
                        Dim nBytes = (CodeData.Length - 3) \ 2
                        If nBytes > 0 Then
                            Dim ArtBytes(nBytes - 1) As Byte
                            ' pick out every two bytes and convert them from hex representation
                            For i = 0 To nBytes - 1
                                ArtBytes(i) = Convert.ToByte(CodeData.Substring((i * 2) + 2, 2), 16)
                            Next
                            ArtWork.Write(ArtBytes, 0, nBytes)
                        End If
                    ElseIf CodeData(1) = "2" Then
                        Dim nBytes = (CodeData.Length - 3) \ 2
                        If nBytes > 0 Then
                            Dim ArtBytes(nBytes - 1) As Byte
                            ' pick out every two bytes and convert them from hex representation
                            For i = 0 To nBytes - 1
                                ArtBytes(i) = Convert.ToByte(CodeData.Substring((i * 2) + 2, 2), 16)
                            Next
                            ArtWork.Write(ArtBytes, 0, nBytes)
                        End If
                        OnkyoArt = Image.FromStream(ArtWork, True, True)
                        ArtworkURL = SaveArtwork(OnkyoArt, "." & ArtworkType, False)
                        Try
                            OnkyoArt.Dispose()
                        Catch ex As Exception
                        End Try
                        OnkyoArt = Nothing
                        Try
                            ArtWork.Close()
                        Catch ex As Exception
                        End Try
                        Try
                            ArtWork.Dispose()
                        Catch ex As Exception
                        End Try
                        ArtWork = Nothing
                    End If
                Case "NLT" ' First byte is: the network source type Second byte is: menu depth (how far you dug down) 3rd,4th byte: selected item from list 5th, 6th: total items in list 2nd to last byte: network icon for net GUI Last byte: always 00 Any characters following that byte are the title for the net GUI
                Case "NAT"
                    Artist = CodeData
                Case "NAL"
                    Album = CodeData
                Case "NTI"
                    Track = CodeData
                Case "NTR"
                    Dim TrackInfo As String() = Split(CodeData.Substring(0, CodeData.Length - 1), "/") ' presented as cccc/tttt with cccc = Current Track Nbr and tttt = total tracks
                    If TrackInfo IsNot Nothing Then
                        If UBound(TrackInfo) > 0 Then
                            MyCurrentTrackNumber = Val(TrackInfo(0)) ' if not present this will be ---
                            MyCurrentNrTracks = Val(TrackInfo(1))
                        End If
                    End If
            End Select
        Catch ex As Exception
            If g_bDebug Then Log("Error in ProcessOnkyoResponse called for Device = " & MyUPnPDeviceName & " and Code = " & Code & " with Data = " & CodeData & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        If SuperDebug Then Log("ProcessOnkyoResponse called for Device = " & MyUPnPDeviceName & " and Code = " & Code & " with Data = " & CodeData, LogType.LOG_TYPE_INFO)
        'If g_bDebug And Code <> "NJA" And Not SuperDebug Then Log("ProcessOnkyoResponse called for Device = " & MyUPnPDeviceName & " and Code = " & Code & " with Data = " & CodeData, LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub OnkyoEstablishTCPConnection()

        If MyOnkyoAsyncSocket Is Nothing Then
            Try
                MyOnkyoAsyncSocket = New AsynchronousClient
            Catch ex As Exception
                Log("Error in OnkyoEstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to open Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyRemoteServiceActive = False
                Exit Sub
            End Try
        End If

        AddHandler MyOnkyoAsyncSocket.DataReceived, AddressOf HandleOnkyoDataReceived

        Try
            MyOnkyoClient = MyOnkyoAsyncSocket.ConnectSocket(MyIPAddress, MyOnkyoPortNbr.ToString)
        Catch ex As Exception
            log("Error in OnkyoEstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to open Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyRemoteServiceActive = False
            Exit Sub
        End Try

        If MyOnkyoClient Is Nothing Then
            Log("Error in OnkyoEstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " . Unable to open Socket", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim WaitForConnection As Integer = 0
        Do While WaitForConnection < 10
            If MyOnkyoAsyncSocket.MySocketIsClosed Then
                wait(1)
                WaitForConnection = WaitForConnection + 1
            Else
                Exit Do
            End If
        Loop

        If WaitForConnection >= 10 Then
            ' unsuccesfull connection
            Log("Error in OnkyoEstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " . Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
            MyOnkyoAsyncSocket.CloseSocket()
            Exit Sub
        End If

        MyRemoteServiceActive = True
        Try
            If g_bDebug Then MyOnkyoAsyncSocket.Receive()
        Catch ex As Exception
            log("Error in OnkyoEstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyOnkyoAsyncSocket.response = False
        ' no need for these here
        'StopAVPollTimer()      ' removed 12/2/2018 in v.039
        'StopRenderPollTimer()  ' removed 12/2/2018 in v0.39
        MyTimeoutActionArray(TOPositionUpdate) = 0

    End Sub

    Private Sub OnkyoCloseTCPConnection()
        If Not MyRemoteServiceActive Then Exit Sub
        Try
            RemoveHandler MyOnkyoAsyncSocket.DataReceived, AddressOf HandleOnkyoDataReceived
        Catch ex As Exception
        End Try
        Try
            MyOnkyoAsyncSocket.CloseSocket()
        Catch ex As Exception
            log("Error in OnkyoCloseTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MyOnkyoClient = Nothing
            MyOnkyoAsyncSocket = Nothing
        Catch ex As Exception
            log("Error in OnkyoCloseTCPConnection 1 for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If HSRefRemote <> -1 Then hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
            MyRemoteServiceActive = False
        Catch ex As Exception
            log("Error in OnkyoCloseTCPConnection 2 for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyTimeoutActionArray(TOPositionUpdate) = TOPositionUpdateValue
    End Sub

    Public Sub OnkyoSendKeyCode(KeyCode As String)
        If g_bDebug Then Log("OnkyoSendKeyCode was called for UPnPDevice = " & MyUPnPDeviceName & " with key = " & KeyCode.ToString, LogType.LOG_TYPE_INFO)
        Dim ReturnString As String = ""

        If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
            If MyOnkyoClient Is Nothing Then
                OnkyoEstablishTCPConnection()
            End If
        Else
            Exit Sub
        End If
        If MyOnkyoClient Is Nothing Then
            Log("Error in OnkyoSendKeyCode. There is no socket for UPnPDevice =  " & MyUPnPDeviceName, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        'Dim sendBytes(KeyCode.Length + 18) As Char
        Dim SendBytes As String = "ISCP" & Chr(0) & Chr(0) & Chr(0) & Chr(16) & Chr(0) & Chr(0) & Chr(0) & Chr(KeyCode.Length + 3) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & "!1"
        'sendBytes(0) = "I"
        'sendBytes(1) = "S"
        'sendBytes(2) = "C"
        'sendBytes(3) = "P"
        'sendBytes(4) = Chr(0)
        'sendBytes(5) = Chr(0)
        'sendBytes(6) = Chr(0)
        'sendBytes(7) = Chr(16)
        'sendBytes(8) = Chr(0)
        'sendBytes(9) = Chr(0)
        'sendBytes(10) = Chr(0)
        'sendBytes(11) = Chr(KeyCode.Length + 3)
        'sendBytes(12) = Chr(1)
        'sendBytes(13) = Chr(0)
        'sendBytes(14) = Chr(0)
        'sendBytes(15) = Chr(0)
        'sendBytes(16) = "!"  'Chr(33) 
        'sendBytes(17) = "1"   'Chr(49)  ' 1 is for the Receiver
        'Dim i As Int32
        'For i = 0 To (KeyCode.Length - 1)
        'SendBytes(18 + i) = KeyCode.Chars(i)
        'Next
        'SendBytes(KeyCode.Length + 18) = Chr(13)
        SendBytes = SendBytes & KeyCode & Chr(13) 'Chr(26) & Chr(13) & Chr(10)
        Dim outBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(SendBytes)
        Try
            MyOnkyoAsyncSocket.response = False
            If Not MyOnkyoAsyncSocket.Send(outBytes) Then
                Log("Error in OnkyoSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code", LogType.LOG_TYPE_ERROR)
                OnkyoCloseTCPConnection()
                outBytes = Nothing
                Exit Sub
            End If
        Catch ex As Exception
            log("Error in OnkyoSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            OnkyoCloseTCPConnection()
            outBytes = Nothing
            Exit Sub
        End Try
        outBytes = Nothing
        Try
            MyOnkyoAsyncSocket.sendDone.WaitOne()
        Catch ex As Exception
            log("Error in OnkyoSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            OnkyoCloseTCPConnection()
            Exit Sub
        End Try

    End Sub

    Private Sub OnkyoGetBasicInfo()
        OnkyoSendKeyCode("PWRQSTN")
        OnkyoSendKeyCode("SLIQSTN")
        OnkyoSendKeyCode("AMTQSTN")
        OnkyoSendKeyCode("MVLQSTN")
        OnkyoSendKeyCode("SLAQSTN")
        OnkyoSendKeyCode("RESQSTN")
        OnkyoSendKeyCode("LMDQSTN")
        OnkyoSendKeyCode("TUNQSTN")
        OnkyoSendKeyCode("XCNQSTN")
        OnkyoSendKeyCode("SCNSTN")
    End Sub

End Class