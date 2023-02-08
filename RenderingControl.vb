Imports System.Xml

Partial Public Class HSPI

    Private RenderingControl As MyUPnPService = Nothing
    Public WithEvents myRenderingControlCallback As New myUPnPControlCallback
    Private MyRendererHSRef As Integer = -1


    Private BrightnessIsConfigurable As Boolean = False
    Private ColorTemperatureIsConfigurable As Boolean = False
    Private ContrastIsConfigurable As Boolean = False
    Public MuteIsConfigurable As Boolean = False
    Private SharpnessIsConfigurable As Boolean = False
    Public VolumeIsConfigurable As Boolean = False
    Public LoudnessIsConfigurable As Boolean = False
    Private VolumeDBIsConfigurable As Boolean = False
    Private GetSlideShowEffectIsConfigurable As Boolean = False
    Private GetImageScaleIsConfigurable As Boolean = False
    Private GetImageRotationIsConfigurable As Boolean = False
    Public SpeedIsConfigurable As Boolean = False
    Private BalanceIsConfigurable As Boolean = False
    Public MyMinimumVolume As Integer = 0
    Public MyMaximumVolume As Integer = 100

    Private MyCurrentVolumeLevel As Integer = 0
    Private MyCurrentMuteState As Boolean = False
    Private MyCurrentVolumeDBLevel As Integer = 0
    Private MyCurrentVolumeBakLevel As Integer = 0
    Private MyCurrentLFVolumeLevel As Integer = 0
    Private MyBalance As Integer = 0
    Private MyCurrentLFVolumeDBLevel As Integer = 0
    Private MyCurrentLFVolumeBakLevel As Integer = 0
    Private MyCurrentRFVolumeLevel As Integer = 0
    Private MyCurrentRFVolumeDBLevel As Integer = 0
    Private MyCurrentRFVolumeBakLevel As Integer = 0
    Private MyCurrentLFMuteState As Boolean = False
    Private MyCurrentRFMuteState As Boolean = False
    Private MyCurrentPresetNameList As String = ""
    Private MyCurrentPresetName As String = ""
    Private MyCurrentMinValue As Integer = 0
    Private MyCurrentMaxValue As Integer = 0
    Private MyCurrentLoudness As Boolean = False
    Private MyCurrentBrightness As Integer = 0
    Private MyCurrentContrast As Integer = 0
    Private MyCurrentSharpness As Integer = 0
    Private MyCurrentColorTemperature As Integer = 0
    Private MyCurrentSlideShowEffect As String = ""
    Private MyCurrentImageScale As Integer = 0
    Private MyCurrentImageRotation As Integer = 0
    Private MyCurrentTreble As Integer = 0
    Private MyCurrentBass As Integer = 0
    Private MyMuteHSRef As Integer = -1
    Private MyVolumeHSRef As Integer = -1
    Private MyVolumeDBHSRef As Integer = -1
    Private MyVolumeBakHSRef As Integer = -1
    Private MyLoudnessHSRef As Integer = -1
    Private MyBrightnessHSRef As Integer = -1
    Private MyColorHSRef As Integer = -1
    Private MyContrastHSRef As Integer = -1
    Private MySharpnessHSRef As Integer = -1
    Private MySlideshowEffectHSRef As Integer = -1
    Private MyImageScaleHSRef As Integer = -1
    Private MyImageRotationHSRef As Integer = -1
    Private VolumeChanged As Boolean = False
    Private VolumeDBChanged As Boolean = False
    Private MyPollRenderStateReEntry As Boolean = False



    Private Sub RenderingControlStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myRenderingControlCallback.ControlStateChange 'RenderingControlStateChange

        If PIDebuglevel > DebugLevel.dlEvents Then Log("RenderingControlStateChange for device = " & MyUPnPDeviceName & ". Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly And Not PIDebuglevel > DebugLevel.dlEvents Then Log("RenderingControlStateChange for device = " & MyUPnPDeviceName & ". Var Name = " & StateVarName, LogType.LOG_TYPE_INFO)
        Dim xmlData As XmlDocument = New XmlDocument
        VolumeChanged = False
        VolumeDBChanged = False
        If StateVarName = "LastChange" Then
            Try
                If Trim(Value.ToString) = "" Then
                    Exit Sub
                End If
                Try
                    xmlData.LoadXml(Value.ToString)
                Catch ex As Exception
                    Log("Error in RenderingControlStateChange for " & MyUPnPDeviceName & " loading XML. XML = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
                Try
                    'Get a list of all the child elements
                    Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                    'If piDebuglevel > DebugLevel.dlEvents Then log( "TransportStateChange Nbr of items in XML Data = " & nodelist.Count) ' this starts with <Event>
                    'If piDebuglevel > DebugLevel.dlEvents Then log( "TransportStateChange Document root node: " & xmlData.DocumentElement.Name)
                    'Parse through all nodes
                    For Each outerNode As XmlNode In nodelist
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("RenderingControlStateChange Outer node name: " & outerNode.Name & " and ID = " & outerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO) ' this will be InstanceID
                        'Check if this matches with our selected item
                        For Each InnerNode As XmlNode In outerNode.ChildNodes
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("RenderingControlStateChange------> Inner node Name: " & InnerNode.Name, LogType.LOG_TYPE_INFO)
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("RenderingControlStateChange------> Inner node Value: " & InnerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO) ' Here are the Values
                            Dim ChannelName As String = ""
                            Try
                                ChannelName = InnerNode.Attributes("channel").Value
                            Catch ex As Exception
                            End Try
                            If ChannelName = "" Then
                                Try
                                    ChannelName = InnerNode.Attributes("Channel").Value
                                Catch ex As Exception
                                End Try
                            End If
                            ProcessRenderingControlXML(InnerNode.Name, ChannelName, InnerNode.Attributes("val").Value)
                        Next
                    Next
                Catch ex As Exception
                    Log("Error in RenderingControlStateChange for UPnPDevice = " & MyUPnPDeviceName & "  processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Catch ex As Exception
                Log("Error: This rendering didn't work too well for zoneplayer = " & MyUPnPDeviceName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RenderingControlStateChange callback for device = " & MyUPnPDeviceName & " received an unknown Variable = " & StateVarName & " and data = " & Value.ToString, LogType.LOG_TYPE_WARNING)
        End If
    End Sub

    Private Sub RenderingControlDied() Handles myRenderingControlCallback.ControlDied 'RenderingControlDied
        'Log( "Rendering Control Callback Died for device = " & MyUPnPDeviceName,LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to device " & MyUPnPDeviceName & " was lost in RenderingControlDied.", LogType.LOG_TYPE_WARNING)
            Disconnect(False)
        Catch ex As Exception
            Log("Error in RenderingControlDied for device = " & MyUPnPDeviceName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub InformHSVolumeChange(NewValue As Integer)
        If MyCurrentVolumeLevel <> NewValue Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("InformHSVolumeChange called for device = " & MyUPnPDeviceName & " with NewValue = " & NewValue.ToString & " and CurrentValue = " & MyCurrentVolumeLevel.ToString, LogType.LOG_TYPE_INFO)
            If MyCurrentVolumeLevel > NewValue Then
                DeviceTrigger("Volume Down")
            ElseIf MyCurrentVolumeLevel < NewValue Then
                DeviceTrigger("Volume Up")
            End If
            SetVolume = NewValue
            PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False) ' this will update iPad clients
        End If
    End Sub

    Private Sub ProcessRenderingControlXML(VariableName As String, VariableChannel As String, VariableValue As String)
        ' Samsung
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/">
        '   <InstanceID val="0">
        '       <PresetNameList val="FactoryDefaults"/>
        '       <Brightness val="45"/>
        '       <Contrast val="100"/>
        '       <Sharpness val="55"/>
        '       <ColorTemperature val="2"/>
        '       <Mute channel="Master" val="0"/>
        '       <Volume channel="Master" val="29"/>
        '       <X_SlideShowEffect val="DEFAULT"/>
        '       <X_ImageScale val="1"/>
        '       <X_ImageRotation val="0"/>
        '   </InstanceID>
        '</Event>
        ' Onkyo
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/">
        '   <InstanceID val="0">
        '       <PresetNameList val="FactoryDefaults,InstallationDefaults"/>
        '       <Mute channel="Master" val="0"/>
        '       <Volume channel="Master" val="29"/>
        '       <VolumeBak channel="Master" val="0"/>
        '       <VolumeDB channel="Master" val="0"/>
        '       <Loudness channel="Master" val="0"/>
        '       <Mute channel="LF" val="0"/>
        '       <Volume channel="LF" val="0"/>
        '       <VolumeBak channel="LF" val="0"/>
        '       <VolumeDB channel="LF" val="0"/>
        '       <Loudness channel="LF" val="0"/>
        '       <Mute channel="RF" val="0"/>
        '       <Volume channel="RF" val="0"/>
        '       <VolumeBak channel="RF" val="0"/>
        '       <VolumeDB channel="RF" val="0"/>
        '       <Loudness channel="RF" val="0"/>
        '   </InstanceID>
        '</Event>

        ' Windows Media Manager
        ' <Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/">
        '   <InstanceID val="0">
        '       <Volume channel="Master" val="31"/>
        '    </InstanceID>
        '</Event>

        ' Sonos
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/">
        '   <InstanceID val="0">
        '       <Volume channel="Master" val="23"/>
        '       <Volume channel="LF" val="100"/>
        '       <Volume channel="RF" val="100"/>
        '       <Mute channel="Master" val="0"/>
        '       <Mute channel="LF" val="0"/>
        '       <Mute channel="RF" val="0"/>
        '       <Bass val="10"/><Treble val="6"/>
        '       <Loudness channel="Master" val="1"/>
        '       <OutputFixed val="0"/>
        '       <HeadphoneConnected val="0"/>
        '       <PresetNameList>FactoryDefaults</PresetNameList>
        '   </InstanceID>
        '</Event>

        Select Case VariableName
            Case "Volume"
                If VariableChannel = "Master" Or VariableChannel = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Master Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue <> "" Then InformHSVolumeChange(Val(VariableValue))
                ElseIf VariableChannel = "LF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (LF Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentLFVolumeLevel <> Val(VariableValue) Then
                        VolumeChanged = True
                        MyCurrentLFVolumeLevel = Val(VariableValue)
                        UpdateBalance()
                    End If
                ElseIf VariableChannel = "RF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (RF Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentRFVolumeLevel <> Val(VariableValue) Then
                        VolumeChanged = True
                        MyCurrentRFVolumeLevel = Val(VariableValue)
                        UpdateBalance()
                    End If
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_WARNING)
                End If
            Case "Mute"
                If VariableChannel = "Master" Or VariableChannel = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Mute Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentMuteState <> VariableValue Then
                        PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
                    End If
                    If VariableValue = False Then SetMuteState = False Else SetMuteState = True
                ElseIf VariableChannel = "LF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Mute LF Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If VariableValue = False Then MyCurrentLFMuteState = False Else MyCurrentLFMuteState = True
                ElseIf VariableChannel = "RF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Mute RF Volume) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If VariableValue = False Then MyCurrentRFMuteState = False Else MyCurrentRFMuteState = True
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " Mute) = " & VariableValue.ToString, LogType.LOG_TYPE_WARNING)
                End If
            Case "Treble"
                If VariableChannel = "Master" Or VariableChannel = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Treble) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    MyCurrentTreble = Val(VariableValue)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " Treble) = " & VariableValue.ToString, LogType.LOG_TYPE_WARNING)
                End If
            Case "Bass"
                If VariableChannel = "Master" Or VariableChannel = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Bass) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    MyCurrentBass = Val(VariableValue)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " Bass) = " & VariableValue.ToString, LogType.LOG_TYPE_WARNING)
                End If
            Case "Loudness"
                If VariableChannel = "Master" Or VariableChannel = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Loudness) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentLoudness <> VariableValue Then
                        SetLoudness = VariableValue
                        PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
                    End If
                ElseIf VariableChannel = "LF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Loudness LF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                ElseIf VariableChannel = "RF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (Loudness RF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " Loudness) = " & VariableValue, LogType.LOG_TYPE_WARNING)
                End If
            Case "VolumeDB"
                If VariableChannel = "Master" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeDB) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentVolumeDBLevel <> Val(VariableValue) Then
                        VolumeDBChanged = True
                        MyCurrentVolumeDBLevel = Val(VariableValue)
                    End If
                ElseIf VariableChannel = "LF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeDB LF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentLFVolumeDBLevel <> Val(VariableValue) Then
                        VolumeDBChanged = True
                        MyCurrentLFVolumeDBLevel = Val(VariableValue)
                    End If
                ElseIf VariableChannel = "RF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeDB RF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    If MyCurrentRFVolumeDBLevel <> Val(VariableValue) Then
                        VolumeDBChanged = True
                        MyCurrentRFVolumeDBLevel = Val(VariableValue)
                    End If
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " VolumeDB) = " & VariableValue, LogType.LOG_TYPE_WARNING)
                End If
            Case "VolumeBak"
                If VariableChannel = "Master" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeBak) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    MyCurrentVolumeBakLevel = Val(VariableValue)
                ElseIf VariableChannel = "LF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeBak LF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    MyCurrentLFVolumeBakLevel = Val(VariableValue)
                ElseIf VariableChannel = "RF" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VolumeBak RF) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                    If VariableValue = "" Then Exit Select
                    MyCurrentRFVolumeBakLevel = Val(VariableValue)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " : Rendering (" & VariableChannel & " VolumeBak) = " & VariableValue, LogType.LOG_TYPE_WARNING)
                End If
            Case "PresetNameList"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : PresetNameList = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                MyCurrentPresetNameList = VariableValue
            Case "Brightness"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Brightness = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentBrightness = Val(VariableValue)
                hs.SetDeviceString(MyBrightnessHSRef, "Brightness = " & MyCurrentBrightness.ToString, True)
                hs.SetDeviceValueByRef(MyBrightnessHSRef, MyCurrentBrightness, True)
            Case "Contrast"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Contrast = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentContrast = Val(VariableValue)
                hs.SetDeviceString(MyContrastHSRef, "Contrast = " & MyCurrentContrast.ToString, True)
                hs.SetDeviceValueByRef(MyContrastHSRef, MyCurrentContrast, True)
            Case "Sharpness"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Sharpness = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentSharpness = Val(VariableValue)
                hs.SetDeviceString(MySharpnessHSRef, "Sharpness = " & MyCurrentSharpness.ToString, True)
                hs.SetDeviceValueByRef(MySharpnessHSRef, MyCurrentSharpness, True)
            Case "ColorTemperature"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : ColorTemperature = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentColorTemperature = Val(VariableValue)
                hs.SetDeviceString(MyColorHSRef, "Color = " & MyCurrentColorTemperature.ToString, True)
                hs.SetDeviceValueByRef(MyColorHSRef, MyCurrentColorTemperature, True)
            Case "X_SlideShowEffect"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : X_SlideShowEffect = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                MyCurrentSlideShowEffect = VariableValue
                hs.SetDeviceString(MySlideshowEffectHSRef, "Slideshow effect = " & MyCurrentSlideShowEffect.ToString, True)
                'hs.SetDeviceValue(MySlideshowEffectHSCode, MyCurrentSlideShowEffect)
            Case "X_ImageScale"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : X_ImageScale = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentImageScale = Val(VariableValue)
                'hs.SetDeviceString(MyImageScaleHSCode, "Image Scale = " & MyCurrentImageScale.ToString)
                'hs.SetDeviceValue(MyImageScaleHSCode, MyCurrentImageScale)
                'log( "ProcessRenderingControlXML for device - " & MyUPnPDeviceName & " just set Image Scale with value = " & MyCurrentImageScale.ToString & " for HSDevice = " & MyImageScaleHSCode)
            Case "X_ImageRotation"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : X_ImageRotation = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                If VariableValue = "" Then Exit Select
                MyCurrentImageRotation = Val(VariableValue)
                hs.SetDeviceString(MyImageRotationHSRef, "Image Rotation = " & MyCurrentImageRotation.ToString, True)
                'hs.SetDeviceValue(MyImageRotationHSCode, MyCurrentImageRotation)
            Case "SubEnabled"   ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (SubEnabled) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SubPolarity"  ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (SubPolarity) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SubCrossover" ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (SubCrossover) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SubGain"      ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (SubGain) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SpeakerSize"  ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (SpeakerSize) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "HeadphoneConnected"   ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (HeadphoneConnected) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "OutputFixed"  ' Sonos
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (OutputFixed) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "LastChange"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (LastChange) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "RedVideoGain" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (RedVideoGain) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "GreenVideoGain" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (GreenVideoGain) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "BlueVideoGain" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (BlueVideoGain) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "RedVideoBlackLevel" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (RedVideoBlackLevel) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "GreenVideoBlackLevel" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (GreenVideoBlackLevel) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "BlueVideoBlackLevel" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (BlueVideoBlackLevel) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "HorizontalKeystone" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (HorizontalKeystone) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "VerticalKeystone", "X_Subtitle", "X_Current3DFormatter", "X_Possible3DFormatter" ' LG
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " : Rendering (VerticalKeystone) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)

            Case Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ProcessRenderingControlXML for Device - " & MyUPnPDeviceName & " received untreated (" & VariableName & ") = " & VariableValue, LogType.LOG_TYPE_WARNING)
        End Select
    End Sub

    Public Sub VolumeUP()
        Dim MyNewVolumeLevel As Integer
        MyNewVolumeLevel = MyCurrentVolumeLevel + MyVolumeStep
        If MyNewVolumeLevel > MyMaximumVolume Then
            MyNewVolumeLevel = MyMaximumVolume
            Exit Sub
        End If
        RCSetVolume(MyNewVolumeLevel)
    End Sub

    Public Sub VolumeDown()
        Dim MyNewVolumeLevel As Integer
        MyNewVolumeLevel = MyCurrentVolumeLevel - MyVolumeStep
        If MyNewVolumeLevel < MyMinimumVolume Then
            MyNewVolumeLevel = MyMinimumVolume
            Exit Sub
        End If
        RCSetVolume(MyNewVolumeLevel)
    End Sub

    Public Sub VolumeDBUp()
        Dim MyNewVolumeDBLevel As Integer
        MyNewVolumeDBLevel = MyCurrentVolumeDBLevel + MyVolumeStep
        If MyNewVolumeDBLevel > MyMaximumVolume Then
            MyNewVolumeDBLevel = MyMaximumVolume
            Exit Sub
        End If
        RCSetVolumeDB(MyNewVolumeDBLevel)
    End Sub

    Public Sub VolumeDBDown()
        Dim MyNewVolumeDBLevel As Integer
        MyNewVolumeDBLevel = MyCurrentVolumeDBLevel - MyVolumeStep
        If MyNewVolumeDBLevel < 0 Then
            MyNewVolumeDBLevel = 0
            Exit Sub
        End If
        RCSetVolumeDB(MyNewVolumeDBLevel)
    End Sub

    Public Sub MuteOn()
        RCSetMute(True)
    End Sub

    Public Sub MuteOff()
        RCSetMute(False)
    End Sub

    Public Sub MuteToggle()
        If MyCurrentMuteState Then
            RCSetMute(False)
        Else
            RCSetMute(True)
        End If
    End Sub

    Public Sub LoudnessOn()
        RCSetLoudness(True)
    End Sub

    Public Sub LoudnessOff()
        RCSetLoudness(False)
    End Sub

    Public Sub LoudnessToggle()
        If MyCurrentLoudness Then
            RCSetLoudness(False)
        Else
            RCSetLoudness(True)
        End If
    End Sub

    Public Sub BrightnessUp()
        Dim MyNewBrightness As Integer
        MyNewBrightness = MyCurrentBrightness + 1
        If MyNewBrightness > 100 Then
            MyNewBrightness = 100
            Exit Sub
        End If
        RCSetBrightness(MyNewBrightness)
    End Sub

    Public Sub BrightnessDown()
        Dim MyNewBrightness As Integer
        MyNewBrightness = MyCurrentBrightness - 1
        If MyNewBrightness < 0 Then
            MyNewBrightness = 0
            Exit Sub
        End If
        RCSetBrightness(MyNewBrightness)
    End Sub

    Public Sub ColorUp()
        Dim MyNewColor As Integer
        MyNewColor = MyCurrentColorTemperature + 1
        If MyNewColor > 100 Then
            MyNewColor = 100
            Exit Sub
        End If
        RCSetColorTemperature(MyNewColor)
    End Sub

    Public Sub ColorDown()
        Dim MyNewColor As Integer
        MyNewColor = MyCurrentColorTemperature - 1
        If MyNewColor < 0 Then
            MyNewColor = 0
            Exit Sub
        End If
        RCSetColorTemperature(MyNewColor)
    End Sub

    Public Sub ContrastUp()
        Dim MyNewContrast As Integer
        MyNewContrast = MyCurrentContrast + 1
        If MyNewContrast > 100 Then
            MyNewContrast = 100
            Exit Sub
        End If
        RCSetContrast(MyNewContrast)
    End Sub

    Public Sub ContrastDown()
        Dim MyNewContrast As Integer
        MyNewContrast = MyCurrentContrast - 1
        If MyNewContrast < 0 Then
            MyNewContrast = 0
            Exit Sub
        End If
        RCSetContrast(MyNewContrast)
    End Sub

    Public Sub SharpnessUp()
        Dim MyNewSharpness As Integer
        MyNewSharpness = MyCurrentSharpness + 1
        If MyNewSharpness > 100 Then
            MyNewSharpness = 100
            Exit Sub
        End If
        RCSetSharpness(MyNewSharpness)
    End Sub

    Public Sub SharpnessDown()
        Dim MyNewSharpness As Integer
        MyNewSharpness = MyCurrentSharpness - 1
        If MyNewSharpness < 0 Then
            MyNewSharpness = 0
            Exit Sub
        End If
        RCSetSharpness(MyNewSharpness)
    End Sub

    Public Sub ImageScaleUp()
        Dim MyNewImageScale As Integer
        MyNewImageScale = MyCurrentImageScale + 1
        If MyNewImageScale > 100 Then
            MyNewImageScale = 100
            Exit Sub
        End If
        RCX_SetImageScale(MyNewImageScale)
    End Sub

    Public Sub ImageScaleDown()
        Dim MyNewImageScale As Integer
        MyNewImageScale = MyCurrentImageScale - 1
        If MyNewImageScale < 0 Then
            MyNewImageScale = 0
            Exit Sub
        End If
        RCSetSharpness(MyNewImageScale)
    End Sub

    Public Sub ImageRotationUp()
        Dim MyNewImageRotation As Integer
        MyNewImageRotation = MyCurrentImageRotation + 1
        If MyNewImageRotation > 100 Then
            MyNewImageRotation = 100
            Exit Sub
        End If
        RCX_SetImageRotation(MyNewImageRotation)
    End Sub

    Public Sub ImageRotationDown()
        Dim MyNewImageRotation As Integer
        MyNewImageRotation = MyCurrentImageRotation - 1
        If MyNewImageRotation < 0 Then
            MyNewImageRotation = 0
            Exit Sub
        End If
        RCX_SetImageRotation(MyNewImageRotation)
    End Sub

    Public Function RCListPresets(Optional InstanceID As Integer = 0) As String
        RCListPresets = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCListPresets called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("ListPresets", InArg, OutArg)
            MyCurrentPresetNameList = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCListPresets : MyCurrentPresetNameList = " & MyCurrentPresetNameList.ToString, LogType.LOG_TYPE_INFO)
            End If
            RCListPresets = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCListPresets for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSelectPreset(Optional InstanceID As Integer = 0) As String
        RCSelectPreset = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCListPresets called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("SelectPreset", InArg, OutArg)
            MyCurrentPresetName = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCSelectPreset : MyPresetName = " & MyCurrentPresetName.ToString, LogType.LOG_TYPE_INFO)
            End If
            RCSelectPreset = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSelectPreset for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetMute(Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCGetMute = ""
        If DeviceStatus = "Offline" Then Exit Function
        If RenderingControl Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlEvents Then Log("RCGetMute called for device " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetMute", InArg, OutArg)
            SetMuteState = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetMute : MyCurrentMute = " & MyCurrentMuteState.ToString, LogType.LOG_TYPE_INFO)
            End If
            MuteIsConfigurable = True
            RCGetMute = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetMute for device = " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetMute(DesiredMute As Boolean, Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCSetMute = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetMute called for device " & MyUPnPDeviceName & " with DesiredMute = " & DesiredMute.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            InArg(2) = DesiredMute
            RenderingControl.InvokeAction("SetMute", InArg, OutArg)
            RCSetMute = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetMute for device = " & MyUPnPDeviceName & " with DesiredMute = " & DesiredMute.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetVolume(Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCGetVolume = ""
        If DeviceStatus = "Offline" Then Exit Function
        If RenderingControl Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlEvents Then Log("RCGetVolume called for device " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetVolume", InArg, OutArg)
            InformHSVolumeChange(OutArg(0))
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetVolume for device " & MyUPnPDeviceName & " : MyCurrentVolume = " & MyCurrentVolumeLevel.ToString, LogType.LOG_TYPE_INFO)
            End If
            VolumeIsConfigurable = True
            RCGetVolume = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetVolume for device = " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetVolume(DesiredVolume As Integer, Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCSetVolume = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetVolume called for device " & MyUPnPDeviceName & " with DesiredVolume = " & DesiredVolume.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            InArg(2) = DesiredVolume
            RenderingControl.InvokeAction("SetVolume", InArg, OutArg)
            RCSetVolume = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetVolume for device = " & MyUPnPDeviceName & " with DesiredVolume = " & DesiredVolume.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetVolumeDB(Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCGetVolumeDB = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetVolumeDB called for device " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetVolumeDB", InArg, OutArg)
            MyCurrentVolumeDBLevel = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetVolumeDB : MyCurrentVolumeDB = " & MyCurrentVolumeDBLevel.ToString, LogType.LOG_TYPE_INFO)
            End If
            VolumeDBIsConfigurable = True
            RCGetVolumeDB = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetVolumeDB for device = " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetVolumeDB(DesiredVolume As Integer, Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCSetVolumeDB = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetVolumeDB called for device " & MyUPnPDeviceName & " with DesiredVolumeDB = " & DesiredVolume.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            InArg(2) = DesiredVolume
            RenderingControl.InvokeAction("SetVolumeDB", InArg, OutArg)
            RCSetVolumeDB = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetVolumeDB for device = " & MyUPnPDeviceName & " with DesiredVolumeDB = " & DesiredVolume.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetVolumeDBRange(Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCGetVolumeDBRange = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetVolumeDBRange called for device " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(1)
            InArg(0) = InstanceID
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetVolumeDBRange", InArg, OutArg)
            MyCurrentMinValue = OutArg(0)
            MyCurrentMaxValue = OutArg(1)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetVolumeDBRange : MyMinValue = " & MyCurrentMinValue.ToString, LogType.LOG_TYPE_INFO)
                Log("RCGetVolumeDBRange : MyMaxValue = " & MyCurrentMaxValue.ToString, LogType.LOG_TYPE_INFO)
            End If
            RCGetVolumeDBRange = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetVolumeDBRange for device = " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetLoudness(Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCGetLoudness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetLoudness called for device " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetLoudness", InArg, OutArg)
            SetLoudness = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetLoudness : MyCurrentLoudness = " & MyCurrentLoudness.ToString, LogType.LOG_TYPE_INFO)
            End If
            LoudnessIsConfigurable = True
            RCGetLoudness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetLoudness for device = " & MyUPnPDeviceName & " with Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetLoudness(DesiredLoudness As Boolean, Optional Channel As String = "Master", Optional InstanceID As Integer = 0) As String
        RCSetLoudness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetLoudness called for device " & MyUPnPDeviceName & " with DesiredLoudness = " & DesiredLoudness.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Channel
            InArg(2) = DesiredLoudness
            RenderingControl.InvokeAction("SetLoudness", InArg, OutArg)
            RCSetLoudness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetLoudness for device = " & MyUPnPDeviceName & " with DesiredLoudness = " & DesiredLoudness.ToString & " and Channel = " & Channel & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetBrightness(Optional InstanceID As Integer = 0) As String
        RCGetBrightness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetBrightness called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("GetBrightness", InArg, OutArg)
            MyCurrentBrightness = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetBrightness : MyCurrentBrightness = " & MyCurrentBrightness.ToString, LogType.LOG_TYPE_INFO)
            End If
            BrightnessIsConfigurable = True
            RCGetBrightness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetBrightness for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetBrightness(DesiredBrightness As Integer, Optional InstanceID As Integer = 0) As String
        RCSetBrightness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetBrightness called for device " & MyUPnPDeviceName & " with DesiredBrightness = " & DesiredBrightness.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = DesiredBrightness
            RenderingControl.InvokeAction("SetBrightness", InArg, OutArg)
            RCSetBrightness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetBrightness for device = " & MyUPnPDeviceName & " with DesiredBrightness = " & DesiredBrightness.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetContrast(Optional InstanceID As Integer = 0) As String
        RCGetContrast = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetContrast called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("GetContrast", InArg, OutArg)
            MyCurrentContrast = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetContrast : MyCurrentContrast = " & MyCurrentContrast.ToString, LogType.LOG_TYPE_INFO)
            End If
            ContrastIsConfigurable = True
            RCGetContrast = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetContrast for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetContrast(DesiredContrast As Integer, Optional InstanceID As Integer = 0) As String
        RCSetContrast = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetContrast called for device " & MyUPnPDeviceName & " with DesiredContrast = " & DesiredContrast.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = DesiredContrast
            RenderingControl.InvokeAction("SetContrast", InArg, OutArg)
            RCSetContrast = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetContrast for device = " & MyUPnPDeviceName & " with DesiredContrast = " & DesiredContrast.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetSharpness(Optional InstanceID As Integer = 0) As String
        RCGetSharpness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetSharpness called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("GetSharpness", InArg, OutArg)
            MyCurrentSharpness = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetSharpness : MyCurrentSharpness = " & MyCurrentSharpness.ToString, LogType.LOG_TYPE_INFO)
            End If
            SharpnessIsConfigurable = True
            RCGetSharpness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetSharpness for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetSharpness(DesiredSharpness As Integer, Optional InstanceID As Integer = 0) As String
        RCSetSharpness = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetSharpness called for device " & MyUPnPDeviceName & " with DesiredSharpness = " & DesiredSharpness.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = DesiredSharpness
            RenderingControl.InvokeAction("SetSharpness", InArg, OutArg)
            RCSetSharpness = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetSharpness for device = " & MyUPnPDeviceName & " with DesiredSharpness = " & DesiredSharpness.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCGetColorTemperature(Optional InstanceID As Integer = 0) As String
        RCGetColorTemperature = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCGetColorTemperature called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("GetColorTemperature", InArg, OutArg)
            MyCurrentColorTemperature = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCGetColorTemperature : MyCurrentColorTemperature = " & MyCurrentColorTemperature.ToString, LogType.LOG_TYPE_INFO)
            End If
            ColorTemperatureIsConfigurable = True
            RCGetColorTemperature = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCGetColorTemperature for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCSetColorTemperature(DesiredColorTemperature As Integer, Optional InstanceID As Integer = 0) As String
        RCSetColorTemperature = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCSetColorTemperature called for device " & MyUPnPDeviceName & " with DesiredColorTemperature = " & DesiredColorTemperature.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = DesiredColorTemperature
            RenderingControl.InvokeAction("SetColorTemperature", InArg, OutArg)
            RCSetColorTemperature = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCSetColorTemperature for device = " & MyUPnPDeviceName & " with DesiredColorTemperature = " & DesiredColorTemperature.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_GetSlideShowEffect(Optional InstanceID As Integer = 0) As String
        RCX_GetSlideShowEffect = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_GetSlideShowEffect called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("X_GetSlideShowEffect", InArg, OutArg)
            MyCurrentSlideShowEffect = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCX_GetSlideShowEffect : MySlideShowEffect = " & MyCurrentSlideShowEffect.ToString, LogType.LOG_TYPE_INFO)
            End If
            GetSlideShowEffectIsConfigurable = True
            RCX_GetSlideShowEffect = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCX_GetSlideShowEffect for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_SetSlideShowEffect(SlideShowEffect As String, Optional InstanceID As Integer = 0) As String
        RCX_SetSlideShowEffect = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_SetSlideShowEffect called for device " & MyUPnPDeviceName & " with SlideShowEffect = " & SlideShowEffect & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = SlideShowEffect
            RenderingControl.InvokeAction("X_SetSlideShowEffect", InArg, OutArg)
            RCX_SetSlideShowEffect = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCX_SetSlideShowEffect for device = " & MyUPnPDeviceName & " with SlideShowEffect = " & SlideShowEffect & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_GetImageScale(Optional InstanceID As Integer = 0) As String
        RCX_GetImageScale = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_GetImageScale called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("X_GetImageScale", InArg, OutArg)
            MyCurrentImageScale = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCX_GetImageScale : MyImageScale = " & MyCurrentImageScale.ToString, LogType.LOG_TYPE_INFO)
            End If
            GetImageScaleIsConfigurable = True
            RCX_GetImageScale = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCX_GetImageScale for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_SetImageScale(ImageScale As Integer, Optional InstanceID As Integer = 0) As String
        RCX_SetImageScale = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_SetImageScale called for device " & MyUPnPDeviceName & " with ImageScale = " & ImageScale.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = ImageScale
            RenderingControl.InvokeAction("X_SetImageScale", InArg, OutArg)
            RCX_SetImageScale = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCX_SetImageScale for device = " & MyUPnPDeviceName & " with ImageScale = " & ImageScale.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_GetImageRotation(Optional InstanceID As Integer = 0) As String
        RCX_GetImageRotation = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_GetImageRotation called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            RenderingControl.InvokeAction("X_GetImageRotation", InArg, OutArg)
            MyCurrentImageRotation = OutArg(0)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("RCX_GetImageRotation : MyImageRotation = " & MyCurrentImageRotation.ToString, LogType.LOG_TYPE_INFO)
            End If
            GetImageRotationIsConfigurable = True
            RCX_GetImageRotation = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in RCX_GetImageRotation for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCX_SetImageRotation(ImageRotation As Integer, Optional InstanceID As Integer = 0) As String
        RCX_SetImageRotation = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCX_SetImageRotation called for device " & MyUPnPDeviceName & " with ImageRotation = " & ImageRotation.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = ImageRotation
            RenderingControl.InvokeAction("X_SetImageRotation", InArg, OutArg)
            RCX_SetImageRotation = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCX_SetImageRotation for device = " & MyUPnPDeviceName & " with ImageRotation = " & ImageRotation.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetVolumeLevel(ByVal Channel As String) As String
        GetVolumeLevel = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetVolumeLevel called for device = " & MyUPnPDeviceName & " with values Channel=" & Channel & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = 0
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetVolume", InArg, OutArg)
            GetVolumeLevel = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetVolumeLevel for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SetVolumeLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        SetVolumeLevel = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetVolumeLevel called for device = " & MyUPnPDeviceName & " with values Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                If NewLevel < MyMinimumVolume Then
                    NewLevel = MyMinimumVolume
                ElseIf NewLevel > MyMaximumVolume Then
                    NewLevel = MyMaximumVolume
                End If
            Else
                Log(IFACE_NAME, "ERROR in SetVolumeLevel for device = " & MyUPnPDeviceName & ". strChannel = " & Channel & " and Channel value must be 'Master', 'LF', or 'RF'")
                SetVolumeLevel = "Error: Channel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            SetVolumeLevel = RCSetVolume(NewLevel, Channel, )
        Catch ex As Exception
            Log("ERROR in SetVolumeLevel for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ChangeVolumeLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        ChangeVolumeLevel = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeVolumeLevel called for device = " & MyUPnPDeviceName & " for Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArgGetVolume(1)
            Dim InArgSetVolume(2)
            Dim OutArg(0)
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArgGetVolume(0) = 0
                InArgGetVolume(1) = Channel
                InArgSetVolume(0) = 0
                InArgSetVolume(1) = Channel
                Try
                    RenderingControl.InvokeAction("GetVolume", InArgGetVolume, OutArg)
                Catch ex As Exception
                    Log("ERROR in ChangeVolumeLevel for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                NewLevel = OutArg(0) + NewLevel
                If NewLevel < MyMinimumVolume Then
                    NewLevel = MyMinimumVolume
                ElseIf NewLevel > MyMaximumVolume Then
                    NewLevel = MyMaximumVolume
                End If
                InArgSetVolume(2) = NewLevel
            Else
                ChangeVolumeLevel = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            RenderingControl.InvokeAction("SetVolume", InArgSetVolume, OutArg)
            ChangeVolumeLevel = "OK"
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeVolumeLevel called for device = " & MyUPnPDeviceName & " for Channel=" & Channel & " New Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in ChangeVolumeLevel for device = " & MyUPnPDeviceName & " for Channel=" & Channel & " New Value=" & NewLevel.ToString & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub SetBalance(inBalance As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetBalance called for device - " & MyUPnPDeviceName & " with Value = " & inBalance.ToString, LogType.LOG_TYPE_INFO)
        If inBalance < -MyMaximumVolume Or inBalance > MyMaximumVolume Then Exit Sub
        If inBalance < MyMinimumVolume Then
            If MyCurrentLFVolumeLevel <> MyMaximumVolume Then
                SetVolumeLevel("LF", MyMaximumVolume)
            End If
            If MyCurrentRFVolumeLevel <> MyMaximumVolume + inBalance Then
                SetVolumeLevel("RF", MyMaximumVolume + inBalance)
            End If
        ElseIf inBalance > MyMinimumVolume Then
            If MyCurrentRFVolumeLevel <> MyMaximumVolume Then
                SetVolumeLevel("RF", MyMaximumVolume)
            End If
            If MyCurrentLFVolumeLevel <> MyMaximumVolume - inBalance Then
                SetVolumeLevel("LF", MyMaximumVolume - inBalance)
            End If
        Else
            If MyCurrentLFVolumeLevel <> MyMaximumVolume Then
                SetVolumeLevel("LF", MyMaximumVolume)
            End If
            If MyCurrentRFVolumeLevel <> MyMaximumVolume Then
                SetVolumeLevel("RF", MyMaximumVolume)
            End If
        End If
    End Sub

    Public Function ChangeBalanceLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        ChangeBalanceLevel = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeBalanceLevel called for device = " & MyUPnPDeviceName & " with values Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        If (Channel <> "LF" And Channel <> "RF") Or NewLevel = 0 Then
            ' wrong input
            Exit Function
        End If
        Dim LeftLevel, RightLevel As Integer
        LeftLevel = GetVolumeLevel("LF")
        RightLevel = GetVolumeLevel("RF")
        ' when the balance is in the middle both LF and RF are equal to 100
        ' if the balance is to the left from midpoint and moving left then only left should be touched else right and left may need ajustment
        Dim DirectionLeft As Boolean
        If Channel = "LF" And NewLevel > MyMinimumVolume Then
            ' we're moving to the left
            DirectionLeft = True
        ElseIf Channel = "LF" And NewLevel < MyMinimumVolume Then
            ' we're moving to the right
            DirectionLeft = False
        ElseIf Channel = "RF" And NewLevel > MyMinimumVolume Then
            ' we're moving to the right
            DirectionLeft = False
        Else
            ' we're moving to the left
            DirectionLeft = True
        End If
        If DirectionLeft Then
            ' decrease right unless left < 100 then increase left first
            If LeftLevel < MyMaximumVolume Then
                ' adjust right and left levels
                ChangeVolumeLevel("LF", NewLevel)
            Else
                ChangeVolumeLevel("RF", -NewLevel)
            End If
        Else
            ' decrease left unless Right < 100 then increase right first
            If RightLevel < MyMaximumVolume Then
                ' adjust right and left levels
                ChangeVolumeLevel("RF", NewLevel)
            Else
                ChangeVolumeLevel("LF", -NewLevel)
            End If
        End If
    End Function


    Private Sub CheckRenderState()
        If Not MyAdminStateActive Then Exit Sub
        If MyUPnPDeviceServiceType = "HST" Then Exit Sub ' no need to ping this device
        If MyPollRenderStateReEntry Then Exit Sub
        MyPollRenderStateReEntry = True
        Dim CurrentVolume As Integer = MyCurrentVolumeLevel
        If VolumeIsConfigurable Then RCGetVolume()
        If CurrentVolume <> MyCurrentVolumeLevel Then
            Dim RenderingInfo As String = ""
            RenderingInfo = "<p>Master Volume = " & MyCurrentVolumeLevel.ToString & "</p><p>Left = " & MyCurrentLFVolumeLevel.ToString & "</p><p>Right = " & MyCurrentRFVolumeLevel.ToString & "</p>"
            hs.SetDeviceString(MyVolumeHSRef, RenderingInfo, True)
            hs.SetDeviceValueByRef(MyVolumeHSRef, MyCurrentVolumeLevel, True)
            PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False) ' this will update iPad clients
        End If
        Dim CurrentMutestate As Boolean = MyCurrentMuteState
        If MuteIsConfigurable Then RCGetMute()
        If MyCurrentMuteState <> CurrentMutestate Then
            PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
            If MyCurrentMuteState Then
                hs.SetDeviceString(MyMuteHSRef, "Master Mute = On", True)
                hs.SetDeviceValueByRef(MyMuteHSRef, msMuted, True)
            Else
                hs.SetDeviceString(MyMuteHSRef, "Master Mute = Off", True)
                hs.SetDeviceValueByRef(MyMuteHSRef, msUnmuted, True)
            End If
        End If
        MyPollRenderStateReEntry = False
    End Sub


End Class


