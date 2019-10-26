Partial Public Class HSPI

    Private ConnectionManager As MyUPnPService = Nothing
    Public WithEvents myConnectionManagerCallback As New myUPnPControlCallback

    Private MyRCSID As Integer = 0
    Private MyAVTransportID As Integer = 0
    Private MyProtocolInfo As String = ""
    Private MyPeerConnectionManager As String = ""
    Private MyPeerConnectionID As String = ""
    Private MyDirection As String = ""
    Private MyStatus As String = ""
    Private MySource As String = ""
    Private MySink As String = ""
    Private MyConnectionIDs As String = ""
    Private MySourceProtocolInfo As String()
    Private MySinkProtocolInfo As String()



    Private Sub ConnectionManagerStateChange(ByVal StateVarName As String, ByVal Value As Object) Handles myConnectionManagerCallback.ControlStateChange
        'If piDebuglevel > DebugLevel.dlEvents Then log( "ConnectionManager Change callback for device = " & MyUPnPDeviceName & ". Var Name = " & StateVarName & " Value = " & Value.ToString)
        'If piDebuglevel > DebugLevel.dlErrorsOnly And Not piDebuglevel > DebugLevel.dlEvents Then log( "ConnectionManager Change callback for device = " & MyUPnPDeviceName & ". Var Name = " & StateVarName)
        If StateVarName = "CurrentConnectionIDs" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly And Not PIDebuglevel > DebugLevel.dlEvents Then Log("ConnectionManager Change callback for device = " & MyUPnPDeviceName & ". Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ElseIf StateVarName = "SinkProtocolInfo" Then
            ExtractSinkProtocolInfo(Value)
        ElseIf StateVarName = "SourceProtocolInfo" Then
            ExtractSourceProtocolInfo(Value)
        Else
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning : ConnectionManager Change callback for device = " & MyUPnPDeviceName & " received an unknown Variable = " & StateVarName & " and data = " & Value.ToString, LogType.LOG_TYPE_WARNING)
        End If
    End Sub

    Private Sub ConnectionManagerDied() Handles myConnectionManagerCallback.ControlDied 'ConnectionManagerDied
        'Log("ConnectionManager Control Callback Died for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_WARNING)
        Try
            Log("UPnP connection to UPnPDevice " & MyUPnPDeviceName & " was lost in ConnectionManagerDied.", LogType.LOG_TYPE_WARNING)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR in ConnectionManagerDied for UPnPDevice - " & MyUPnPDeviceName & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub ExtractSourceProtocolInfo(Info As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSourceProtocolInfo called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Info = Trim(Info)
        If Info = "" Then Exit Sub
        MySourceProtocolInfo = Split(Info, ",")
        If PIDebuglevel > DebugLevel.dlEvents Then
            If Not MySourceProtocolInfo Is Nothing Then
                For Each SourceProtocol In MySourceProtocolInfo
                    Log("ExtractSourceProtocolInfo found Protocol = " & SourceProtocol.ToString, LogType.LOG_TYPE_INFO)
                Next
            End If
        End If
    End Sub

    Private Sub ExtractSinkProtocolInfo(Info As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSinkProtocolInfo called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Info = Trim(Info)
        If Info = "" Then Exit Sub
        MySinkProtocolInfo = Split(Info, ",")
        If PIDebuglevel > DebugLevel.dlEvents Then
            If Not MySinkProtocolInfo Is Nothing Then
                For Each SinkProtocol In MySinkProtocolInfo
                    Log("ExtractSinkProtocolInfo found Protocol = " & SinkProtocol.ToString, LogType.LOG_TYPE_INFO)
                Next
            End If
        End If
    End Sub

    Private Function CheckForProtocol(Source As Boolean, Protocol As String) As Boolean
        If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckForProtocol called for device - " & MyUPnPDeviceName & " with SourceFlag = " & Source.ToString & " and Protocol = " & Protocol.ToString, LogType.LOG_TYPE_INFO)
        ' the format is <protocol>“:”<network>“:”<contentFormat>“:”<additionalInfo> where each of the 4 elements MAY be a wildcard “*”.
        ' we apparantly need to check for the additional info when available as well. Looks as follows DLNA.ORG_PN=AVC_MP4_BL_L31_HD_AAC;DLNA.ORG_FLAGS=8d700000000000000000000000000000 
        CheckForProtocol = False
        Dim ProtocolInfo As String()
        Dim ProtocolToCheck As String()
        If Source Then
            If MySourceProtocolInfo Is Nothing Then
                CheckForProtocol = True
                Exit Function
            End If
            ProtocolToCheck = MySourceProtocolInfo
        Else
            If MySinkProtocolInfo Is Nothing Then
                CheckForProtocol = True
                Exit Function
            End If
            ProtocolToCheck = MySinkProtocolInfo
        End If
        ProtocolInfo = Split(Protocol, ":")
        For Each SourceSinkProtocol In ProtocolToCheck
            Dim SourceSinkProtocolInfo As String()
            SourceSinkProtocolInfo = Split(SourceSinkProtocol.ToString, ":")
            Dim Index As Integer = 0
            Dim Different As Boolean = False
            For Index = 0 To 2
                Try
                    If SourceSinkProtocolInfo(Index) <> "*" And ProtocolInfo(Index) <> "*" Then
                        If UCase(Trim(SourceSinkProtocolInfo(Index))) <> UCase(Trim(ProtocolInfo(Index))) Then
                            Different = True
                            Exit For   ' they are different, next one
                        End If
                    End If
                Catch ex As Exception
                    Log("Error in CheckForProtocol for device - " & MyUPnPDeviceName & " and Protocol = " & Protocol.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
            Next
            If Not Different Then
                ' they were the same now let's check whether there was additional info
                CheckForProtocol = True
                If UBound(SourceSinkProtocolInfo, 1) > 2 And UBound(ProtocolInfo, 1) > 2 Then
                    ' there is additional information in the format DLNA.ORG_PN=AVC_MP4_BL_L31_HD_AAC;DLNA.ORG_FLAGS=8d700000000000000000000000000000
                    Dim AdditionalSourceSinkInfos As String()
                    AdditionalSourceSinkInfos = Split(SourceSinkProtocolInfo(3).ToString, ";")
                    Dim AdditionalSourceSinkInfo As String
                    For Each AdditionalSourceSinkInfo In AdditionalSourceSinkInfos
                        If AdditionalSourceSinkInfo.IndexOf("DLNA.ORG_PN=") <> -1 Then
                            AdditionalSourceSinkInfo = AdditionalSourceSinkInfo.Remove(0, AdditionalSourceSinkInfo.IndexOf("=") + 1)
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CheckForProtocol found Additional SourceSinkProtocolInfo for device - " & MyUPnPDeviceName & " with Additional Protocol = " & AdditionalSourceSinkInfo & " and Protocol = " & Protocol.ToString)
                            Dim AdditionalProtocolInfos As String()
                            AdditionalProtocolInfos = Split(ProtocolInfo(3).ToString, ";")
                            Dim AdditionalProtocolInfo As String
                            For Each AdditionalProtocolInfo In AdditionalProtocolInfos
                                If AdditionalProtocolInfo.IndexOf("DLNA.ORG_PN=") <> -1 Then
                                    AdditionalProtocolInfo = AdditionalProtocolInfo.Remove(0, AdditionalProtocolInfo.IndexOf("=") + 1)
                                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CheckForProtocol found Additional ProtocolInfo for device - " & MyUPnPDeviceName & " with Additional Protocol = " & AdditionalSourceSinkInfo & " and Protocol = " & Protocol.ToString)
                                    If AdditionalProtocolInfo <> AdditionalSourceSinkInfo Then
                                        ' keep this 
                                        CheckForProtocol = False
                                        GoTo TheyAreDifferent
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
TheyAreDifferent:
                If CheckForProtocol Then
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CheckForProtocol found Protocol for device - " & MyUPnPDeviceName & " with Protocol = " & Protocol & " and SourceProtocol = " & SourceProtocol.ToString)
                    Exit Function
                End If

            End If
        Next
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Warning CheckForProtocol did not find Protocol for device - " & MyUPnPDeviceName & " with Protocol = " & Protocol)
    End Function



    Public Function CMGetCurrentConnectionInfo(Optional ConnectionID As Integer = 0) As String
        CMGetCurrentConnectionInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CMGetCurrentConnectionInfo called for device " & MyUPnPDeviceName & " and ConnectionID = " & ConnectionID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(6)
            InArg(0) = ConnectionID ' ConnectionID
            ConnectionManager.InvokeAction("GetCurrentConnectionInfo", InArg, OutArg)
            MyRCSID = OutArg(0)                 ' UI4
            MyAVTransportID = OutArg(1)         ' UI4
            MyProtocolInfo = OutArg(2)          ' String
            MyPeerConnectionManager = OutArg(3) ' String
            MyPeerConnectionID = OutArg(4)      ' String
            MyDirection = OutArg(5)             ' String
            MyStatus = OutArg(6)                ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("CMGetCurrentConnectionInfo : MyRCSID                 = " & MyRCSID.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyAVTransportID         = " & MyAVTransportID.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyProtocolInfo          = " & MyProtocolInfo.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyPeerConnectionManager = " & MyPeerConnectionManager.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyPeerConnectionID      = " & MyPeerConnectionID.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyDirection             = " & MyDirection.ToString, LogType.LOG_TYPE_INFO)
                Log("CMGetCurrentConnectionInfo : MyStatus                = " & MyStatus.ToString, LogType.LOG_TYPE_INFO)
            End If
            CMGetCurrentConnectionInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CMGetCurrentConnectionInfo for device = " & MyUPnPDeviceName & " and ConnectionID = " & ConnectionID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function CMGetProtocolInfo() As String
        CMGetProtocolInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CMGetProtocolInfo called for device " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(1)
            ConnectionManager.InvokeAction("GetProtocolInfo", InArg, OutArg)
            MySource = OutArg(0) ' String
            MySink = OutArg(1)   ' String
            ExtractSourceProtocolInfo(MySource)
            ExtractSinkProtocolInfo(MySink)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("CMGetProtocolInfo : MySource = " & MySource.ToString, LogType.LOG_TYPE_INFO)
                'log( "CMGetProtocolInfo : MySink   = " & MySink.ToString)
            End If
            CMGetProtocolInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CMGetProtocolInfo for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function CMGetCurrentConnectionIDs() As String
        CMGetCurrentConnectionIDs = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CMGetCurrentConnectionIDs called for device " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            ConnectionManager.InvokeAction("GetCurrentConnectionIDs", InArg, OutArg)
            MyConnectionIDs = OutArg(0) ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("CMGetCurrentConnectionIDs : MyConnectionIDs = " & MyConnectionIDs.ToString, LogType.LOG_TYPE_INFO)
            End If
            CMGetCurrentConnectionIDs = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CMGetCurrentConnectionIDs for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function




End Class

