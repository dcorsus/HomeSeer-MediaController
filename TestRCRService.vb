Partial Public Class HSPI


    Public Function RCRAddMessage(MessageID As String, MessageType As String, Message As String) As String
        RCRAddMessage = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCRAddMessage called for device " & MyUPnPDeviceName & " and MessageID = " & MessageID & " and MessageType = " & MessageType & " and Message = " & Message, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = MessageID
            InArg(1) = MessageType
            InArg(2) = Message
            RemoteControlService.InvokeAction("AddMessage", InArg, OutArg)
            RCRAddMessage = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCRAddMessage for device = " & MyUPnPDeviceName & " and MessageID = " & MessageID & " and MessageType = " & MessageType & " and Message = " & Message & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCRRemoveMessage(MessageID As String) As String
        RCRRemoveMessage = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCRRemoveMessage called for device " & MyUPnPDeviceName & " and MessageID = " & MessageID, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = MessageID
            RemoteControlService.InvokeAction("RemoveMessage", InArg, OutArg)
            RCRRemoveMessage = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCRRemoveMessage for device = " & MyUPnPDeviceName & " and MessageID = " & MessageID & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RCRSendKeyCode(KeyCode As Integer, KeyDescription As String) As String
        RCRSendKeyCode = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RCRSendKeyCode called for device " & MyUPnPDeviceName & " and KeyCode = " & KeyCode.ToString & " and KeyDescription = " & KeyDescription, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = KeyCode
            InArg(1) = KeyDescription
            RemoteControlService.InvokeAction("SendKeyCode", InArg, OutArg)
            RCRSendKeyCode = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RCRSendKeyCode for device = " & MyUPnPDeviceName & " and KeyCode = " & KeyCode.ToString & " and KeyDescription = " & KeyDescription & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


End Class
