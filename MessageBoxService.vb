Partial Public Class HSPI

    Public Function PMBAddMessage(MessageID As String, MessageType As String, Message As String) As String
        PMBAddMessage = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PMBAddMessage called for device " & MyUPnPDeviceName & " and MessageID = " & MessageID & " and MessageType = " & MessageType & " and Message = " & Message, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = MessageID
            InArg(1) = MessageType
            InArg(2) = Message
            MessageBoxService.InvokeAction("AddMessage", InArg, OutArg)
            PMBAddMessage = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PMBAddMessage for device = " & MyUPnPDeviceName & " and MessageID = " & MessageID & " and MessageType = " & MessageType & " and Message = " & Message & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PMBRemoveMessage(MessageID As String) As String
        PMBRemoveMessage = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PMBRemoveMessage called for device " & MyUPnPDeviceName & " and MessageID = " & MessageID, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = MessageID
            MessageBoxService.InvokeAction("RemoveMessage", InArg, OutArg)
            PMBRemoveMessage = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PMBRemoveMessage for device = " & MyUPnPDeviceName & " and MessageID = " & MessageID & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub CreateMessageButtons(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

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
        Pair.Value = dsActivateOnLine
        Pair.Status = "Activated Online"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "OKBtn.png"
        GraphicsPair.Set_Value = dsActivateOnLine
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

    End Sub


End Class
