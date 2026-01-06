Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Xml
Imports System.Drawing
Imports System.IO

Partial Public Class HSPI

    Private Function RetrieveDIALAppList(AppURL As String) As Boolean
        RetrieveDIALAppList = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveDIALAppList called for device - " & MyUPnPDeviceName & " with AppURL = " & AppURL.ToString, LogType.LOG_TYPE_INFO)
        If String.IsNullOrEmpty(AppURL) Then Exit Function
        ' http://www.dial-multiscreen.org/dial-registry/namespace-database
        ' YouTube Netflix Pandora   Hulu YahooScreen Vimeo Plex FiberTV AmazonInstantVideo VUDU Movies  
        ' the URL to the DIAL information is returned in the header with a header field = "Application-URL"
        ' So first do get on the Location Info
        CheckDIALAppExists(AppURL, "Netflix")
        CheckDIALAppExists(AppURL, "YouTube")
        CheckDIALAppExists(AppURL, "Pandora")
        CheckDIALAppExists(AppURL, "Hulu")
        CheckDIALAppExists(AppURL, "Plex")
        CheckDIALAppExists(AppURL, "AmazonInstantVideo")
        CheckDIALAppExists(AppURL, "VUDU")

    End Function


    Private Function CheckDIALAppExists(AppURL As String, AppName As String) As String
        CheckDIALAppExists = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckDIALAppExists called for device - " & MyUPnPDeviceName & " with AppURL = " & AppURL & " and AppName = " & AppName, LogType.LOG_TYPE_INFO)
        If AppURL = "" Then Exit Function
        If AppName = "" Then Exit Function
        If AppURL((AppURL.Length) - 1) = "/" Then
            AppURL = AppURL + AppName
        Else
            AppURL = AppURL + "/" + AppName
        End If
        Dim RequestUri As Uri = Nothing
        Try
            RequestUri = New Uri(AppURL)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            'If MyAuthenticationCookieContainer IsNot Nothing Then
            'wRequest.CookieContainer = MyAuthenticationCookieContainer
            'Else
            'MyAuthenticationCookieContainer = New CookieContainer
            'Try
            'If System.IO.File.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json") Then
            'Dim Stream As Stream = System.IO.File.Open(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json", FileMode.Open)
            'Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            'MyAuthenticationCookieContainer = formatter.Deserialize(Stream)
            'wRequest.CookieContainer = MyAuthenticationCookieContainer
            'Stream.Dispose()
            'End If
            'Catch ex As Exception
            'Log("Error in CheckDIALAppExists for device - " & MyUPnPDeviceName & " restoring the Cookie Container with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'MyAuthenticationCookieContainer = Nothing
            'End Try
            'End If

            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Try
                'MyApplicationURL = webResponse.Headers("Application-URL")
                'If MyApplicationURL <> "" Then
                'URLDoc = MyApplicationURL
                'Else '
                'MyApplicationURL = URLDoc
                'End If
                Dim webStream As Stream = Nothing
                webStream = webResponse.GetResponseStream
                Dim xmlDoc As New XmlDocument
                xmlDoc.Load(webStream)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(" CheckDIALAppExists for device = " & MyUPnPDeviceName & " got response = " & xmlDoc.OuterXml, LogType.LOG_TYPE_INFO)
                webStream.Close()
                webResponse.Close()
            Catch ex As Exception

            End Try
            webResponse.Close()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in CheckDIALAppExists for device - " & MyUPnPDeviceName & " retrieving the document with AppURL = " & AppURL.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        RequestUri = Nothing
    End Function


    Private Sub CreateHSDIALRemoteButtons(ReCreate As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSDIALRemoteButtons called for device - " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
        HSRefRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Remote.ToString & "HSCode", -1)
        If HSRefRemote = -1 Then
            HSRefRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Remote.ToString)
        Else
            If Not ReCreate Then Exit Sub
        End If

        If HSRefRemote = -1 Then Exit Sub

        If ReCreate Then
            hs.DeviceVSP_ClearAll(HSRefRemote, True)  ' added v24
            hs.DeviceVGP_ClearAll(HSRefRemote, True)  ' added v24
        End If
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psCreateRemoteButtons
        Pair.Status = "Create Remote Buttons"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
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


End Class
