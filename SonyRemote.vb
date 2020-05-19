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

    Private MySonyRegisterURL As String = ""
    Private MySonyAppControlURL As String = ""
    Public MySonyRegisterMode As String = ""
    Private MySonyRemoteCommandListURL As String = ""
    Private MySonyContentListURL As String = ""
    Private MySonySystemInformationURL As String = ""
    Private MySonyGetContentURL As String = ""
    Private MySonySendContentURL As String = ""
    Private MySonySendTextURL As String = ""
    Private MySonyWebServiceList As String = ""
    Private MySonyGetStatusURL As String = ""
    Private MySonyActionHeader As String = "X-CERS-DEVICE-ID"
    Private MySonyRDISEntryPort As String = ""
    Private MyCurrentSonyRemoteStatus As String = ""
    Private MyCurrentSonyCommandInfo As String = ""
    Private MyCurrentSonySingerCapability As Integer = 0
    Private MyCurrentSonyTransportPort As Integer = 0
    Private MyCurrentSonyPartyState As String = ""
    Private MyCurrentSonyPartyMode As String = ""
    Private MyCurrentSonyPartySong As String = ""
    Private MyCurrentSonySessionID As Integer = 0
    Private MyCurrentSonyNumberOfListeners As Integer = 0
    Private MyCurrentSonyListenerList As String = ""
    Private MyCurrentSonySingerUUID As String = ""
    Private MyCurrentSonySingerSessionID As Integer = 0
    Private MyCurrentSonyListenerSessionID As Integer = 0
    Private MyAuthenticationCookieContainer As CookieContainer = Nothing

    Private SonyPartyService As MyUPnPService = Nothing
    Public WithEvents mySonyPartyCallback As New myUPnPControlCallback 'myUPnPSonyPartyCallback

    'https://developer.sony.com/develop/google-tv/get-started/


    ' urn:schemas-sony-com:service:IRCC:1
    ' http://www.remotecentral.com/cgi-bin/mboard/rs232-ip/thread.cgi?171,3

    'GET /getStatus HTTP/1.1
    'Host: 192.168.1.132:50002
    'User-Agent: MediaRemote/3.1.0 CFNetwork/548.1.4 Darwin/11.0.0
    'X-CERS-DEVICE-INFO: iPhone OS5.1/MediaRemote3.0.0/iPad1,1
    'X-CERS-DEVICE-ID: MediaRemote:b8-ff-61-c7-fb-6d
    'Connection: close

    '<?xml version="1.0" encoding="UTF-8"?>
    '<statusList>
    '  <status name="party">
    '   <statusItem field="status" value="idle"/>
    '</status>
    '</statusList>

    'http://192.168.1.132:50002/getSystemInformation
    '<systemInformation>
    ' <name>NetBox</name>
    ' <generation>2011</generation>
    ' <remoteType bundled="true">RMT-D302</remoteType>
    '   <actionHeader name="CERS-DEVICE-ID"/>
    '   <supportContentsClass>
    '       <class>video</class>
    '       <class>music</class>
    '   </supportContentsClass>
    '   <supportSource>
    '       <source>Net</source>
    '   </supportSource>
    '</systemInformation>


    'GET /getText HTTP/1.1
    'Host: 192.168.1.132:50002
    'User-Agent: MediaRemote/3.1.0 CFNetwork/548.1.4 Darwin/11.0.0
    'X-CERS-DEVICE-INFO: iPhone OS5.1/MediaRemote3.0.0/iPad1,1
    'X-CERS-DEVICE-ID: MediaRemote:b8-ff-61-c7-fb-6d
    'Connection: close



    '    LIST OF COMMAND CODES FOR THE SONY BDP-S780 BLU-RAY PLAYER
    'The following list of command codes were contained in the xml file that was return after sending
    'the above request remote command packet.
    '========================================================================
    'Confirm:AAAAAwAAHFoAAAA9Aw==
    'Up:AAAAAwAAHFoAAAA5Aw==
    'Down:AAAAAwAAHFoAAAA6Aw==
    'Right:AAAAAwAAHFoAAAA8Aw==
    'Left:AAAAAwAAHFoAAAA7Aw==
    'Home:AAAAAwAAHFoAAABCAw==
    'Options:AAAAAwAAHFoAAAA/Aw==
    'Return:AAAAAwAAHFoAAABDAw==
    'Num1:AAAAAwAAHFoAAAAAAw==
    'Num2:AAAAAwAAHFoAAAABAw==
    'Num3:AAAAAwAAHFoAAAACAw==
    'Num4:AAAAAwAAHFoAAAADAw==
    'Num5:AAAAAwAAHFoAAAAEAw==
    'Num6:AAAAAwAAHFoAAAAFAw==
    'Num7:AAAAAwAAHFoAAAAGAw==
    'Num8:AAAAAwAAHFoAAAAHAw==
    'Num9:AAAAAwAAHFoAAAAIAw==
    'Num0:AAAAAwAAHFoAAAAJAw==
    'Power:AAAAAwAAHFoAAAAVAw==
    'Display:AAAAAwAAHFoAAABBAw==
    'Audio:AAAAAwAAHFoAAABkAw==
    'SubTitle:AAAAAwAAHFoAAABjAw==
    'Favorites:AAAAAwAAHFoAAABeAw==
    'Yellow:AAAAAwAAHFoAAABpAw==
    'Blue:AAAAAwAAHFoAAABmAw==
    'Red:AAAAAwAAHFoAAABnAw==
    'Green:AAAAAwAAHFoAAABoAw==
    'Play:AAAAAwAAHFoAAAAaAw==
    'Stop:AAAAAwAAHFoAAAAYAw==
    'Pause:AAAAAwAAHFoAAAAZAw==
    'Rewind:AAAAAwAAHFoAAAAbAw==
    'Forward:AAAAAwAAHFoAAAAcAw==
    'Prev:AAAAAwAAHFoAAABXAw==
    'Next:AAAAAwAAHFoAAABWAw==
    'Replay:AAAAAwAAHFoAAAB2Aw==
    'Advance:AAAAAwAAHFoAAAB1Aw==
    'Angle:AAAAAwAAHFoAAABlAw==
    'TopMenu:AAAAAwAAHFoAAAAsAw==
    'PopUpMenu:AAAAAwAAHFoAAAApAw==
    'Eject:AAAAAwAAHFoAAAAWAw==
    'Karaoke:AAAAAwAAHFoAAABKAw==
    'Qriocity:AAAAAwAAHFoAAABMAw==
    'Netflix:AAAAAwAAHFoAAABLAw==
    'Mode3D:AAAAAwAAHFoAAABNAw==


    ' AQAAAwAADfoAAAB6Aw
    ' AAAAAwAADfoAAAB8Aw==



    '    <?xml version="1.0" encoding="UTF-8"?>
    '<systemInformation>
    '  <name>NetBox</name>
    '  <generation>2011</generation>
    '  <remoteType bundled="true">RMT-D302</remoteType>
    '  <actionHeader name="CERS-DEVICE-ID"/>
    '  <supportContentsClass>
    '    <class>video</class>
    '    <class>music</class>
    '  </supportContentsClass>
    '  <supportSource>
    '    <source>Net</source>
    '  </supportSource>
    '</systemInformation>

    Public ReadOnly Property SonyPartyState
        Get
            SonyPartyState = MyCurrentSonyPartyState
        End Get
    End Property

    Public ReadOnly Property SonySingerSessionID
        Get
            SonySingerSessionID = MyCurrentSonySingerSessionID
        End Get
    End Property

    Public ReadOnly Property SonySingerUUID
        Get
            SonySingerUUID = MyCurrentSonySingerUUID
        End Get
    End Property


    Public ReadOnly Property SonySessionID
        Get
            SonySessionID = MyCurrentSonySessionID
        End Get
    End Property

    Private Sub SonyPartyStateChange(ByVal StateVarName As String, ByVal Value As Object) Handles mySonyPartyCallback.ControlStateChange 'SonyPartyStateChange
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SonyPartyStateChange for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly And Not PIDebuglevel > DebugLevel.dlEvents Then Log("SonyPartyStateChange for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName, LogType.LOG_TYPE_INFO)

        Try
            SonyPartyX_GetDeviceInfo()
        Catch ex As Exception
            Log("Error in SonyPartyStateChange 1 for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            SonyPartyX_GetState()
        Catch ex As Exception
            Log("Error in SonyPartyStateChange 2 for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Select Case StateVarName.ToUpper
                Case "X_PARTYSTATE"
                    SonySetHSState(Value.ToString)
                Case "X_NUMBEROFLISTENERS"
                    MyCurrentSonyNumberOfListeners = Value
                Case Else
                    Log("Warning : SonyPartyStateChange for device = " & MyUPnPDeviceName & " found unrecognized StateVarName = " & StateVarName.ToString & " and Value = " & Value.ToString, LogType.LOG_TYPE_WARNING)
            End Select
        Catch ex As Exception
            Log("Error in SonyPartyStateChange for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub SonySetHSState(PartyState As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonySetHSState called for device = " & MyUPnPDeviceName & " with PartyState = " & PartyState.ToString & " and HSRef = " & HSRefParty.ToString, LogType.LOG_TYPE_INFO)
        If HSRefParty = -1 Then Exit Sub
        Dim StateInfo As String = ""
        Try
            Select Case PartyState.ToUpper
                Case "IDLE"
                    hs.SetDeviceValueByRef(HSRefParty, dsDeactivated, True)
                    PlayChangeNotifyCallback(player_status_change.PartyOff, player_state_values.UpdateHSServerOnly)
                    hs.SetDeviceString(HSRefParty, "Idle", True)
                    'Case "PARTY"
                    '    hs.SetDeviceValue(MySonyPartyHSCode, 201)
                    '    PlayChangeNotifyCallback(player_status_change.PartyOn, player_state_values.UpdateHSServerOnly)
                    '    hs.SetDeviceString(MySonyPartyHSCode, "Party")
                Case "ABORT_SINGING"
                Case "START_SINGING"
                Case "LISTENING"
                    hs.SetDeviceValueByRef(HSRefParty, dsActivateOnLine, True)
                    PlayChangeNotifyCallback(player_status_change.PartyOn, player_state_values.UpdateHSServerOnly)
                    hs.SetDeviceString(HSRefParty, "Listening to " & MyReferenceToMyController.GetDeviceGivenNameByUDN(MyCurrentSonySingerUUID), True)
                Case "SINGING"
                    hs.SetDeviceValueByRef(HSRefParty, dsActivatedOffLine, True)
                    PlayChangeNotifyCallback(player_status_change.PartyOn, player_state_values.UpdateHSServerOnly)
                    hs.SetDeviceString(HSRefParty, "Singing", True)
                Case "NOT_READY"
                Case Else
                    Log("Warning : SonySetHSState for device = " & MyUPnPDeviceName & " found unrecognized PartyState = " & PartyState.ToString, LogType.LOG_TYPE_WARNING)
            End Select
        Catch ex As Exception
            Log("Error in SonySetHSState for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try


    End Sub

    Private Sub SonyPartyDied() Handles mySonyPartyCallback.ControlDied 'SonyPartyDied
        'Log("SonyPartyDied for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_WARNING)
        Try
            Log("UPnP connection to device " & MyUPnPDeviceName & " was lost in SonyPartyDied.", LogType.LOG_TYPE_WARNING)
            Disconnect(False)
        Catch ex As Exception
            Log("Error in SonyPartyDied for device = " & MyUPnPDeviceName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub RetrieveSonyActionList(URLDoc As String)
        ' http://192.168.1.132:50002/actionList

        ' This is old
        '<actionList>
        '  <action name="register" mode="1" url="http://192.168.1.132:50002/register"/>
        '  <action name="getText" url="http://192.168.1.132:50002/getText"/>
        '  <action name="sendText" url="http://192.168.1.132:50002/sendText"/>
        '  <action name="getContentInformation" url="http://192.168.1.132:50002/getContentInformation"/>
        '  <action name="getSystemInformation" url="http://192.168.1.132:50002/getSystemInformation"/>
        '  <action name="getRemoteCommandList" url="http://192.168.1.132:50002/getRemoteCommandList"/>
        '  <action name="getStatus" url="http://192.168.1.132:50002/getStatus"/>
        '</actionList>

        ' This is new
        '<actionList>
        '  <action name="register" mode="2" url="http://192.168.1.147:31038/cers?action=register"/>
        '  <action name="getWebServiceList" url="http://192.168.1.147:31038/cers?action=getWebServiceList"/>
        '  <action name="getContentUrl" url="http://192.168.1.147:31038/cers?action=getContentUrl"/>
        '  <action name="getSystemInformation" url="http://192.168.1.147:31038/cers?action=getSystemInformation"/>
        '  <action name="sendContentUrl" url="http://192.168.1.147:31038/cers?action=sendContentUrl"/>
        '  <action name="sendText" url="http://192.168.1.147:31038/cers?action=sendText"/>
        '</actionList>

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonyActionList called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then Exit Sub
        Dim xmlDoc As New XmlDocument
        Dim RequestUri = New Uri(URLDoc)
        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) ' HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            '            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "MediaController/MediaController/MediaController,1")
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView:" & AdjustMacAddressforSony(MyMacAddress))
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            xmlDoc.Load(webStream)
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RetrieveSonyActionList for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & URLDoc.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
            xmlDoc = Nothing
        End Try

        If xmlDoc.HasChildNodes Then
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlDoc.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & " Document root node: " & xmlDoc.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO) 
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & " Document root node: " & xmlDoc.DocumentElement.Name, LogType.LOG_TYPE_INFO) 
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    Dim ActionName As String = ""
                    Dim ActionURL As String = ""
                    ActionName = outerNode.Attributes("name").Value
                    ActionURL = outerNode.Attributes("url").Value
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & "------> Action Name: " & ActionName, LogType.LOG_TYPE_INFO)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & "------> Action URL: " & ActionURL, LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & "------> Action Name: " & ActionName, LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonyActionList for device - " & MyUPnPDeviceName & "------> Action URL: " & ActionURL, LogType.LOG_TYPE_INFO) 
                    If ActionName.ToUpper = "REGISTER" Then
                        MySonyRegisterURL = ActionURL
                        MySonyRegisterMode = outerNode.Attributes("mode").Value
                    ElseIf ActionName.ToUpper = "GETSTATUS" Then
                        MySonyGetStatusURL = ActionURL
                    ElseIf ActionName.ToUpper = "GETREMOTECOMMANDLIST" Then
                        MySonyRemoteCommandListURL = ActionURL
                    ElseIf ActionName.ToUpper = "GETCONTENTINFORMATION" Then
                        MySonyContentListURL = ActionURL
                    ElseIf ActionName.ToUpper = "GETWEBSERVICELIST" Then
                        MySonyWebServiceList = ActionURL
                    ElseIf ActionName.ToUpper = "GETSYSTEMINFORMATION" Then
                        MySonySystemInformationURL = ActionURL
                    ElseIf ActionName.ToUpper = "GETCONTENTURL" Then
                        MySonyGetContentURL = ActionURL
                    ElseIf ActionName.ToUpper = "SENDCONTENTURL" Then
                        MySonySendContentURL = ActionURL
                    ElseIf ActionName.ToUpper = "SENDTEXT" Then
                        MySonySendTextURL = ActionURL
                    End If
                Next
            Catch ex As Exception
                Log("Error in RetrieveSonyActionList for UPnPDevice = " & MyUPnPDeviceName & "  processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        xmlDoc = Nothing
    End Sub

    Private Function AdjustMacAddressforSony(inMacAddress As String) As String
        AdjustMacAddressforSony = inMacAddress
        Dim index As Integer = 2
        Try
            For index = 10 To 2 Step -2
                inMacAddress = inMacAddress.Insert(index, "-")
            Next
            AdjustMacAddressforSony = inMacAddress.ToUpper
        Catch ex As Exception
        End Try
    End Function

    Private Function SonyRegister(URLDoc As String, SendRenew As Boolean) As Boolean

        ' To register, go first to 

        ' <action name="register" mode="2" url="http://192.168.1.147:31038/cers?action=register"/>
        'GET /cers?action=register&name=Dirk%27s%20iPhone%20%28TV%20SideView%29&registrationType=initial&deviceId=TVSideView%3Af7-07-4a-01-f0-3a HTTP/1.1
        'Host: 192.168.1.147:31038
        'User-Agent: TVSideView/2.5.0 CFNetwork/672.1.14 Darwin/14.0.0
        'Connection: close()
        'X-CERS-DEVICE-INFO: iPhone OS7.1.1/TVSideView2.5.0/iPhone5,1

        'HTTP/1.1 200 OK
        'Pragma: no-Cache
        'Cache-Control: no-Cache,no-store


        SonyRegister = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyRegister called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString & " and SendRenew = " & SendRenew.ToString, LogType.LOG_TYPE_INFO)
        Dim sonyPIN As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")

        If MySonyRegisterMode = "JSON" Then
            Return SendJSONAuthentication(sonyPIN)
        ElseIf MySonyRegisterMode = "3" Then
            Return SendMode3Authentication(sonyPIN, SendRenew)
        End If

        ' http://192.168.1.131:50002/register?action=register&name=DeviceName&RegistrationType=initial&deviceId=MediaRemote:MA-CA-DD-RE-SS

        Dim RegistrationType As String = "initial"
        If SendRenew Then RegistrationType = "renewal"
        If URLDoc.IndexOf("?") <> -1 Then ' MySonyRegisterMode = "1" Then
            'URLDoc = URLDoc & "?name=Dirk%27s%20iPhone%20%28TV%20SideView%29&registrationType=initial&deviceId=TVSideView%3A00219b23aaf7"
            'URLDoc = URLDoc & "?action=register&name=" & sIFACE_NAME & "&registrationType=" & RegistrationType & "&deviceId=" & "TVSideView" & "%3A" & MyMacAddress
            'URLDoc = URLDoc & "?action=register&name=" & sIFACE_NAME & "&registrationType=" & RegistrationType & "&deviceId=" & "TVSideView" & "%3A" & AdjustMacAddressforSony(MyMacAddress)
            URLDoc = URLDoc & "&name=" & sIFACE_NAME & "&registrationType=" & RegistrationType & "&deviceId=" & "TVSideView" & "%3A" & AdjustMacAddressforSony(MyMacAddress)
        ElseIf MySonyRegisterMode = "3" Then
            ' https://github.com/KHerron/SonyAPILib/blob/master/SonyAPILib/SonyAPILib/sonyAPILib.cs
            URLDoc = URLDoc & "?name=" & sIFACE_NAME & "&registrationType=" & RegistrationType & "&deviceId=" & "TVSideView" & "%3A" & AdjustMacAddressforSony(MyMacAddress) & "&wolSupport=true"
        Else  'If MySonyRegisterMode = "2" Then
            URLDoc = URLDoc & "?name=" & sIFACE_NAME & "&registrationType=" & RegistrationType & "&deviceId=" & "TVSideView" & "%3A" & AdjustMacAddressforSony(MyMacAddress)
        End If


        Dim RequestUri = New Uri(URLDoc)
        Dim webResponse As WebResponse = Nothing
        Dim webStream As Stream = Nothing
        Dim ResponseHTML As String = ""

        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.KeepAlive = False
            '            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "MediaController/MediaController/MediaController,1")
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            webResponse = wRequest.GetResponse
            webStream = webResponse.GetResponseStream
            webStream.Close()
            webResponse.Close()
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyRegister for device - " & MyUPnPDeviceName & " registered successfully with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)
        Catch ex As WebException ' if we are trying to renew and the response = 403 than the device is not registered and the registration flag should be reset!!
            Log("Error in SonyRegister for device - " & MyUPnPDeviceName & " registering with URL = " & URLDoc.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ResponseHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
            End If
            Log("Error in SonyRegister for device - " & MyUPnPDeviceName & " registering with URL = " & URLDoc.ToString & " with HTML = " & ResponseHTML, LogType.LOG_TYPE_ERROR)
            Exit Function
        Catch ex As Exception
            Log("Error in SonyRegister for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
            Exit Function
        End Try
        Return True
    End Function

    Private Sub GetSonyRemoteCommandList(URLDoc As String)

        'GET /cers?action=getWebServiceList&lang=en HTTP/1.1'
        'Host: 192.168.1.147:31038
        'X-CERS-DEVICE-INFO: iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1
        'User-Agent: TVSideView/2.5.0 CFNetwork/672.1.14 Darwin/14.0.0
        'Connection: close()
        'X-CERS-DEVICE-ID: TVSideView:f7-07-4a-01-f0-3a

        ' getSystemInformation returns the following. Note the actionHeader !!

        '<?xml version="1.0" encoding="UTF-8"?>
        '<systemInformation>
        '  <name>NetBox</name>
        '  <generation>2011</generation>
        '  <remoteType bundled="true">RMT-D302</remoteType>
        '  <actionHeader name="CERS-DEVICE-ID"/>
        '  <supportContentsClass>
        '    <class>video</class>
        '   <class>music</class>
        '  </supportContentsClass>
        '  <supportSource>
        '     <source>Net</source>
        '  </supportSource>
        '</systemInformation>

        ' example at http://msdn.microsoft.com/en-us/library/debx8sh9.aspx
        ' SONY SMP-N200
        '<?xml version="1.0" encoding="UTF-8"?>
        '<remoteCommandList>
        '  <command name="Power" type="ircc" value="AAAAAwAADfoAAAAVAw==" />
        '  <command name="Red" type="ircc" value="AAAAAwAADvoAAABnAw==" />
        '  <command name="Green" type="ircc" value="AAAAAwAADvoAAABoAw==" />
        '  <command name="Yellow" type="ircc" value="AAAAAwAADvoAAABpAw==" />
        '  <command name="Blue" type="ircc" value="AAAAAwAADvoAAABmAw==" />
        '  <command name="Display" type="ircc" value="AAAAAwAADfoAAABUAw==" />
        '  <command name="Up" type="ircc" value="AAAAAwAADfoAAAB5Aw==" />
        '  <command name="Qriocity" type="ircc" value="AAAAAwAADvoAAABMAw==" />
        '  <command name="Left" type="ircc" value="AAAAAwAADfoAAAB7Aw==" />
        '  <command name="Confirm" type="ircc" value="AAAAAwAADfoAAAALAw==" />
        '  <command name="Right" type="ircc" value="AAAAAwAADfoAAAB8Aw==" />
        '  <command name="Return" type="ircc" value="AAAAAwAADfoAAAAOAw==" />
        '  <command name="Down" type="ircc" value="AAAAAwAADfoAAAB6Aw==" />
        '  <command name="Options" type="ircc" value="AAAAAwAADvoAAAAXAw==" />
        '  <command name="Prev" type="ircc" value="AAAAAwAADfoAAAAwAw==" />
        '  <command name="Home" type="ircc" value="AAAAAwAADfoAAABTAw==" />
        '  <command name="Next" type="ircc" value="AAAAAwAADfoAAAAxAw==" />
        '  <command name="Rewind" type="ircc" value="AAAAAwAADfoAAAAiAw==" />
        '  <command name="Play" type="ircc" value="AAAAAwAADfoAAAAyAw==" />
        '  <command name="Forward" type="ircc" value="AAAAAwAADfoAAAAjAw==" />
        '  <command name="Pause" type="ircc" value="AAAAAwAADfoAAAA5Aw==" />
        '  <command name="Stop" type="ircc" value="AAAAAwAADfoAAAA4Aw==" />
        '  <command name="PartyOn" type="url" value="http://192.168.1.132:50002/setParty?action=start" />
        '  <command name="PartyOff" type="url" value="http://192.168.1.132:50002/setParty?action=stop" />
        '</remoteCommandList>

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyRemoteCommandList called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyRemoteCommandList called for device - " & MyUPnPDeviceName & " has no Remote Command List to retrieve", LogType.LOG_TYPE_INFO)
            'processSonyCommand("Power", "AAAAAQAAAAEAAAAVAw==", 21, 3, 1)
            'processSonyCommand("PowerOff", "AAAAAQAAAAEAAAAvAw==", 22, 3, 2)

            Exit Sub
            ' this device wasn't publishing it, go for default
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyRemoteCommandList called for device - " & MyUPnPDeviceName & " has no Remote Command List to retrieve, create default", LogType.LOG_TYPE_INFO)
            processSonyCommand("Confirm", "AAAAAwAAHFoAAAA9Aw==", 20, 2, 1)
            processSonyCommand("Up", "AAAAAwAAHFoAAAA5Aw==", 21, 2, 2)
            processSonyCommand("Down", "AAAAAwAAHFoAAAA6Aw==", 22, 2, 3)
            processSonyCommand("Right", "AAAAAwAAHFoAAAA8Aw==", 23, 2, 4)
            processSonyCommand("Left", "AAAAAwAAHFoAAAA7Aw==", 24, 3, 1)
            processSonyCommand("Home", "AAAAAwAAHFoAAABCAw==", 25, 3, 2)
            processSonyCommand("Options", "AAAAAwAAHFoAAAA/Aw==", 26, 3, 3)
            processSonyCommand("Return", "AAAAAwAAHFoAAABDAw==", 27, 3, 4)
            processSonyCommand("Num1", "AAAAAwAAHFoAAAAAAw==", 28, 4, 1)
            processSonyCommand("Num2", "AAAAAwAAHFoAAAABAw==", 29, 4, 2)
            processSonyCommand("Num3", "AAAAAwAAHFoAAAACAw==", 30, 4, 3)
            processSonyCommand("Num4", "AAAAAwAAHFoAAAADAw==", 31, 4, 4)
            processSonyCommand("Num5", "AAAAAwAAHFoAAAAEAw==", 32, 5, 1)
            processSonyCommand("Num6", "AAAAAwAAHFoAAAAFAw==", 33, 5, 2)
            processSonyCommand("Num7", "AAAAAwAAHFoAAAAGAw==", 34, 5, 3)
            processSonyCommand("Num8", "AAAAAwAAHFoAAAAHAw==", 35, 5, 4)
            processSonyCommand("Num9", "AAAAAwAAHFoAAAAIAw==", 36, 6, 1)
            processSonyCommand("Num0", "AAAAAwAAHFoAAAAJAw==", 37, 6, 2)
            processSonyCommand("Power", "AAAAAwAAHFoAAAAVAw==", 38, 6, 3)
            processSonyCommand("Display", "AAAAAwAAHFoAAABBAw==", 39, 6, 4)
            processSonyCommand("Audio", "AAAAAwAAHFoAAABkAw==", 40, 7, 1)
            processSonyCommand("SubTitle", "AAAAAwAAHFoAAABjAw==", 41, 7, 2)
            processSonyCommand("Favorites", "AAAAAwAAHFoAAABeAw==", 42, 7, 3)
            processSonyCommand("Yellow", "AAAAAwAAHFoAAABpAw==", 43, 7, 4)
            processSonyCommand("Blue", "AAAAAwAAHFoAAABmAw==", 44, 8, 1)
            processSonyCommand("Red", "AAAAAwAAHFoAAABnAw==", 45, 8, 2)
            processSonyCommand("Green", "AAAAAwAAHFoAAABoAw==", 46, 8, 3)
            processSonyCommand("Play", "AAAAAwAAHFoAAAAaAw==", 47, 8, 4)
            processSonyCommand("Stop", "AAAAAwAAHFoAAAAYAw==", 48, 9, 1)
            processSonyCommand("Pause", "AAAAAwAAHFoAAAAZAw==", 49, 9, 2)
            processSonyCommand("Rewind", "AAAAAwAAHFoAAAAbAw==", 50, 9, 3)
            processSonyCommand("Forward", "AAAAAwAAHFoAAAAcAw==", 51, 9, 4)
            processSonyCommand("Prev", "AAAAAwAAHFoAAABXAw==", 52, 10, 1)
            processSonyCommand("Next", "AAAAAwAAHFoAAABWAw==", 53, 10, 2)
            processSonyCommand("Replay", "AAAAAwAAHFoAAAB2Aw==", 54, 10, 3)
            processSonyCommand("Advance", "AAAAAwAAHFoAAAB1Aw==", 55, 10, 4)
            processSonyCommand("Angle", "AAAAAwAAHFoAAABlAw==", 56, 11, 1)
            processSonyCommand("TopMenu", "AAAAAwAAHFoAAAAsAw==", 57, 11, 2)
            processSonyCommand("PopUpMenu", "AAAAAwAAHFoAAAApAw==", 58, 11, 3)
            processSonyCommand("Eject", "AAAAAwAAHFoAAAAWAw==", 59, 11, 4)
            processSonyCommand("Karaoke", "AAAAAwAAHFoAAABKAw==", 60, 12, 1)
            processSonyCommand("Qriocity", "AAAAAwAAHFoAAABMAw==", 61, 12, 2)
            processSonyCommand("Netflix", "AAAAAwAAHFoAAABLAw==", 62, 12, 3)
            processSonyCommand("Mode3D", "AAAAAwAAHFoAAABNAw==", 63, 12, 4)
            Exit Sub
        End If

        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument
        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.Host = RequestUri.Authority
            wRequest.KeepAlive = True
            wRequest.AllowAutoRedirect = True
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1")
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
            Dim sonyPIN As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")
            If sonyPIN <> "" Then
                Dim authInfo As String = ":" & sonyPIN
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo))
                wRequest.Headers.Add("Authorization", "Basic " & authInfo)
            End If
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webResponse As WebResponse = Nothing
        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webStream As Stream = Nothing
        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            xmlDoc.Load(webStream)
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " loading the XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Dim ButtonIndex As Integer = 20
        Dim RowIndex As Integer = 2
        Dim ColumnIndex As Integer = 1
        Try
            If xmlDoc.HasChildNodes Then
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlDoc.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & " Document root node: " & xmlDoc.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    Dim CommandName As String = outerNode.Attributes("name").Value
                    Dim CommandType As String = outerNode.Attributes("type").Value
                    Dim CommandValue As String = outerNode.Attributes("value").Value
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & "------> Command Name: " & CommandName, LogType.LOG_TYPE_INFO)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & "------> Command Type: " & CommandType, LogType.LOG_TYPE_INFO)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("GetSonyRemoteCommandList for device - " & MyUPnPDeviceName & "------> Command Value: " & CommandValue, LogType.LOG_TYPE_INFO)
                    If CommandType.ToUpper = "IRCC" Then processSonyCommand(CommandName, CommandValue, ButtonIndex, RowIndex, ColumnIndex)
                    ButtonIndex += 1
                    ColumnIndex += 1
                    If ColumnIndex > 4 Then
                        RowIndex += 1
                        ColumnIndex = 1
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in GetSonyRemoteCommandList for UPnPDevice = " & MyUPnPDeviceName & "  processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub GetSonyContentList(URLDoc As String)
        ' example at http://msdn.microsoft.com/en-us/library/debx8sh9.aspx
        ' SONY SMP-N200


        ' this Blue Ray player
        ' <?xml version="1.0" encoding="UTF-8"?>
        '<contentInformation>
        '   <infoItem field = "class" value="video" />
        '   <infoItem field = "source" value="BD" />
        '   <infoItem field = "mediaType" value="BD-ROM" />
        '   <infoItem field = "mediaFormat" value="BDMV" />
        '</contentInformation>

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyContentList called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then Exit Sub
        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument
        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.KeepAlive = True
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.AllowAutoRedirect = True
            wRequest.Host = RequestUri.Authority
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
            Dim sonyPIN As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")
            If sonyPIN <> "" Then
                Dim authInfo As String = ":" & sonyPIN
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo))
                wRequest.Headers.Add("Authorization", "Basic " & authInfo)
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub ' there is none, exit function
        End Try
        Dim webResponse As WebResponse = Nothing
        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Catch ex As Exception
            Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webStream As Stream = Nothing
        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            xmlDoc.Load(webStream)
        Catch ex As Exception
            Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " loading the XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyContentList for device - " & MyUPnPDeviceName & " found XML = " & xmlDoc.InnerXml.ToString, LogType.LOG_TYPE_INFO)
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in GetSonyContentList for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub GetSonyWebServices(URLDoc As String)


        'GET /cers?action=getWebServiceList&lang=en HTTP/1.1'
        'Host: 192.168.1.147:31038
        'X-CERS-DEVICE-INFO: iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1
        'User-Agent: TVSideView/2.5.0 CFNetwork/672.1.14 Darwin/14.0.0
        'Connection: close()
        'X-CERS-DEVICE-ID: TVSideView:f7-07-4a-01-f0-3a

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyWebServices called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then Exit Sub
        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument
        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) ' HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.ContentLength = 0
            wRequest.Host = RequestUri.Authority
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
            Dim sonyPIN As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")
            If sonyPIN <> "" Then
                Dim authInfo As String = ":" & sonyPIN
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo))
                wRequest.Headers.Add("Authorization", "Basic " & authInfo)
            End If
        Catch ex As Exception
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webResponse As WebResponse = Nothing
        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Catch ex As Exception
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webStream As Stream = Nothing
        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            xmlDoc.Load(webStream)
        Catch ex As Exception
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " loading the XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in GetSonyWebServices for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyWebServices for device - " & MyUPnPDeviceName & " found XML = " & xmlDoc.InnerXml.ToString, LogType.LOG_TYPE_INFO)
        xmlDoc = Nothing

    End Sub

    Private Sub RetrieveSonySystemInformation(URLDoc As String)
        ' https://forums.homeseer.com/forum/media-plug-ins/media-discussion/media-controller-dcorsus/1386002-can-t-add-remote-buttons-for-sony-ubp-x1100es-blu-ray-player

        ' <action name="getSystemInformation" url="http://192.168.1.147:31038/cers?action=getSystemInformation"/>

        '<systemInformation>
        '    <name>Internet TV Box</name>
        '    <generation>NSZGS7</generation>
        '    <remoteType bundled="true">GTB1GBOX_US</remoteType>
        '    <actionHeader name="CERS-DEVICE-ID"/>
        '    <supportContentsClass>
        '        <class>url</class>
        '    </supportContentsClass>
        '    <supportSource>
        '        <source>Net</source>
        '    </supportSource>
        '</systemInformation>

        '<systemInformation>
        '   <name>BDPlayer</name>
        '   <generation>2017</generation>
        '   <remoteType>RMT-B119A</remoteType>
        '   <remoteType>RMT-B120A</remoteType>
        '   <remoteType>RMT-B122A</remoteType>
        '   <remoteType>RMT-B123A</remoteType>
        '   <remoteType bundled = "true" > RMT - B126A</remoteType>
        '   <remoteType>RMT-B119J</remoteType>
        '   <remoteType>RMT-B127J</remoteType>
        '   <remoteType>RMT-B119P</remoteType>
        '   <remoteType>RMT-B120P</remoteType>
        '   <remoteType>RMT-B121P</remoteType>
        '   <remoteType>RMT-B122P</remoteType>
        '   <remoteType>RMT-B127P</remoteType>
        '   <remoteType>RMT-B119C</remoteType>
        '   <remoteType>RMT-B120C</remoteType>
        '   <remoteType>RMT-B122C</remoteType>
        '   <remoteType>RMT-B127C</remoteType>
        '   <remoteType>RMT-B127T</remoteType>
        '   <remoteType>RMT-B115A</remoteType>
        '   <actionHeader name = "CERS-DEVICE-ID" />
        '   <supportContentsClass>
        '        <class>video</Class>
        '       <class>music</Class>
        '   </supportContentsClass>
        '   <supportSource>
        '       <source>BD</source>
        '       <source>DVD</source>
        '       <source>CD</source>
        '       <source>Net</source>
        '   </supportSource>
        '   <supportFunction>
        '       <function name = "Notification" />
        '        <function name="WOL">
        '             <functionItem field="MAC" value="94-db-56-08-24-25"/>
        '        </function>
        '   </supportFunction>
        '</systemInformation>


        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RetrieveSonySystemInformation called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then Exit Sub
        Dim xmlDoc As New XmlDocument
        Dim RequestUri = New Uri(URLDoc)
        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1")
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            xmlDoc.Load(webStream)
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RetrieveSonySystemInformation for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & URLDoc.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
            xmlDoc = Nothing
        End Try
        Try
            MySonyActionHeader = xmlDoc.GetElementsByTagName("actionHeader").Item(0).Attributes("name").Value
        Catch ex As Exception
            MySonyActionHeader = "X-CERS-DEVICE-ID"
        End Try
        MySonyActionHeader = "X-CERS-DEVICE-ID" ' test dcor
        Try
            Dim functionNodeList As XmlNodeList = xmlDoc.GetElementsByTagName("functionItem")
            If functionNodeList IsNot Nothing AndAlso functionNodeList.Count > 0 Then
                For Each functionNode As XmlNode In functionNodeList
                    If functionNode.Attributes("field").Value = "MAC" Then
                        Dim macaddr As String = functionNode.Attributes("value").Value
                        If macaddr <> "" Then WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diMACAddress.ToString, macaddr.Replace("-", ":"))
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in RetrieveSonySystemInformation for device - " & MyUPnPDeviceName & " retrieving the MAC with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        xmlDoc = Nothing
    End Sub

    Private Sub GetSonyStatus(URLDoc As String)
        ' <?xml version="1.0" encoding="UTF-8"?>
        '<statusList>
        '   <status name = "disc">
        '       <statusItem field="type" value="BD" />
        '       <statusItem field = "mediaType" value="BD-ROM" />
        '       <statusItem field = "mediaFormat" value="BDMV" />
        '   </status>
        '</statusList>

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyStatus called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Then Exit Sub
        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument
        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1")
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress))
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub ' there is none, exit function
        End Try
        Dim webResponse As WebResponse = Nothing
        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Catch ex As Exception
            Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim webStream As Stream = Nothing
        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Try
            xmlDoc.Load(webStream)
        Catch ex As Exception
            Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " loading the XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSonyStatus for device - " & MyUPnPDeviceName & " found XML = " & xmlDoc.InnerXml.ToString, LogType.LOG_TYPE_INFO)
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in GetSonyStatus for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub SonySetupRemoteInfo()
        MySonyRegisterMode = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyRemoteRegisterType.ToString, "")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonySetupRemoteInfo called for device = " & MyUPnPDeviceName & " with RegisterMode = " & MySonyRegisterMode, LogType.LOG_TYPE_INFO)
        MyRemoteServiceActive = True
        If MySonyRegisterMode = "1" Then
            GetSonyRemoteCommandList(MySonyRemoteCommandListURL)
            GetSonyContentList(MySonyContentListURL)
            CreateButtonInfoInInifile()
            If HSRefRemote = -1 Then CreateHSSonyRemoteButtons(False)
        ElseIf MySonyRegisterMode = "2" Then
            GetSonyRemoteCommandList(MySonyRemoteCommandListURL)
            GetSonyWebServices(MySonyWebServiceList)
            If HSRefRemote = -1 Then CreateHSSonyRemoteButtons(False)
        ElseIf MySonyRegisterMode = "3" Then    ' added 5/17/2020 in v58
            GetSonyRemoteCommandList(MySonyRemoteCommandListURL)
            GetSonyContentList(MySonyContentListURL)
            GetSonyStatus(MySonyGetStatusURL)  ' added in v60
            CreateButtonInfoInInifile()
            If HSRefRemote = -1 Then CreateHSSonyRemoteButtons(False)
        ElseIf MySonyRegisterMode = "JSON" Then
            GetSonyRemoteCommandList(MySonyRemoteCommandListURL)
            GetSonyWebServices(MySonyWebServiceList)
            If HSRefRemote = -1 Then CreateHSSonyRemoteButtons(False)
        Else    ' maybe some future methods that are non JSON added 5/17/2020 in v.58
            GetSonyRemoteCommandList(MySonyRemoteCommandListURL)
            GetSonyContentList(MySonyContentListURL)
            CreateButtonInfoInInifile()
            If HSRefRemote = -1 Then CreateHSSonyRemoteButtons(False)
        End If
    End Sub

    Private Sub SonyProcessIRCCInfo(IRCCXML As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO) ' 
        Dim xmlDoc As New XmlDocument
        xmlDoc.XmlResolver = Nothing
        Dim ButtonIndex As Integer = 20
        Dim RowIndex As Integer = 2
        Dim ColumnIndex As Integer = 1

        Try
            xmlDoc.LoadXml(IRCCXML)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        '<av:X_IRCCCodeList>
        '	<av:X_IRCCCode command="Power">AAAAAQAAAAEAAAAVAw==</av:X_IRCCCode>
        '</av:X_IRCCCodeList>
        Dim X_IRCCCodeList As XmlNodeList = Nothing
        Try
            X_IRCCCodeList = xmlDoc.GetElementsByTagName("av:X_IRCCCode")
            If X_IRCCCodeList IsNot Nothing Then
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved av:X_IRCCCodeList with " & X_IRCCCodeList.Count.ToString & "Nodes ", LogType.LOG_TYPE_INFO) 
                If X_IRCCCodeList.Count > 0 Then
                    For NodeIndex As Integer = 0 To X_IRCCCodeList.Count - 1
                        Dim Command As String = X_IRCCCodeList.Item(NodeIndex).Attributes("command").Value
                        Dim Value As String = X_IRCCCodeList.Item(NodeIndex).InnerText
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved IRCC Command = " & Command & " and Value = " & Value, LogType.LOG_TYPE_INFO)
                        processSonyCommand(Command, Value, ButtonIndex, RowIndex, ColumnIndex)
                        ButtonIndex += 1
                        ColumnIndex += 1
                        If ColumnIndex > 5 Then
                            ColumnIndex = 1
                            RowIndex += 1
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving av:X_IRCCCode with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        X_IRCCCodeList = Nothing

        '<av:X_IRCC_DeviceInfo>
        '		<av:X_IRCC_Version>1.0</av:X_IRCC_Version>
        '		<av:X_IRCC_CategoryList>'
        '			<av:X_IRCC_Category>
        '				<av:X_CategoryInfo>AAEAAAAB</av:X_CategoryInfo>
        '			</av:X_IRCC_Category>
        '			<av:X_IRCC_Category>
        '				<av:X_CategoryInfo>AAIAAACk</av:X_CategoryInfo>
        '			</av:X_IRCC_Category>
        '			<av:X_IRCC_Category>
        '				<av:X_CategoryInfo>AAIAAACX</av:X_CategoryInfo>
        '			</av:X_IRCC_Category>
        '			<av:X_IRCC_Category>
        '				<av:X_CategoryInfo>AAIAAAB3</av:X_CategoryInfo>
        '			</av:X_IRCC_Category>
        '			<av:X_IRCC_Category>
        '				<av:X_CategoryInfo>AAIAAAAa</av:X_CategoryInfo>
        '			</av:X_IRCC_Category>
        '		</av:X_IRCC_CategoryList>
        '</av:X_IRCC_DeviceInfo>
        Dim X_IRCC_CategoryList As XmlNodeList = Nothing
        Try
            X_IRCC_CategoryList = xmlDoc.GetElementsByTagName("av:X_IRCC_Category")
            If X_IRCC_CategoryList IsNot Nothing Then
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved av:X_IRCC_CategoryList with " & X_IRCC_CategoryList.Count.ToString & "Nodes ", LogType.LOG_TYPE_INFO) 
                If X_IRCC_CategoryList.Count > 0 Then
                    For NodeIndex As Integer = 0 To X_IRCC_CategoryList.Count - 1
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved X_IRCC_Category = " & X_IRCC_CategoryList.Item(NodeIndex).InnerText, LogType.LOG_TYPE_INFO)
                    Next
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving av:X_IRCC_Category with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        X_IRCC_CategoryList = Nothing


        Dim X_ScalarWebAPI_BaseURL As String = ""
        Try
            X_ScalarWebAPI_BaseURL = xmlDoc.GetElementsByTagName("av:X_ScalarWebAPI_BaseURL").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved X_ScalarWebAPI_BaseURL = " & X_ScalarWebAPI_BaseURL, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving av:X_ScalarWebAPI_BaseURL with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Dim X_ScalarWebAPI_ActionList_URL As String = ""
        Try
            X_ScalarWebAPI_ActionList_URL = xmlDoc.GetElementsByTagName("av:X_ScalarWebAPI_ActionList_URL").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved X_ScalarWebAPI_ActionList_URL = " & X_ScalarWebAPI_ActionList_URL, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving av:X_ScalarWebAPI_ActionList_URL with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If (X_ScalarWebAPI_BaseURL = "") And (X_ScalarWebAPI_ActionList_URL = "") Then Exit Sub ' we're done

        If (X_ScalarWebAPI_BaseURL = "") And (X_ScalarWebAPI_ActionList_URL <> "") Then X_ScalarWebAPI_BaseURL = X_ScalarWebAPI_ActionList_URL ' use this

        Try
            ' restore the Cookiecontainer
            If MyAuthenticationCookieContainer Is Nothing Then
                MyAuthenticationCookieContainer = New CookieContainer
            End If
            If System.IO.File.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json") Then
                Dim Stream As Stream = System.IO.File.Open(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json", FileMode.Open)
                Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                MyAuthenticationCookieContainer = formatter.Deserialize(Stream)
                Stream.Dispose()
            End If
        Catch ex As Exception
            Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " restoring the Cookie Container with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyAuthenticationCookieContainer = Nothing
        End Try

        Dim X_ScalarWebAPI_ServiceList As XmlNodeList = Nothing
        Dim JSONSystemURL As String = ""
        Dim ReturnJSON As String = ""

        ' We look for 
        '<av:X_ScalarWebAPI_DeviceInfo>
        '		<av:X_ScalarWebAPI_Version>1.0</av:X_ScalarWebAPI_Version>
        '		<av:X_ScalarWebAPI_BaseURL>http://192.168.1.150/sony</av:X_ScalarWebAPI_BaseURL>
        '		<av:X_ScalarWebAPI_ServiceList>				
        '			<av:X_ScalarWebAPI_ServiceType>guide</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>accessControl</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>encryption</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>contentshare</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>avContent</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>cec</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>audio</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>system</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>appControl</av:X_ScalarWebAPI_ServiceType>
        '			<av:X_ScalarWebAPI_ServiceType>videoScreen</av:X_ScalarWebAPI_ServiceType>
        '		</av:X_ScalarWebAPI_ServiceList>
        '</av:X_ScalarWebAPI_DeviceInfo>

        ' Alternatively on Sony Cameras
        '<av:X_ScalarWebAPI_DeviceInfo xmlns:av="urn:schemas-sony-com:av">
        '   <av:X_ScalarWebAPI_Version>1.0</av:X_ScalarWebAPI_Version>
        '   <av:X_ScalarWebAPI_ServiceList>
        '       <av:X_ScalarWebAPI_Service>
        '           <av:X_ScalarWebAPI_ServiceType>guide</av:X_ScalarWebAPI_ServiceType>
        '           <av:X_ScalarWebAPI_ActionList_URL>http://192.168.122.1:8080/sony</av:X_ScalarWebAPI_ActionList_URL>
        '       </av:X_ScalarWebAPI_Service>
        '   <av:X_ScalarWebAPI_Service>
        '       <av:X_ScalarWebAPI_ServiceType>camera</av:X_ScalarWebAPI_ServiceType>
        '       <av:X_ScalarWebAPI_ActionList_URL>http://192.168.122.1:8080/sony</av:X_ScalarWebAPI_ActionList_URL> </av:X_ScalarWebAPI_Service>
        '   </av:X_ScalarWebAPI_ServiceList>
        '</av:X_ScalarWebAPI_DeviceInfo>

        Try
            X_ScalarWebAPI_ServiceList = xmlDoc.GetElementsByTagName("av:X_ScalarWebAPI_ServiceType")
            If X_ScalarWebAPI_ServiceList IsNot Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved X_ScalarWebAPI_ServiceList with " & X_ScalarWebAPI_ServiceList.Count.ToString & "Nodes ", LogType.LOG_TYPE_INFO)
                If X_ScalarWebAPI_ServiceList.Count > 0 Then
                    For NodeIndex As Integer = 0 To X_ScalarWebAPI_ServiceList.Count - 1
                        Try
                            'ReturnJSON = SendJSON(X_ScalarWebAPI_BaseURL & "/" & X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText, "{""id"":2,""method"":""getVersions"",""version"":""1.0"",""params"":[]}")
                            'ReturnJSON = SendJSON(X_ScalarWebAPI_BaseURL & "/" & X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText, "{""id"":3,""method"":""getMethodTypes"",""version"":""1.0"",""params"":[""1.0""]}")
                            'ReturnJSON = SendJSON(X_ScalarWebAPI_BaseURL & "/" & X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText, "{""id"":3,""method"":""getMethodTypes"",""version"":""1.0"",""params"":[""1.1""]}")
                        Catch ex As Exception

                        End Try
                        If X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText = "system" Then
                            ' OK we have a JSON port to retrieve stufff from
                            JSONSystemURL = X_ScalarWebAPI_BaseURL & "/system"
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved JSON System URL = " & JSONSystemURL, LogType.LOG_TYPE_INFO)
                            MySonyRegisterMode = "JSON"
                        ElseIf X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText = "accessControl" Then
                            MySonyRegisterURL = X_ScalarWebAPI_BaseURL & "/accessControl"
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved JSON AccessControl URL = " & MySonyRegisterURL, LogType.LOG_TYPE_INFO)
                        ElseIf X_ScalarWebAPI_ServiceList.Item(NodeIndex).InnerText = "appControl" Then
                            MySonyAppControlURL = X_ScalarWebAPI_BaseURL & "/appControl"
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " retrieved JSON AppControl URL = " & MySonyRegisterURL, LogType.LOG_TYPE_INFO)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " while retieving av:X_ScalarWebAPI_ServiceType with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        X_ScalarWebAPI_ServiceList = Nothing

        ' JSON in is always
        '   ID
        '   Method
        '   Version
        '   Param

        ' JSON Returned = 
        '   ID
        '   Result

        If JSONSystemURL <> "" Then ' check whether WOL is enabled
            'ReturnJSON = SendJSON(JSONSystemURL, "{""id"":3,""method"":""getWolMode"",""version"":""1.0"",""params"":[""1.0""]}")
            ReturnJSON = SendJSON(JSONSystemURL, "{""id"":3,""method"":""getWolMode"",""version"":""1.0"",""params"":[]}")
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " received JSON reply = " & ReturnJSON, LogType.LOG_TYPE_INFO) 'dcorsony
            ' {"id":3,"result":[{"enabled":true}]}
            ' if need be issue {"id":3,"method":"setWolMode","version":"1.0","params":[{"enabled":true}]}
            ReturnJSON = SendJSON(JSONSystemURL, "{""id"":3,""method"":""setWolMode"",""version"":""1.0"",""params"":[{""enabled"":true}]}")
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " received JSON reply = " & ReturnJSON, LogType.LOG_TYPE_INFO) 'dcorsony
        End If

        Try
            If JSONSystemURL <> "" Then
                ReturnJSON = SendJSON(JSONSystemURL, "{""id"":19,""method"":""getSystemSupportedFunction"",""version"":""1.0"",""params"":[]}")
                ' {"result":[[{"option":"WOL","value":"AC:9B:0A:DF:91:31"}]],"id":19}
                Try
                    Dim json As New JavaScriptSerializer
                    Dim JSONdata
                    'JSONdata = json.Deserialize(ReturnJSON, GetType(SonyJSONSystemSupportedFunction))
                    JSONdata = json.DeserializeObject(ReturnJSON)
                    For Each Entry As Object In JSONdata
                        If Entry.Key = "result" Then
                            For Each Entry2 As Object In Entry.value
                                For Each Resultpair As Object In Entry2
                                    Dim Name As String = ""
                                    Dim Value As String = ""
                                    For Each ValuePair As Object In Resultpair
                                        If ValuePair.key = "option" Then
                                            Name = ValuePair.value
                                        ElseIf ValuePair.key = "value" Then
                                            Value = ValuePair.value
                                        End If
                                    Next
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & "  for SystemSupportedFunction found JSON name = " & Name & " and Value = " & Value, LogType.LOG_TYPE_INFO) 'dcorsonyy
                                    If Name = "WOL" Then
                                        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diMACAddress.ToString, Value.Replace("-", ":")) ' in the format 70:18:8B:97:34:14
                                        'processSonyCommand("WOL", Value, psWOL, RowIndex, ColumnIndex)
                                        'ColumnIndex += 1
                                        'If ColumnIndex > 5 Then
                                        'ColumnIndex = 1
                                        'RowIndex += 1
                                        'End If
                                    End If
                                Next
                            Next
                        ElseIf Entry.key = "id" Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " found SystemSupportedFunction JSON ID = " & Entry.value, LogType.LOG_TYPE_INFO)
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " found SystemSupportedFunction ID = " & Entry.id & " and Value = " & Entry.value, LogType.LOG_TYPE_INFO)
                        End If
                    Next
                    JSONdata = Nothing
                    json = Nothing
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " processing SystemSupportedFunctions with error = " & ex.Message & " and JSON = " & ReturnJSON, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " processing SystemSupportedFunctions1 with JSON data = " & ReturnJSON & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If JSONSystemURL <> "" Then
                ' {"id":20,
                '  "result":[{"bundled":true,"type":"RM-J1100"},[{"name":"PowerOff","value":"AAAAAQAAAAEAAAAvAw=="},
                '                                                {"name":"Input","value":"AAAAAQAAAAEAAAAlAw=="},
                '                                                {"name":"GGuide","value":"AAAAAQAAAAEAAAAOAw=="}
                '                                               ]
                '                                           ]}
                ReturnJSON = SendJSON(JSONSystemURL, "{""id"":20,""method"":""getRemoteControllerInfo"",""version"":""1.0"",""params"":[]}")

                Try
                    Dim json As New JavaScriptSerializer
                    Dim JSONdata
                    'JSONdata = json.Deserialize(ReturnJSON, GetType(SonyJSONRemoteControllerInfo))
                    JSONdata = json.DeserializeObject(ReturnJSON)
                    'Dim Remotecodes As Object = JSONdata.result(1)
                    For Each Entry As System.Collections.Generic.KeyValuePair(Of String, Object) In JSONdata
                        If Entry.Key = "result" Then
                            'Dim Remotecodes As Object = Entry.value(1)
                            For Each Entry2 As System.Collections.Generic.Dictionary(Of String, Object) In Entry.Value(1)
                                Dim Name As String = ""
                                Dim Value As String = ""
                                Try
                                    For Each ValuePair As System.Collections.Generic.KeyValuePair(Of String, Object) In Entry2
                                        If ValuePair.Key = "name" Then
                                            Name = ValuePair.Value
                                        ElseIf ValuePair.Key = "value" Then
                                            Value = ValuePair.Value
                                        End If
                                    Next
                                Catch ex As Exception
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " processing RemoteControllerInfo for Entry2 with error = " & ex.Message & " and JSON = " & ReturnJSON, LogType.LOG_TYPE_ERROR)
                                End Try
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " for RemoteControllerInfo found JSON name = " & Name & " and Value = " & Value, LogType.LOG_TYPE_INFO) 'dcorsony
                                If (Name <> "") And (Value <> "") Then
                                    processSonyCommand(Name, Value, ButtonIndex, RowIndex, ColumnIndex)
                                    ButtonIndex += 1
                                    ColumnIndex += 1
                                    If ColumnIndex > 5 Then
                                        ColumnIndex = 1
                                        RowIndex += 1
                                    End If
                                End If
                            Next
                        ElseIf Entry.Key = "id" Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " found RemoteControllerInfo JSON ID = " & Entry.Value, LogType.LOG_TYPE_INFO)
                        Else
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyProcessIRCCInfo for device = " & MyUPnPDeviceName & " found RemoteControllerInfo ID = " & Entry.Key & " and Value = " & Entry.Value, LogType.LOG_TYPE_INFO)
                        End If
                    Next

                    JSONdata = Nothing
                    json = Nothing
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " processing RemoteControllerInfo with error = " & ex.Message & " and JSON = " & ReturnJSON, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyProcessIRCCInfo for device - " & MyUPnPDeviceName & " processing RemoteControllerInfo1 with JSON data = " & ReturnJSON & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Function SendJSON(URLDoc As String, JSONin As String) As String

        SendJSON = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSON called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString & " and JSONin = " & JSONin, LogType.LOG_TYPE_INFO)
        If URLDoc = "" Or JSONin = "" Then Exit Function

        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument

        Dim data = Encoding.UTF8.GetBytes(JSONin)

        Try
            Dim RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "POST"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.ContentType = "application/json"
            wRequest.ContentLength = data.Length
            If MyAuthenticationCookieContainer IsNot Nothing Then
                wRequest.CookieContainer = MyAuthenticationCookieContainer
            End If
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendJSON for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function ' there is none, exit function
        End Try

        Dim stream = wRequest.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()

        Dim webResponse As WebResponse = Nothing
        Dim webStream As Stream = Nothing
        Dim ResponseHTML As String = ""

        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ResponseHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
                webStream.Dispose()
            End If
            Log("Error in SendJSON for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message & " and Response = " & ResponseHTML, LogType.LOG_TYPE_ERROR)
            Exit Function
        Catch ex As Exception
            Log("Error in SendJSON for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in SendJSON for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            Dim strmRdr As New System.IO.StreamReader(webStream)
            SendJSON = strmRdr.ReadToEnd()
            strmRdr.Dispose()
        Catch ex As Exception
            Log("Error in SendJSON for device - " & MyUPnPDeviceName & " reading the ResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendJSON for device - " & MyUPnPDeviceName & " received JSON = " & SendJSON.ToString, LogType.LOG_TYPE_INFO)

        Try
            data = Nothing
            stream.Dispose()
        Catch ex As Exception
        End Try
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in SendJSON for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function SendJSONAuthentication(SonyPIN As String) As Boolean
        ' https://braviacontroller.codeplex.com/SourceControl/latest#WindowsFormsApplication1/Form1.cs

        SendJSONAuthentication = False
        ' this could be the JSON way of registering
        ' http://192.168.1.61/sony/accessControl
        ' {"id":13,"method":"actRegister","version":"1.0","params":[{"clientid":"TVSideView:34c48639-af3d-40e7-b1b2-74091375368c","nickname":"cm_tenderloin (TV SideView)"},[{"clientid":"TVSideView:34c48639-af3d-40e7-b1b2-74091375368c","value":"yes","nickname":"cm_tenderloin (TV SideView)","function":"WOL"}]]}
        Dim ReturnJSON As String = ""
        'Dim SonyPIN As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication called for device - " & MyUPnPDeviceName & " with SonyRegisterURL = " & MySonyRegisterURL.ToString & " and PIN = " & SonyPIN, LogType.LOG_TYPE_INFO)
        If MySonyRegisterURL = "" Then Exit Function
        SonyPIN = Trim(SonyPIN)

        Dim JSONRegisterString As String = "{""id"":13,""method"":""actRegister"",""version"":""1.0"",""params"":[{""clientid"":""" & "TVSideView:" & AdjustMacAddressforSony(MyMacAddress) & """,""nickname"":""(MediaController)""},[{""clientid"":""TVSideView:" & AdjustMacAddressforSony(MyMacAddress) & """,""value"":""yes"",""nickname"":""(MediaController)"",""function"":""WOL""}]]}"
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication send Registration String for device - " & MyUPnPDeviceName & " with String = " & JSONRegisterString.ToString, LogType.LOG_TYPE_INFO)
        'ReturnJSON = SendJSONAuthentication(MySonyRegisterURL & "/system", JSONRegisterString, "")
        'WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)

        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication called for device - " & MyUPnPDeviceName & " with JSON = " & JSONRegisterString, LogType.LOG_TYPE_INFO)

        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument

        Try
            If MyAuthenticationCookieContainer Is Nothing Then
                MyAuthenticationCookieContainer = New CookieContainer()
            ElseIf SonyPIN = "" Then
                MyAuthenticationCookieContainer = Nothing
                MyAuthenticationCookieContainer = New CookieContainer()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " getting a cookiecontainer error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Dim data = Encoding.UTF8.GetBytes(JSONRegisterString)

        Try
            Dim RequestUri = New Uri(MySonyRegisterURL)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "POST"
            wRequest.KeepAlive = False
            wRequest.AllowAutoRedirect = True
            wRequest.CookieContainer = MyAuthenticationCookieContainer
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.ContentType = "application/json"
            wRequest.ContentLength = data.Length
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1") '"iPhone OS7.1.1/MediaRemote2.5.0/iPhone5,1") ' iPhone OS9.2/MediaRemote3.0.0/iPhone7,1
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
            If SonyPIN <> "" Then
                Dim authInfo As String = ":" & SonyPIN
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo))
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " created Authentication string = " & "Basic " & authInfo, LogType.LOG_TYPE_INFO)
                wRequest.Headers.Add("Authorization", "Basic " & authInfo)
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function ' there is none, exit function
        End Try

        Dim stream = wRequest.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()

        Dim webResponse As HttpWebResponse = Nothing
        Dim webStream As Stream = Nothing
        Dim ResponseHTML As String = ""

        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ResponseHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
                webStream.Dispose()
                Try
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Response Header = " & webResponse.Headers.ToString(), LogType.LOG_TYPE_INFO)
                Catch ex1 As Exception
                End Try
            End If
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message & " and Response = " & ResponseHTML, LogType.LOG_TYPE_ERROR)
            Exit Function
        Catch ex As Exception
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            ' write Authentication cookie to file!
            Dim file As FileStream = System.IO.File.Create(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json")
            Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            'Dim answerCookie As String = Newtonsoft.Json.JsonConvert.SerializeObject(wRequest.CookieContainer.GetCookies(New Uri(MySonyAppControlURL)))
            formatter.Serialize(file, MyAuthenticationCookieContainer)
            file.Close()
        Catch ex As Exception
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " storing the authentication Cookie with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Response Header = " & webResponse.Headers.ToString(), LogType.LOG_TYPE_INFO) ' dcorsony  
        Catch ex As Exception
        End Try

        Try
            For Each cookieValue As Cookie In webResponse.Cookies
                If cookieValue.Name = "auth" Then
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSonyCookieExpiryDate.ToString, cookieValue.Expires)
                    Log("Warning for device - " & MyUPnPDeviceName & ". Authentication will expire = " & cookieValue.Expires, LogType.LOG_TYPE_WARNING)
                    Exit Try
                End If
            Next
        Catch ex As Exception
        End Try

        Try
            For Each cookieValue As Cookie In webResponse.Cookies
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie: " & cookieValue.ToString(), LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie name: " & cookieValue.Name & " has Cookie value: " & cookieValue.Value, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Domain: " & cookieValue.Domain, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Path: " & cookieValue.Path, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Port: " & cookieValue.Port, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Secure: " & cookieValue.Secure, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie TimeStamp: " & cookieValue.TimeStamp, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Expires: " & cookieValue.Expires, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Expired: " & cookieValue.Expired, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Discard: " & cookieValue.Discard, LogType.LOG_TYPE_INFO) ' dcorsony
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie Comment: " & cookieValue.Comment, LogType.LOG_TYPE_INFO) ' dcorsony
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " has Cookie CommentUri: " & cookieValue.CommentUri.AbsoluteUri, LogType.LOG_TYPE_INFO) ' dcorsony
            Next
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " reading the cookies with error = " & ex.Message, LogType.LOG_TYPE_ERROR) 'dcorsony
        End Try

        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            Dim strmRdr As New System.IO.StreamReader(webStream)
            ReturnJSON = strmRdr.ReadToEnd()
            strmRdr.Dispose()
        Catch ex As Exception
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " reading the ResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendJSONAuthentication for device - " & MyUPnPDeviceName & " received JSON = " & ReturnJSON.ToString, LogType.LOG_TYPE_INFO)

        ' I think successful registration response with a string like this {"result":[],"id":13}

        SendJSONAuthentication = True
        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)

        Try
            data = Nothing
            stream.Dispose()
        Catch ex As Exception
        End Try
        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in SendJSONAuthentication for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function SendMode3Authentication(SonyPIN As String, sendRenew As Boolean) As Boolean
        ' https://github.com/KHerron/SonyAPILib/blob/master/SonyAPILib/SonyAPILib/sonyAPILib.cs

        SendMode3Authentication = False

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMode3Authentication called for device - " & MyUPnPDeviceName & " with SonyRegisterURL = " & MySonyRegisterURL.ToString & " and PIN = " & SonyPIN, LogType.LOG_TYPE_INFO)
        If MySonyRegisterURL = "" Then Exit Function
        SonyPIN = Trim(SonyPIN)

        Dim registerURL As String = ""
        If sendRenew Then
            registerURL = MySonyRegisterURL & "?name=" & sIFACE_NAME & "&registrationType=renewal&deviceId=TVSideView%3A" & AdjustMacAddressforSony(MyMacAddress) & "&wolSupport=true"
        Else
            registerURL = MySonyRegisterURL & "?name=" & sIFACE_NAME & "&registrationType=initial&deviceId=TVSideView%3A" & AdjustMacAddressforSony(MyMacAddress) & "&wolSupport=true "
        End If

        Dim wRequest As HttpWebRequest = Nothing
        Dim webResponse As WebResponse = Nothing
        Dim webStream As Stream = Nothing
        Dim ResponseHTML As String = ""
        Dim authInfo As String = ""

        Try
            Dim RequestUri = New Uri(registerURL)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.Host = RequestUri.Authority
            wRequest.KeepAlive = True
            wRequest.AllowAutoRedirect = True
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("X-CERS-DEVICE-INFO", "iPhone OS9.2/MediaController1.0.0/MediaController,1")
            wRequest.Headers.Add(MySonyActionHeader, "TVSideView" & ":" & AdjustMacAddressforSony(MyMacAddress)) ' X-CERS-DEVICE-ID
            If SonyPIN <> "" Then
                authInfo = ":" & SonyPIN
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo))
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMode3Authentication for device - " & MyUPnPDeviceName & " created Authentication string = " & "Basic " & authInfo, LogType.LOG_TYPE_INFO)
                wRequest.Headers.Add("Authorization", "Basic " & authInfo)
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " creating a webrequest with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function ' there is none, exit function
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMode3Authentication for device - " & MyUPnPDeviceName & " is sending registerString = " & registerURL.ToString & " with authInfo = " & authInfo, LogType.LOG_TYPE_INFO)

        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException ' if we are trying to renew and the response = 403 than the device is not registered and the registration flag should be reset!!
            ' a 401 Unauthorized may indicate our PIN is invalid
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " registering with URL = " & registerURL.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, "")
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ResponseHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
            End If
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " registering with URL = " & registerURL.ToString & " with HTML = " & ResponseHTML, LogType.LOG_TYPE_ERROR)
            Exit Function
        Catch ex As Exception
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMode3Authentication for device - " & MyUPnPDeviceName & " has Response Header = " & webResponse.Headers.ToString(), LogType.LOG_TYPE_INFO) ' dcorsony  
        Catch ex As Exception
        End Try


        Try
            webStream = webResponse.GetResponseStream
        Catch ex As Exception
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " doing a GetResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Dim receivedResponse As String = ""

        Try
            Dim strmRdr As New System.IO.StreamReader(webStream)
            receivedResponse = strmRdr.ReadToEnd()
            strmRdr.Dispose()
        Catch ex As Exception
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " reading the ResponseStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMode3Authentication for device - " & MyUPnPDeviceName & " received = " & receivedResponse.ToString, LogType.LOG_TYPE_INFO)

        SendMode3Authentication = True
        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)

        Try
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in SendMode3Authentication for device - " & MyUPnPDeviceName & " closing everything with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Private Sub processSonyCommand(CommandName As String, CommandValue As String, ButtonIndex As Integer, Row As Integer, Column As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("processSonyCommand called for device - " & MyUPnPDeviceName & " with CommandName = " & CommandName & " and CommandValue = " & CommandValue, LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = gRemoteControlPath
        WriteStringIniFile(MyUDN & " - Default Codes", CommandName, CommandValue & ":;:-:" & ButtonIndex.ToString, objRemoteFile)
        WriteStringIniFile(MyUDN, ButtonIndex.ToString, CommandName & ":;:-:" & CommandValue & ":;:-:" & Row.ToString & ":;:-:" & Column.ToString, objRemoteFile)
    End Sub

    Private Sub CreateHSSonyRemoteButtons(ReCreate As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSSonyRemoteButtons called for device - " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
        HSRefRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Remote.ToString & "HSCode", -1)
        If HSRefRemote = -1 Then
            HSRefRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Remote.ToString)
        Else
            If Not ReCreate Then Exit Sub ' already exist
        End If
        If HSRefRemote = -1 Then Exit Sub
        If ReCreate Then
            hs.DeviceVSP_ClearAll(HSRefRemote, True)  ' added v24
            hs.DeviceVGP_ClearAll(HSRefRemote, True)  ' added v24
        End If
        Dim Pair As VSPair

        Dim Column As Integer = 1
        Dim registrationMode = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyRemoteRegisterType.ToString, "")

        If (registrationMode <> "JSON") And (registrationMode = "3") Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psRegister
            Pair.Status = "Register"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Row = 1
            Pair.Render_Location.Column = Column
            Column += 1
            hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        End If

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psCreateRemoteButtons
        Pair.Status = "Create Remote Buttons"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = Column
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

    Private Sub SonyPartyAll()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyAll called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If MyCurrentSonyPartyMode <> "IDLE" Then
            SonyExitParty()
            wait(2)
        End If

        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Dim ListenerList As String = ""
        Try
            For Each PartyDevice In PartyDevices
                If PartyDevice.Key <> "" And PartyDevice.Key <> MyUDN Then
                    Dim DLNADevice As HSPI = Nothing
                    DLNADevice = MyReferenceToMyController.GetAPIByUDN(PartyDevice.Key)
                    If Not DLNADevice Is Nothing Then
                        If DLNADevice.DeviceOnLine Then
                            If ListenerList <> "" Then ListenerList = ListenerList & ","
                            ListenerList = ListenerList & "uuid:" & PartyDevice.Key
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in SonyPartyAll for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If ListenerList <> "" Then
            SonyPartyX_Start("PARTY", ListenerList)
        End If

    End Sub

    Private Sub SonyAddPartyByNumber(PlayerNumber As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyAddPartyByNumber called for UPnPDevice = " & MyUPnPDeviceName & " with Device number to add = " & PlayerNumber, LogType.LOG_TYPE_INFO)
        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Try
            For Each PartyDevice In PartyDevices
                If PartyDevice.Value = PlayerNumber Then
                    Dim DLNADevice As HSPI = Nothing
                    DLNADevice = MyReferenceToMyController.GetAPIByUDN(PartyDevice.Key)
                    If Not DLNADevice Is Nothing Then
                        If DLNADevice.DeviceOnLine Then
                            If MyCurrentSonyListenerList <> "" Then
                                If MyCurrentSonyListenerList.IndexOf(PartyDevice.Key) = 0 Then
                                    SonyPartyX_Start("PARTY", MyCurrentSonyListenerList & ",uuid:" & PartyDevice.Key)
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in SonyAddPartyByNumber for UPnPDevice = " & MyUPnPDeviceName & " with Device number to add = " & PlayerNumber & "; player already in Listenerlist", LogType.LOG_TYPE_WARNING)
                                End If
                            Else
                                SonyPartyX_Start("PARTY", "uuid:" & PartyDevice.Key)
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in SonyAddPartyByNumber for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub SonyAddPartyByName(PlayerName As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyAddPartyByName called for UPnPDevice = " & MyUPnPDeviceName & " with Device to add = " & PlayerName, LogType.LOG_TYPE_INFO)
        'If MyCurrentSonyPartyMode <> "IDLE" Then SonyExitParty()

        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Try
            For Each PartyDevice In PartyDevices
                If PartyDevice.Key <> "" And PartyDevice.Key <> MyUDN Then
                    Dim DLNADevice As HSPI = Nothing
                    DLNADevice = MyReferenceToMyController.GetAPIByUDN(PartyDevice.Key)
                    If Not DLNADevice Is Nothing Then
                        Dim DLNAFriendlyName As String = DLNADevice.DeviceName
                        If DLNADevice.DeviceOnLine And DLNAFriendlyName = PlayerName Then
                            If MyCurrentSonyListenerList <> "" Then
                                If MyCurrentSonyListenerList.IndexOf(PartyDevice.Key) = 0 Then
                                    SonyPartyX_Start("PARTY", MyCurrentSonyListenerList & ",uuid:" & PartyDevice.Key)
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in SonyAddPartyByName for UPnPDevice = " & MyUPnPDeviceName & " with Device to add = " & PlayerName & "; player already in Listenerlist", LogType.LOG_TYPE_WARNING)

                                End If
                            Else
                                SonyPartyX_Start("PARTY", "uuid:" & PartyDevice.Key)
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in SonyAddPartyByName for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in SonyAddPartyByName for UPnPDevice = " & MyUPnPDeviceName & " with Device to add = " & PlayerName & " but player not found", LogType.LOG_TYPE_WARNING)

    End Sub


    Private Sub SonyPartyHost()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyHost called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If MyCurrentSonyPartyMode <> "IDLE" Then
            SonyExitParty()
            wait(2)
        End If
        SonyPartyX_Start("PARTY", "")
    End Sub

    Private Sub SonyExitParty()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyExitParty called for UPnPDevice = " & MyUPnPDeviceName & " and PartyMode = " & MyCurrentSonyPartyMode.ToString & " and PartyState = " & MyCurrentSonyPartyState.ToString, LogType.LOG_TYPE_INFO)
        If MyCurrentSonyPartyMode = "IDLE" Then Exit Sub
        Select Case MyCurrentSonyPartyState
            Case "ABORT_SINGING"
            Case "START_SINGING"
            Case "LISTENING"
                If MyCurrentSonySingerUUID <> "" Then
                    Try
                        Dim Singer As HSPI
                        Singer = MyReferenceToMyController.GetAPIByUDN(MyCurrentSonySingerUUID)
                        Singer.SonyPartyX_Leave(MyCurrentSonySingerSessionID, "uuid:" & MyUDN)
                    Catch ex As Exception
                        Log("Error in SonyExitParty for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Case "SINGING"
                SonyPartyX_Abort(MyCurrentSonySessionID)
            Case Else
                Log("Warning : SonyExitParty for device = " & MyUPnPDeviceName & " found unrecognized PartyState = " & MyCurrentSonyPartyState.ToString, LogType.LOG_TYPE_WARNING)
        End Select
    End Sub

    Private Sub SonyJoinParty()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyJoinParty called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim Singers As String = MyReferenceToMyController.GetSonyPartySingers()
        If Singers = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyJoinParty called for UPnPDevice = " & MyUPnPDeviceName & " but no singers found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        Try
            Dim Singer As HSPI
            Singer = MyReferenceToMyController.GetAPIByUDN(Singers)
            SonyPartyX_Invite("LINK", "uuid:" & Singer.DeviceUDN, Singer.SonySessionID)
        Catch ex As Exception
            Log("Error in SonyJoinParty for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub SonyUpdatePartyButtons()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyUpdatePartyButtons called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        HSRefParty = GetIntegerIniFile(MyUDN, "di" & HSDevices.Party.ToString & "HSCode", -1)
        If HSRefParty = -1 Then
            HSRefParty = CreateHSServiceDevice(HSRefParty, HSDevices.Party.ToString)
            MyReferenceToMyController.AddPartyDevice(MyUDN)
        End If
        If HSRefParty = -1 Then Exit Sub
        hs.DeviceVSP_ClearAll(HSRefParty, True)
        hs.DeviceVGP_ClearAll(HSRefParty, True)
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPartyAll
        Pair.Status = "Party All"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 1
        hs.DeviceVSP_AddPair(HSRefParty, Pair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psExitParty
        Pair.Status = "Exit Party"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 2
        hs.DeviceVSP_AddPair(HSRefParty, Pair)


        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsIdle
        Pair.Status = "Idle"
        hs.DeviceVSP_AddPair(HSRefParty, Pair)
        Dim GraphicsPair As New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "NOKBtn.png"
        GraphicsPair.Set_Value = dsIdle
        hs.DeviceVGP_AddPair(HSRefParty, GraphicsPair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsSinger
        Pair.Status = "Singer"
        hs.DeviceVSP_AddPair(HSRefParty, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "Speaker.jpg"
        GraphicsPair.Set_Value = dsSinger
        hs.DeviceVGP_AddPair(HSRefParty, GraphicsPair)
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = dsListener
        Pair.Status = "Listener"
        hs.DeviceVSP_AddPair(HSRefParty, Pair)
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "Listen.jpg"
        GraphicsPair.Set_Value = dsListener
        hs.DeviceVGP_AddPair(HSRefParty, GraphicsPair)

        Dim PartyDevices As New System.Collections.Generic.Dictionary(Of String, String)()
        PartyDevices = GetIniSection("Party Devices") '  As Dictionary(Of String, String)
        Dim PartyButtonString As String = ""
        Dim PartyValueString As String = ""
        Dim PartyIndex As Integer = 2
        Try
            For Each PartyDevice In PartyDevices
                If PartyDevice.Key <> "" And PartyDevice.Key <> MyUDN Then
                    Dim DLNAFriendlyName As String = ""
                    DLNAFriendlyName = GetStringIniFile(PartyDevice.Key, "diGivenName", "")
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = Val(PartyDevice.Value) + psStartPartyDevices
                    Pair.Status = "Add " & DLNAFriendlyName
                    Pair.Render = Enums.CAPIControlType.Button
                    Pair.Render_Location.Row = PartyIndex
                    Pair.Render_Location.Column = 1
                    PartyIndex += 1
                    hs.DeviceVSP_AddPair(HSRefParty, Pair)
                End If
            Next
        Catch ex As Exception
            Log("Error in SonyUpdatePartyButtons for UPnPDevice = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TreatSetIOExSony(ButtonValue As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExSony called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            Case psRemoteOff  ' Remote Off
                SetAdministrativeStateRemote(False)
            Case psRemoteOn  ' Remote On
                SetAdministrativeStateRemote(True)
            Case psRegister  ' Register
                If SonyRegister(MySonyRegisterURL, False) Then
                    SetAdministrativeStateRemote(True)
                End If
            Case psCreateRemoteButtons
                CreateHSSonyRemoteButtons(True)
                CreateRemoteButtons(HSRefRemote)
            Case psWOL
                SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), PlugInIPAddress, GetSubnetMask())
            Case Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExSony called for UPnPDevice = " & MyUPnPDeviceName & " and Buttonvalue = " & ButtonValue & " and Registration = " & GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) & " and DeviceState = " & DeviceStatus, LogType.LOG_TYPE_INFO)
                If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) And UCase(DeviceStatus) = "ONLINE" Then
                    Dim objRemoteFile As String = gRemoteControlPath
                    Dim ButtonInfoString As String = GetStringIniFile(MyUDN, ButtonValue.ToString, "", objRemoteFile)
                    Dim ButtonInfos As String()
                    ButtonInfos = Split(ButtonInfoString, ":;:-:")
                    If UBound(ButtonInfos, 1) > 2 Then
                        SonyX_SendIRCC(ButtonInfos(1))
                    End If
                End If
        End Select
    End Sub

    Private Sub TreatSetIOExSonyParty(ButtonValue As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExSonyParty called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            Case psPartyAll  ' "Party All"
                SonyPartyAll()
            Case psExitParty  ' "Exit Party"
                SonyExitParty()
            Case Else
                ' the counting start here from StartPartyDevicesIndex
                SonyAddPartyByNumber(ButtonValue - psStartPartyDevices)
        End Select
    End Sub

    Public Function SonyX_SendIRCC(IRCCCode As String) As String
        SonyX_SendIRCC = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyX_SendIRCC called for device " & MyUPnPDeviceName & " and IRCCCode = " & IRCCCode.ToString, LogType.LOG_TYPE_INFO)
        If MySonyRegisterMode = "JSON" Then
            ' check whether we are still authenticated
            If GetBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False) Then
                Try
                    Dim CookieExpery As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyCookieExpiryDate.ToString, "")
                    If CookieExpery <> "" Then
                        Dim MYCookieExpiryDate As DateTime = DateTime.Parse(CookieExpery)
                        Dim ExperyRes As Integer = DateTime.Compare(DateTime.Now, MYCookieExpiryDate)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyX_SendIRCC called for device " & MyUPnPDeviceName & " and IRCCCode = " & IRCCCode.ToString & " compared date = " & CookieExpery & " with TimeNow and result = " & ExperyRes.ToString, LogType.LOG_TYPE_INFO)
                        If ExperyRes > 0 Then
                            SendJSONAuthentication(GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, ""))
                        End If
                    Else
                        SendJSONAuthentication(GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyAuthenticationPIN.ToString, ""))
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyX_SendIRCC for device - " & MyUPnPDeviceName & " parsing the Cookie Expiry Date = " & GetStringIniFile(MyUDN, DeviceInfoIndex.diSonyCookieExpiryDate.ToString, "") & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in SonyX_SendIRCC called for device " & MyUPnPDeviceName & " and IRCCCode = " & IRCCCode.ToString & " You need to Authenticate/Register first!!!", LogType.LOG_TYPE_WARNING)
            End If
        End If

        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = IRCCCode
            RemoteControlService.InvokeAction("X_SendIRCC", InArg, OutArg, MyAuthenticationCookieContainer)
            SonyX_SendIRCC = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyX_SendIRCC for device = " & MyUPnPDeviceName & " and IRCCCode = " & IRCCCode.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyX_GetStatus(CategoryCode As String) As String
        SonyX_GetStatus = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyX_GetStatus called for device " & MyUPnPDeviceName & " and CategoryCode = " & CategoryCode.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(1)
            InArg(0) = CategoryCode
            RemoteControlService.InvokeAction("X_GetStatus", InArg, OutArg)
            MyCurrentSonyRemoteStatus = OutArg(0)
            MyCurrentSonyCommandInfo = OutArg(1)
            SonyX_GetStatus = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyX_GetStatus for device = " & MyUPnPDeviceName & " and CategoryCode = " & CategoryCode.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_GetDeviceInfo() As String
        SonyPartyX_GetDeviceInfo = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_GetDeviceInfo called for device " & MyUPnPDeviceName & " and DeviceStatus = " & DeviceStatus.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(1)
            SonyPartyService.InvokeAction("X_GetDeviceInfo", InArg, OutArg)
            MyCurrentSonySingerCapability = OutArg(0)
            MyCurrentSonyTransportPort = OutArg(1)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_GetDeviceInfo for device = " & MyUPnPDeviceName & " retrieved SingerCapability = " & MyCurrentSonySingerCapability.ToString & " and TransportPort = " & MyCurrentSonyTransportPort.ToString, LogType.LOG_TYPE_INFO)
            SonyPartyX_GetDeviceInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_GetDeviceInfo for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_GetState() As String
        SonyPartyX_GetState = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_GetState called for device " & MyUPnPDeviceName & " and DeviceStatus = " & DeviceStatus.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(7)
            SonyPartyService.InvokeAction("X_GetState", InArg, OutArg)
            MyCurrentSonyPartyState = OutArg(0)
            MyCurrentSonyPartyMode = OutArg(1)
            MyCurrentSonyPartySong = OutArg(2)
            MyCurrentSonySessionID = OutArg(3)
            MyCurrentSonyNumberOfListeners = OutArg(4)
            MyCurrentSonyListenerList = OutArg(5)
            MyCurrentSonySingerUUID = OutArg(6)
            MyCurrentSonySingerSessionID = OutArg(7)
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonyPartyState        = " & MyCurrentSonyPartyState, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonyPartyMode         = " & MyCurrentSonyPartyMode, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonyPartySong         = " & MyCurrentSonyPartySong, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonySessionID         = " & MyCurrentSonySessionID.ToString, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonyNumberOfListeners = " & MyCurrentSonyNumberOfListeners.ToString, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonyListenerList      = " & MyCurrentSonyListenerList, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonySingerUUID        = " & MyCurrentSonySingerUUID, LogType.LOG_TYPE_INFO)
                Log("SonyPartyX_GetState for device = " & MyUPnPDeviceName & " retrieved SonySingerSessionID   = " & MyCurrentSonySingerSessionID.ToString, LogType.LOG_TYPE_INFO)
            End If
            SonyPartyX_GetState = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_GetState for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Start(PartyMode As String, ListenerList As String) As String
        ' allowed values for X_PartyMode are : IDLE, PARTY, LINK
        ' allowed values for X_PartyState are : NOT_READY, IDLE, SINGING, LISTENING 
        ' call this first with Partymode "party" or perhaps "link" and listenerlist empty
        ' Listenerlist = uuid:5f9ec1b3-ed59-1900-4530-0007f5238374
        '                uuid:00000000-0000-1010-8000-30f9ed246323
        SonyPartyX_Start = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Start called for device " & MyUPnPDeviceName & " and PartyMode = " & PartyMode.ToString & " and ListenerList = " & ListenerList.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = PartyMode
            InArg(1) = ListenerList
            SonyPartyService.InvokeAction("X_Start", InArg, OutArg)
            MyCurrentSonySingerSessionID = OutArg(0)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Start for device = " & MyUPnPDeviceName & " received SingerSessionID = " & MyCurrentSonySingerSessionID.ToString, LogType.LOG_TYPE_INFO)
            SonyPartyX_Start = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Start for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Entry(SingerSessionID As Integer, ListenerList As String) As String
        SonyPartyX_Entry = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Entry called for device " & MyUPnPDeviceName & " and SingerSessionID = " & SingerSessionID.ToString & " and ListenerList = " & ListenerList.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = SingerSessionID
            InArg(1) = ListenerList
            SonyPartyService.InvokeAction("X_Entry", InArg, OutArg)
            SonyPartyX_Entry = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Entry for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Leave(SingerSessionID As Integer, ListenerList As String) As String
        ' works on the singer with the listenerUUID and SingerSessionID = SingersessionID on the listener side
        SonyPartyX_Leave = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Leave called for device " & MyUPnPDeviceName & " and SingerSessionID = " & SingerSessionID.ToString & " and ListenerList = " & ListenerList.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = SingerSessionID
            InArg(1) = ListenerList
            SonyPartyService.InvokeAction("X_Leave", InArg, OutArg)
            SonyPartyX_Leave = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Leave for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Abort(SingerSessionID As Integer) As String
        ' Use this to go back to Idle after using x-start
        ' Only when singer!
        SonyPartyX_Abort = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Abort called for device " & MyUPnPDeviceName & " and SingerSessionID = " & SingerSessionID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = SingerSessionID
            SonyPartyService.InvokeAction("X_Abort", InArg, OutArg)
            SonyPartyX_Abort = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Abort for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Invite(PartyMode As String, SingerUUID As String, SingerSessionID As Integer) As String
        ' allowed values for X_PartyMode are : IDLE, PARTY, LINK
        SonyPartyX_Invite = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Invite called for device " & MyUPnPDeviceName & " and PartyMode = " & PartyMode.ToString & " and SingerUUID = " & SingerUUID.ToString & " and SingerSessionID = " & SingerSessionID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = PartyMode
            InArg(1) = SingerUUID
            InArg(2) = SingerSessionID
            SonyPartyService.InvokeAction("X_Invite", InArg, OutArg)
            MyCurrentSonyListenerSessionID = OutArg(0)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Invite for device = " & MyUPnPDeviceName & " received ListenerSessionID = " & MyCurrentSonyListenerSessionID.ToString, LogType.LOG_TYPE_INFO)
            SonyPartyX_Invite = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Invite for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SonyPartyX_Exit(ListenerSessionID As Integer) As String
        ' Use this to go back to Idle after using x-start
        ' works on listener with listenerSessionID = sessionID but leaves Singer not updated
        SonyPartyX_Exit = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SonyPartyX_Exit called for device " & MyUPnPDeviceName & " and ListenerSessionID = " & ListenerSessionID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = ListenerSessionID
            SonyPartyService.InvokeAction("X_Exit", InArg, OutArg)
            SonyPartyX_Exit = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SonyPartyX_Exit for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

End Class

<Serializable()> _
Public Class SonyJSONRemoteControllerInfo
    Public id As Integer
    Public result As Object()
End Class

<Serializable()> _
Public Class SonyJSONSystemSupportedFunction
    Public result As Object()
    Public id As Integer
End Class