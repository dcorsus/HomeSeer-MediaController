Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Xml
Imports System.Drawing
Imports System.IO

Partial Public Class HSPI

    Private MyApplicationURL As String = ""

    ' http://sdkdocs.roku.com/display/sdkdoc/External+Control+Guide
    ' https://sdkdocs.roku.com/display/sdkdoc/External+Control+API#ExternalControlAPI-KeypressKeyValues
    ' https://sdkdocs.roku.com/display/sdkdoc/External+Control+API#ExternalControlAPI-keypress/key

    'The External Control Service is a simple, RESTful service accessed via the http protocol on port 8060. Once you have the Roku ipaddress to connect to, you can issue the following External Control commands to the Roku:

    'query/apps This 'query/apps' returns a map of all the channels installed on the Roku box paired with their app id. This command is accessed via an http GET.
    'keydown is followed by a slash and the name of the key pressed. Keydown is equivalent to pressing down the remote key whose value is the argument passed. This command is sent via a POST with no body.
    'keyup is followed by a slash and the name of the key to release. Keyup is equivalent to releasing the remote key whose value is the argument passed. This command is sent via a POST with no body.
    'keypress is followed by a slash and the name of the key pressed. Keypress is equivalent to pressing down and releasing the remote key whose value is the argument passed. This command is sent via a POST with no body.
    'launch is followed by a slash and an app id, optionally followed by a question mark and a list of URL parameters that are sent to the app id as an roAssociativeArray passed to the RunUserInterface() or Main() entry point. This command is sent via a POST with no body.
    'query/icon is followed by a slash and an app id and returns an icon corresponding to that app. The binary data with an identifying MIME-type header is returned. This command is accessed via an http GET.

    'Example: GET /query/icon/12

    'input enables a developer to send custom events to their Brightscript application. It takes a user defined list of name-value pairs sent as query string uri parameters. The external control server places these name-value pairs into a BrightScript associative array and passes them directly through to the currently executing channel script via a Message Port attached to a created roInput object. Please refer to Section 3.1 below for more detailed recommendations on how to pass your data. Messages of type roInputEvent have a GetInfo() method that will obtain the associative array. The arguments must be URL-encoded. This command is sent via a POST with no body.

    'Example: POST /input?acceleration.x=0.0&acceleration.y=0.0&acceleration.z=9.8


    ' POSTs
    ' /keypress/Home
    ' /keydown/Left
    ' /keyup/Left
    ' /launch/dev?url=http%3A%2F%2Fvideo.ted.com%2Ftalks%2Fpodcast%2FVilayanurRamac\handran_2007_480.mp4&streamformat=mp4
    ' /launch/dev?contentID=my_content_id&options=my_options
    ' /launch/11?contentID=14 brings you to the summary screen of the channel aka Channel Store
    ' /launch/12 gets to launch the channel directly
    ' /query/icon/12  -> returns icon for appid = 12 = Netflix in my case
    ' 8060/query/device-info

    ' Send Keys via Keypress or Keydown
    ' Home
    ' Rev
    ' Fwd
    ' Play
    ' Select
    ' Left
    ' Right
    ' Down
    ' Up
    ' Back
    ' InstantReplay
    ' Info
    ' Backspace
    ' Search
    ' Enter
    ' Lit_  -> add here the ascii character you want to enter


    'POST /keydown/Down HTTP/1.1
    'Host: 192.168.1.148:8060
    'User-Agent: Mozilla/5.0 (Windows NT 6.1; WOW64; rv:27.0) Gecko/20100101 Firefox/27.0
    'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8
    'Accept-Language: en-US,en;q=0.5
    'Accept-Encoding: gzip, deflate
    'Referer: http://remoku.tv/
    'Connection: keep-alive
    'Content-Type: application/x-www-form-urlencoded
    'Content-Length: 0

    Private Function RokuRetrieveDIALAppList(URLDoc As String) As Boolean

        ' GET http://192.168.1.111:8060/query/apps
        '<apps>
        '   <app id="5127" version="1.0.28">Roku Spotlight</app>
        '   <app id="11" version="2.6.6">Roku Channel Store</app>
        '   <app id="12" version="3.1.21">Netflix</app>
        '   <app id="13" version="3.2.7">Amazon Instant Video</app>
        '   <app id="2285" version="2.1.1">Hulu Plus</app>
        '   <app id="28" version="2.0.23">Pandora</app>
        '   <app id="6119" version="1.2.0">Popcornflix</app>
        '   <app id="9161" version="1.1.1">Tested</app>
        '</apps>
        RokuRetrieveDIALAppList = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList called for device - " & MyUPnPDeviceName & " with URL = " & URLDoc.ToString, LogType.LOG_TYPE_INFO)
        If String.IsNullOrEmpty(URLDoc) Then Exit Function
        ' the URL to the DIAL information is returned in the header with a header field = "Application-URL"
        ' So first do get on the Location Info
        Dim RequestUri As Uri = Nothing
        Try
            RequestUri = New Uri(URLDoc)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            If MyAuthenticationCookieContainer IsNot Nothing Then
                wRequest.CookieContainer = MyAuthenticationCookieContainer
            Else
                MyAuthenticationCookieContainer = New CookieContainer
                Try
                    If System.IO.File.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json") Then
                        Dim Stream As Stream = System.IO.File.Open(CurrentAppPath & "/html/" & tIFACE_NAME & "/" & MyUDN & "_Authenticationcookie.json", FileMode.Open)
                        Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                        MyAuthenticationCookieContainer = formatter.Deserialize(Stream)
                        wRequest.CookieContainer = MyAuthenticationCookieContainer
                        Stream.Dispose()
                    End If
                Catch ex As Exception
                    Log("Error in RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " restoring the Cookie Container with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    MyAuthenticationCookieContainer = Nothing
                End Try
            End If

            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Try
                MyApplicationURL = webResponse.Headers("Application-URL")
                If MyApplicationURL <> "" Then
                    URLDoc = MyApplicationURL
                Else '
                    MyApplicationURL = URLDoc
                End If
                If PIDebuglevel > DebugLevel.dlEvents Then Log(" RokuRetrieveDIALAppList for device = " & MyUPnPDeviceName & " found Application-URL = " & MyApplicationURL.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                MyApplicationURL = URLDoc
            End Try
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & URLDoc.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyApplicationURL = URLDoc
            Exit Function
        End Try
        RequestUri = Nothing

        Dim xmlDoc As New XmlDocument
        Dim DocURL As Uri = Nothing
        Try
            DocURL = New Uri(URLDoc)
        Catch ex As Exception
        End Try

        ' added in v.49. Retrieve the capabilities so we can add buttons for TVs etc
        Try
            If DocURL IsNot Nothing Then
                RequestUri = New Uri(DocURL, "query/device-info")
            Else
                RequestUri = New Uri(URLDoc & "/query/device-info")
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList called for device - " & MyUPnPDeviceName & " is retrieving the device-info with URL = " & RequestUri.ToString, LogType.LOG_TYPE_INFO)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As System.Net.HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            xmlDoc.Load(webStream)
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & RequestUri.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            xmlDoc = Nothing
            'Exit Function
        End Try

        ' info looks like this
        '<device-info> 
        '   <udn>29780014-180c-1056-80cf-30d16b133d9f</udn>
        '   <serial-number>YN00D6808655</serial-number>
        '   <device-id>9G487D808655</device-id>
        '   <advertising-id>f92d29d8-967d-5029-ade3-02fe351eb170</advertising-id>
        '   <vendor-name>Sharp</vendor-name>
        '   <model-name>LC-43LBU591U</model-name>
        '   <model-number>7201X</model-number>
        '   <model-region>US</model-region>
        '   <Is-tv>true</Is-tv>
        '   <Is-stick>false</Is-stick>
        '   <screen-size>43</screen-size>
        '   <panel-id>0</panel-id>
        '   <tuner-type>ATSC</tuner-type>
        '   <supports-ethernet>true</supports-ethernet>
        '   <wifi-mac>30:d1 : 6b:13:3d:9f</wifi-mac>
        '   <wifi-driver>realtek</wifi-driver>
        '   <ethernet-mac>a8:82:00:aa:61:29</ethernet-mac>
        '   <network-type>wifi</network-type>
        '   <network-name>xxxxx</network-name>
        '   <friendly-device-name>Jupiter Rockette</friendly-device-name>
        '   <friendly-model-name>Sharp•Roku TV</friendly-model-name>
        '   <default-device-name>Sharp•Roku TV - YN00D6808655</default-device-name>
        '   <user-device-name>Jupiter Rockette</user-device-name>
        '   <software-version>8.2.0</software-version>
        '   <software-build>4167</software-build>
        '   <secure-device>true</secure-device>
        '   <language>en</language>
        '   <country>US</country>
        '   <locale>en_US</locale>
        '   <time-zone-auto>True</time-zone-auto>
        '   <time-zone>US/Pacific</time-zone>
        '   <time-zone-name>United States/Pacific</time-zone-name>
        '   <time-zone-tz>America/Los_Angeles</time-zone-tz>
        '   <time-zone-offset>-480</time-zone-offset>
        '   <clock-format>12-hour</clock-format>
        '   <uptime>611055</uptime>
        '   <power-mode>PowerOn</power-mode>
        '   <supports-suspend>True</supports-suspend>
        '   <supports-find-remote>False</supports-find-remote>
        '   <supports-audio-guide>True</supports-audio-guide>
        '   <supports-rva>True</supports-rva>
        '   <developer-enabled>False</developer-enabled>
        '   <keyed-developer-id/>
        '   <search-enabled>True</search-enabled>
        '   <search-channels-enabled>True</search-channels-enabled>
        '   <voice-search-enabled>True</voice-search-enabled>
        '   <notifications-enabled>True</notifications-enabled>
        '   <notifications-first-use>True</notifications-first-use>
        '   <supports-Private-listening>True</supports-Private-listening>
        '   <supports-Private-listening-dtv>True</supports-Private-listening-dtv>
        '   <supports-warm-standby>True</supports-warm-standby>
        '   <headphones-connected>False</headphones-connected>
        '   <expert-pq-enabled>0.5</expert-pq-enabled>
        '   <supports-ecs-textedit>True</supports-ecs-textedit>
        '   <supports-ecs-microphone>True</supports-ecs-microphone>
        '   <supports-wake-On-wlan>True</supports-wake-On-wlan>
        '   <has-play-On-roku>True</has-play-On-roku>
        '   <has-mobile-screensaver>True</has-mobile-screensaver>
        '   <support-url>www.sharptvusa.com/support</support-url>
        '</device-info>

        Dim isTV As Boolean = False

        Try
            If xmlDoc IsNot Nothing Then
                Dim DeviceInfo As XmlNode = xmlDoc.SelectSingleNode("device-info")
                If DeviceInfo IsNot Nothing Then
                    Dim TagName As String = ""
                    Dim TagContent As String = ""
                    For Each Entry As XmlNode In DeviceInfo.ChildNodes
                        TagName = Entry.Name
                        TagContent = Entry.InnerText
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " found device-info TagName = " & TagName & ", TagContent = " & TagContent, LogType.LOG_TYPE_INFO)
                        If TagName.ToLower = "Is-tv" Then
                            If TagContent.ToLower = "true" Then isTV = True
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " found isTV", LogType.LOG_TYPE_INFO)
                        ElseIf TagName.ToLower = "ethernet-mac" Then
                            WriteStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, TagContent.Replace("-", ":"))
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " found eMAC = " & TagContent, LogType.LOG_TYPE_INFO)
                        ElseIf TagName.ToLower = "wifi-mac" Then
                            WriteStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, TagContent)
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " found WifiMAC = " & TagContent, LogType.LOG_TYPE_INFO)

                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RokuRetrieveDIALAppList for UPnPDevice = " & MyUPnPDeviceName & "  processing device-info XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If DocURL IsNot Nothing Then
                RequestUri = New Uri(DocURL, "query/apps")
            Else
                RequestUri = New Uri(URLDoc & "/query/apps")
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList called for device - " & MyUPnPDeviceName & " is retrieving the app-info with URL = " & RequestUri.ToString, LogType.LOG_TYPE_INFO)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As System.Net.HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            xmlDoc.Load(webStream)
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & RequestUri.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            xmlDoc = Nothing
            Exit Function
        End Try


        Dim objRemoteFile As String = gRemoteControlPath
        'WriteStringIniFile(MyUDN & " - Default Codes", "Home", "/keypress/Home" & ":;:-:20", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Rev", "/keypress/Rev" & ":;:-:21", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Fwd", "/keypress/Fwd" & ":;:-:22", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Play", "/keypress/Play" & ":;:-:23", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Select", "/keypress/Select" & ":;:-:24", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Left", "/keypress/Left" & ":;:-:25", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Right", "/keypress/Right" & ":;:-:26", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Down", "/keypress/Down" & ":;:-:27", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Up", "/keypress/Up" & ":;:-:28", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "InstantReplay", "/keypress/InstantReplay" & ":;:-:29", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Backspace", "/keypress/Backspace" & ":;:-:30", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Search", "/keypress/Search" & ":;:-:31", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Enter", "/keypress/Enter" & ":;:-:32", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Info", "/keypress/Info" & ":;:-:32", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Enter", "/keyDown/Back" & ":;:-:32", objRemoteFile)
        'WriteStringIniFile(MyUDN & " - Default Codes", "Enter", "/keyUp/Back" & ":;:-:32", objRemoteFile)
        ' Back
        ' Home
        ' Right
        ' Info  <- this is the options key
        ' InstantReplay <- this is the return or half circle back button
        ' /keypress/LIT_d when "d" is pressed
        ' Rev
        ' Play
        ' Fwd
        ' Back <- this is the back arrow 
        ' Select
        ' Up
        ' Left
        ' Down

        WriteStringIniFile(MyUDN, "20", "Home" & ":;:-:" & "/keypress/Home" & ":;:-:1:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "21", "Rev" & ":;:-:" & "/keypress/Rev" & ":;:-:1:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "22", "Fwd" & ":;:-:" & "/keypress/Fwd" & ":;:-:1:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "23", "Play" & ":;:-:" & "/keypress/Play" & ":;:-:1:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "24", "Select" & ":;:-:" & "/keypress/Select" & ":;:-:2:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "25", "Left" & ":;:-:" & "/keypress/Left" & ":;:-:2:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "26", "Right" & ":;:-:" & "/keypress/Right" & ":;:-:2:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "27", "Down" & ":;:-:" & "/keypress/Down" & ":;:-:2:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "28", "Up" & ":;:-:" & "/keypress/Up" & ":;:-:3:;:-:1", objRemoteFile)
        WriteStringIniFile(MyUDN, "29", "InstantReplay" & ":;:-:" & "/keypress/InstantReplay" & ":;:-:3:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "30", "Backspace" & ":;:-:" & "/keypress/Backspace" & ":;:-:3:;:-:3", objRemoteFile)
        WriteStringIniFile(MyUDN, "31", "Search" & ":;:-:" & "/keypress/Search" & ":;:-:3:;:-:4", objRemoteFile)
        WriteStringIniFile(MyUDN, "32", "Enter" & ":;:-:" & "/keypress/Enter" & ":;:-:4:;:-:1", objRemoteFile)
        ' added 11/9/2017
        WriteStringIniFile(MyUDN, "33", "Info" & ":;:-:" & "/keypress/Info" & ":;:-:4:;:-:2", objRemoteFile)
        WriteStringIniFile(MyUDN, "34", "Back" & ":;:-:" & "/keypress/Back" & ":;:-:4:;:-:3", objRemoteFile)
        'Roku Devices that support the "Find Remote" support

        'FindRemote

        'Note that query/device-info includes a supports-find-remote flag that indicates whether the Roku device supports FindRemote.
        'However this does Not specifically indicate that the device has a paired remote that supports "Find remote" as well.

        'Some Roku devices, such As Roku TVs, also support
        '   VolumeDown()
        '   VolumeMute
        '   VolumeUP()
        '   PowerOff

        If isTV Then
            WriteStringIniFile(MyUDN, "35", "Power Off" & ":;:-:" & "/keypress/PowerOff" & ":;:-:5:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "36", "VolDown" & ":;:-:" & "/keypress/VolumeDown" & ":;:-:5:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "37", "VolUP" & ":;:-:" & "/keypress/VolumeUP" & ":;:-:5:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "38", "Mute" & ":;:-:" & "/keypress/VolumeMute" & ":;:-:5:;:-:4", objRemoteFile)
            'Roku TV devices also support changing the channel when watching the TV tuner input
            '   ChannelUp
            '   ChannelDown
            WriteStringIniFile(MyUDN, "39", "ChannelUp" & ":;:-:" & "/keypress/ChannelUp" & ":;:-:6:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "40", "ChannelDown" & ":;:-:" & "/keypress/ChannelDown" & ":;:-:6:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "41", "InputTuner" & ":;:-:" & "/keypress/InputTuner" & ":;:-:6:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "42", "InputHDMI1" & ":;:-:" & "/keypress/InputHDMI1" & ":;:-:6:;:-:4", objRemoteFile)
            'Roku TV devices also support keys to set the current TV input UI
            '   InputTuner
            '   InputHDMI1
            '   InputHDMI2
            '   InputHDMI3
            '   InputHDMI4
            '   InputAV1
            WriteStringIniFile(MyUDN, "43", "InputHDMI2" & ":;:-:" & "/keypress/InputHDMI2" & ":;:-:7:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "44", "InputHDMI3" & ":;:-:" & "/keypress/InputHDMI3" & ":;:-:7:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "45", "InputHDMI4" & ":;:-:" & "/keypress/InputHDMI4" & ":;:-:7:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "46", "InputAV1" & ":;:-:" & "/keypress/InputAV1" & ":;:-:7:;:-:4", objRemoteFile)
        End If

        Dim ButtonIndex As Integer = 50
        Dim RowIndex As Integer = 8
        Dim ColumnIndex As Integer = 1
        Dim ChannelImage As Image

        If xmlDoc.HasChildNodes Then
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlDoc.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " Document root node: " & xmlDoc.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    Dim AppName As String = ""
                    Dim AppID As String = ""
                    'AppName = outerNode.Attributes("app").InnerText
                    AppName = outerNode.InnerText
                    AppID = outerNode.Attributes("id").Value
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & "------> App Name: " & AppName, LogType.LOG_TYPE_INFO)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & "------> App ID: " & AppID, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN & " - Default Codes", AppName, "/launch/" & AppID & ":;:-:" & ButtonIndex.ToString, objRemoteFile)
                    WriteStringIniFile(MyUDN, ButtonIndex.ToString, AppName & ":;:-:" & "/launch/" & AppID & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString, objRemoteFile)
                    Try
                        ChannelImage = GetPicture(URLDoc & "/query/icon/" & AppID.ToString)
                        If Not ChannelImage Is Nothing Then
                            Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                            Dim SuccesfullSave As Boolean = False
                            SuccesfullSave = hs.WriteHTMLImage(ChannelImage, FileArtWorkPath & "RokuChannelImage_" & AppID.ToString & ".png", True)
                            If Not SuccesfullSave Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RokuRetrieveDIALAppList for Device = " & MyUPnPDeviceName & " had error storing Roku Channel Image at " & FileArtWorkPath & "RokuChannelImage_" & AppID.ToString & ".png", LogType.LOG_TYPE_ERROR)
                            Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for Device = " & MyUPnPDeviceName & " stored Roku Channel Image at " & FileArtWorkPath & "RokuChannelImage_" & AppID.ToString & ".png", LogType.LOG_TYPE_INFO)
                            End If
                            ChannelImage.Dispose()
                        End If
                    Catch ex As Exception
                        Log("Error in RokuRetrieveDIALAppList retrieving the Roku Channel Image with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    ButtonIndex = ButtonIndex + 1
                    ColumnIndex = ColumnIndex + 1
                    If ColumnIndex > 4 Then
                        ColumnIndex = 1
                        RowIndex = RowIndex + 1
                    End If
                Next
                RokuRetrieveDIALAppList = True
            Catch ex As Exception
                Log("Error in RokuRetrieveDIALAppList for UPnPDevice = " & MyUPnPDeviceName & "  processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        RowIndex += 1
        ButtonIndex += 1
        ColumnIndex = 1

        ' tv-channels'
        Try
            If DocURL IsNot Nothing Then
                RequestUri = New Uri(DocURL, "query/tv-channels")
            Else
                RequestUri = New Uri(URLDoc & "/query/tv-channels")
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList called for device - " & MyUPnPDeviceName & " is retrieving the tv-channels with URL = " & RequestUri.ToString, LogType.LOG_TYPE_INFO)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As System.Net.HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            xmlDoc.Load(webStream)
            webStream.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " retrieving the document with URL = " & RequestUri.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            xmlDoc = Nothing
            Exit Function
        End Try

        Try
            If xmlDoc IsNot Nothing Then
                Dim ChannelsInfo As XmlNode = xmlDoc.SelectSingleNode("tv-channels")
                If ChannelsInfo IsNot Nothing Then
                    Dim TagName As String = ""
                    Dim TagContent As String = ""
                    For Each Channel As XmlNode In ChannelsInfo.ChildNodes
                        Dim ChannelName As String = ""
                        Dim ChannelNumber As String = ""
                        Dim ChannelType As String = ""
                        Dim ChannelUserHidden As String = ""
                        TagName = Channel.Name
                        TagContent = Channel.InnerText
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & " found channel-info TagName = " & TagName & ", TagContent = " & TagContent, LogType.LOG_TYPE_INFO)
                        If TagName.ToLower = "number" Then
                            ChannelNumber = TagContent
                        ElseIf TagName.ToLower = "name" Then
                            ChannelName = TagContent
                        ElseIf TagName.ToLower = "type" Then
                            ChannelType = TagContent
                        ElseIf TagName.ToLower = "user-hidden" Then
                            ChannelUserHidden = TagContent
                        End If
                        If ChannelName <> "" And ChannelNumber = "" Then
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & "------> Channel Name: " & ChannelName, LogType.LOG_TYPE_INFO)
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("RokuRetrieveDIALAppList for device - " & MyUPnPDeviceName & "------> Channel Number: " & ChannelName, LogType.LOG_TYPE_INFO)
                            WriteStringIniFile(MyUDN & " - Default Codes", ChannelName, "/launch/tvinput.dtv?ch=" & ChannelName & ":;:-:" & ButtonIndex.ToString, objRemoteFile)
                            WriteStringIniFile(MyUDN, ButtonIndex.ToString, ChannelName & ":;:-:" & "/launch/tvinput.dtv?ch=" & ChannelName & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString, objRemoteFile)
                            ButtonIndex = ButtonIndex + 1
                            ColumnIndex = ColumnIndex + 1
                            If ColumnIndex > 4 Then
                                ColumnIndex = 1
                                RowIndex = RowIndex + 1
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RokuRetrieveDIALAppList for UPnPDevice = " & MyUPnPDeviceName & "  processing channel-info XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        xmlDoc = Nothing

    End Function


    Private Sub CreateHSRokuRemoteButtons(ReCreate As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSRokuRemoteButtons called for device - " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
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

    Private Sub TreatSetIOExRoku(ButtonValue As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExRoku called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            'Case psRemoteOff ' Remote Off
            '    SetAdministrativeStateRemote(False)
            'Case psRemoteOn
            '    SetAdministrativeStateRemote(True)
            Case psCreateRemoteButtons
                CreateHSRokuRemoteButtons(True)
                CreateRemoteButtons(HSRefRemote)
            Case Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExRoku called for UPnPDevice = " & MyUPnPDeviceName & " and Buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
                If UCase(DeviceStatus) = "ONLINE" Then
                    Dim objRemoteFile As String = gRemoteControlPath
                    Dim ButtonInfoString As String = GetStringIniFile(MyUDN, ButtonValue.ToString, "", objRemoteFile)
                    Dim ButtonInfos As String()
                    ButtonInfos = Split(ButtonInfoString, ":;:-:")
                    If UBound(ButtonInfos, 1) > 2 Then
                        PostToRoku(ButtonInfos(1))
                    End If
                End If
        End Select


    End Sub

    Private Sub PostToRoku(KeyCode As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostToRoku called for device - " & MyUPnPDeviceName & " with KeyCode = " & KeyCode.ToString, LogType.LOG_TYPE_INFO)
        If KeyCode = "" Then Exit Sub
        Dim RequestUri = New Uri(MyApplicationURL & KeyCode) ' need to fix this to take / into account 
        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "POST"
            ' Set the ContentType property of the WebRequest.
            wRequest.ContentType = "application/x-www-form-urlencoded"
            ' Set the ContentLength property of the WebRequest.
            wRequest.ContentLength = 0
            wRequest.KeepAlive = False
            ' Get the request stream.
            'Dim dataStream As Stream = wRequest.GetRequestStream()
            ' Write the data to the request stream.
            'dataStream.Write(byteArray, 0, byteArray.Length)
            ' Close the Stream object.
            'dataStream.Close()
            ' Get the response.
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim webStream As Stream = Nothing
            webStream = webResponse.GetResponseStream
            webStream.Close()
            webResponse.Close()
            'dataStream.Close()
        Catch ex As Exception
            Log("Error in PostToRoku for device - " & MyUPnPDeviceName & " with KeyCode = " & KeyCode.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
    End Sub

End Class