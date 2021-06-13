Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.IO
Imports System.Xml
Imports System.Web.Script.Serialization
Imports Microsoft.Win32.SafeHandles
Imports System.Runtime.InteropServices
Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Numerics


'Imports Pairing

Partial Public Class HSPI

    Dim MySamsungAsyncSocket As AsynchronousClient = Nothing
    Dim MySamsungWebSocket As WebSocketClient = Nothing
    Dim MySamsungClient As Socket
    Dim WebSocketRequest As HttpWebRequest = Nothing
    Dim SamsungAESKey As Byte() = Nothing
    Dim SamsungSessionID As String = ""
    Dim SamsungDeviceID As String = "MediaControllerPI"
    Dim SamsungPIN As String = ""

    Friend WithEvents myRetrieveInitInfoTimer As Timers.Timer

    Dim wbKey As String = "abbb120c09e7114243d1fa0102163b27"
    Dim transKey As String = "6c9474469ddf7578f3e5ad8a4c703d99"
    Dim publicKey As String = "2cb12bb2cbf7cec713c0fff7b59ae68a96784ae517f41d259a45d20556177c0ffe951ca60ec03a990c9412619d1bee30adc7773088c5721664cffcedacf6d251cb4b76e2fd7aef09b3ae9f9496ac8d94ed2b262eee37291c8b237e880cc7c021fb1be0881f3d0bffa4234d3b8e6a61530c00473ce169c025f47fcc001d9b8051"
    Dim privateKey As String = "2fd6334713816fae018cdee4656c5033a8d6b00e8eaea07b3624999242e96247112dcd019c4191f4643c3ce1605002b2e506e7f1d1ef8d9b8044e46d37c0d5263216a87cd783aa185490436c4a0cb2c524e15bc1bfeae703bcbc4b74a0540202e8d79cadaae85c6f9c218bc1107d1f5b4b9bd87160e782f4e436eeb17485ab4d"
    Dim prime As String = "b361eb0ab01c3439f2c16ffda7b05e3e320701ebee3e249123c3586765fd5bf6c1dfa88bb6bb5da3fde74737cd88b6a26c5ca31d81d18e3515533d08df619317063224cf0943a2f29a5fe60c1c31ddf28334ed76a6478a1122fb24c4a94c8711617ddfe90cf02e643cd82d4748d6d4a7ca2f47d88563aa2baf6482e124acd7dd"

    Const BLOCK_SIZE As Integer = 16
    Const SHA_DIGEST_LENGTH As Integer = 20

    'Private Declare Sub applySamyGOKeyTransform Lib "SamsungCrypto.dll" (ByRef Input As Byte(), ByRef Output As Byte())    ' removed 5/25/2019, not sure what this was doing here?
    Private Declare Sub applySamyGOKeyTransform Lib "MediaControllerCrypto.dll" (ByRef Input As Byte(), ByRef Output As Byte())
    <DllImport("MediaControllerCrypto.dll", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub applySamyGOKeyTransform(Input As IntPtr, Output As IntPtr)
    End Sub

    Private Sub WriteSamsungKeyInfoToInfoFile()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("WriteSamsungKeyInfoToInfoFile called", LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = (gRemoteControlPath)
        WriteStringIniFile(MyUDN & " - Default Codes", "Power", "KEY_POWER" & ":;:-:20", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PowerOn", "KEY_POWERON" & ":;:-:21", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PowerOff", "KEY_POWEROFF" & ":;:-:22", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Source", "KEY_SOURCE" & ":;:-:23", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Mute", "KEY_MUTE" & ":;:-:24", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "1", "KEY_1" & ":;:-:25", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "2", "KEY_2" & ":;:-:26", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "3", "KEY_3" & ":;:-:27", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "4", "KEY_4" & ":;:-:28", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "5", "KEY_5" & ":;:-:29", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "6", "KEY_6" & ":;:-:30", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "7", "KEY_7" & ":;:-:31", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "8", "KEY_8" & ":;:-:32", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "9", "KEY_9" & ":;:-:33", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "0", "KEY_0" & ":;:-:34", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "11", "KEY_11" & ":;:-:35", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "12", "KEY_12" & ":;:-:36", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "-", "KEY_PLUS100" & ":;:-:37", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PRE-CH", "KEY_PRECH" & ":;:-:38", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "TV", "KEY_TV" & ":;:-:39", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Vol +", "KEY_VOLUP" & ":;:-:40", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Vol -", "KEY_VOLDOWN" & ":;:-:41", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Channel +", "KEY_CHUP" & ":;:-:42", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Channel -", "KEY_CHDOWN" & ":;:-:43", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Return", "KEY_RETURN" & ":;:-:44", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Up", "KEY_UP" & ":;:-:45", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Down", "KEY_DOWN" & ":;:-:46", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Left", "KEY_LEFT" & ":;:-:47", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Right", "KEY_RIGHT" & ":;:-:48", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Enter", "KEY_ENTER" & ":;:-:49", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Info", "KEY_INFO" & ":;:-:50", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Exit", "KEY_EXIT" & ":;:-:51", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Tools", "KEY_TOOLS" & ":;:-:52", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "CH List", "KEY_CHLIST" & ":;:-:53", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Red", "KEY_RED" & ":;:-:54", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Green", "KEY_GREEN" & ":;:-:55", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Blue", "KEY_BLUE" & ":;:-:56", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Yellow", "KEY_YELLOW" & ":;:-:57", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "REC", "KEY_REC" & ":;:-:58", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "RW", "KEY_RW" & ":;:-:59", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "FF", "KEY_FF" & ":;:-:60", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "InfoLink", "KEY_INFOLINK" & ":;:-:61", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "FavCH", "KEY_FAVCH" & ":;:-:62", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Content", "KEY_CONTENT" & ":;:-:63", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Menu", "KEY_MENU" & ":;:-:64", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "WLink", "KEY_WLINK" & ":;:-:65", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "CC", "KEY_CC" & ":;:-:66", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Guide", "KEY_GUIDE" & ":;:-:67", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Subtitle", "KEY_SUBTITLE" & ":;:-:68", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Aspect", "KEY_ASPECT" & ":;:-:69", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PIP On/Off", "KEY_PIP_ONOFF" & ":;:-:70", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PIP CH Up", "KEY_PIP_CHUP" & ":;:-:71", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PIP CH Down", "KEY_PIP_CHDOWN" & ":;:-:72", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "3D", "KEY_3D" & ":;:-:73", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "AD", "KEY_AD" & ":;:-:74", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "PMode", "KEY_PMODE" & ":;:-:75", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "SMode", "KEY_SMODE" & ":;:-:76", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Step", "KEY_STEP" & ":;:-:77", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Sleep", "KEY_SLEEP" & ":;:-:78", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Dolby SRR", "KEY_DOLBY_SRR" & ":;:-:79", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "MTS", "KEY_MTS" & ":;:-:80", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Disc Menu", "KEY_DISC_MENU" & ":;:-:81", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "EMode", "KEY_EMODE" & ":;:-:82", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "AV1", "KEY_AV1" & ":;:-:83", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "AV2", "KEY_AV2" & ":;:-:84", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "AV3", "KEY_AV3" & ":;:-:85", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "SVIDEO1", "KEY_SVIDEO1" & ":;:-:86", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "SVIDEO2", "SVIDEO2" & ":;:-:87", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "SVIDEO3", "KEY_SVIDEO3" & ":;:-:88", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "COMPONENT1", "KEY_COMPONENT1" & ":;:-:89", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "COMPONENT2", "KEY_COMPONENT2" & ":;:-:90", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "DVI", "KEY_DVI" & ":;:-:91", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "HDMI", "KEY_HDMI" & ":;:-:92", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "HDMI1", "KEY_HDMI1" & ":;:-:93", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "HDMI2", "KEY_HDMI2" & ":;:-:94", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "HDMI3", "KEY_HDMI3" & ":;:-:95", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "HDMI4", "KEY_HDMI4" & ":;:-:96", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Internet", "KEY_RSS" & ":;:-:97", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Play", "KEY_PLAY" & ":;:-:98", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Pause", "KEY_PAUSE" & ":;:-:99", objRemoteFile)
        WriteStringIniFile(MyUDN & " - Default Codes", "Stop", "KEY_STOP" & ":;:-:100", objRemoteFile)


        '    KEY_ADDDEL
        '    KEY_PIP_SWAP

        '    KEY_CAPTION
        '    KEY_PMODE
        '    KEY_TTX_MIX
        '    KEY_PICTURE_SIZE
        '    KEY_PIP_SIZE
        '    KEY_MAGIC_CHANNEL
        '    KEY_PIP_SCAN
        '    KEY_DEVICE_CONNECT
        '    KEY_HELP
        '    KEY_ANTENA
        '    KEY_CONVERGENCE
        '    KEY_AUTO_PROGRAM
        '    KEY_FACTORY
        '    KEY_3SPEED
        '    KEY_RSURF
        '    KEY_ASPECT
        '    KEY_TOPMENU
        '    KEY_GAME
        '    KEY_QUICK_REPLAY
        '    KEY_STILL_PICTURE
        '    KEY_DTV
        '    KEY_INSTANT_REPLAY
        '    KEY_LINK
        '    KEY_FF_
        '    KEY_GUIDE
        '    KEY_REWIND_
        '    KEY_ANGLE
        '    KEY_RESERVED1
        '    KEY_ZOOM1
        '    KEY_PROGRAM
        '    KEY_BOOKMARK
        '    KEY_DISC_MENU
        '    KEY_PRINT
        '    KEY_SUB_TITLE
        '    KEY_CLEAR
        '    KEY_VCHIP
        '    KEY_REPEAT
        '    KEY_DOOR
        '    KEY_OPEN
        '    KEY_WHEEL_LEFT
        '    KEY_TURBO
        '    KEY_FM_RADIO
        '    KEY_DVR_MENU
        '    KEY_PCMODE
        '    KEY_TTX_SUBFACE
        '    KEY_CH_LIST
        '    KEY_DNIe
        '    KEY_SRS
        '    KEY_CONVERT_AUDIO_MAINSUB
        '    KEY_MDC
        '    KEY_SEFFECT
        '    KEY_DVR
        '    KEY_DTV_SIGNAL
        '    KEY_LIVE
        '    KEY_PERPECT_FOCUS
        '    KEY_HOME
        '    KEY_ESAVING
        '    KEY_WHEEL_RIGHT
        '    KEY_CONTENTS
        '    KEY_VCR_MODE
        '    KEY_CATV_MODE
        '    KEY_DSS_MODE
        '    KEY_TV_MODE
        '    KEY_DVD_MODE
        '    KEY_STB_MODE
        '    KEY_CALLER_ID
        '    KEY_SCALE
        '    KEY_ZOOM_MOVE
        '    KEY_CLOCK_DISPLAY
        '    KEY_MAGIC_BRIGHT

        '    KEY_W_LINK
        '    KEY_DTV_LINK
        '    KEY_APP_LIST
        '    KEY_BACK_MHP
        '    KEY_ALT_MHP
        '    KEY_DNSe
        '    KEY_RSS
        '    KEY_ENTERTAINMENT
        '    KEY_ID_INPUT
        '    KEY_ID_SETUP
        '    KEY_ANYNET

        '    KEY_ANYVIEW
        '    KEY_MS
        '    KEY_MORE
        '    KEY_PANNEL_POWER
        '    KEY_PANNEL_CHUP
        '    KEY_PANNEL_CHDOWN
        '    KEY_PANNEL_VOLUP
        '    KEY_PANNEL_VOLDOW
        '    KEY_PANNEL_ENTER
        '    KEY_PANNEL_MENU
        '    KEY_PANNEL_SOURCE

        '    KEY_ZOOM2
        '    KEY_PANORAMA
        '    KEY_4_3
        '    KEY_16_9
        '    KEY_DYNAMIC
        '    KEY_STANDARD
        '    KEY_MOVIE1
        '    KEY_CUSTOM
        '    KEY_AUTO_ARC_RESET
        '    KEY_AUTO_ARC_LNA_ON
        '    KEY_AUTO_ARC_LNA_OFF
        '    KEY_AUTO_ARC_ANYNET_MODE_OK
        '    KEY_AUTO_ARC_ANYNET_AUTO_START
        '    KEY_AUTO_FORMAT
        '    KEY_DNET
        '    KEY_SETUP_CLOCK_TIMER
        '    KEY_AUTO_ARC_CAPTION_ON
        '    KEY_AUTO_ARC_CAPTION_OFF
        '    KEY_AUTO_ARC_PIP_DOUBLE
        '    KEY_AUTO_ARC_PIP_LARGE
        '    KEY_AUTO_ARC_PIP_SMALL
        '    KEY_AUTO_ARC_PIP_WIDE
        '    KEY_AUTO_ARC_PIP_LEFT_TOP
        '    KEY_AUTO_ARC_PIP_RIGHT_TOP
        '    KEY_AUTO_ARC_PIP_LEFT_BOTTOM
        '    KEY_AUTO_ARC_PIP_RIGHT_BOTTOM
        '    KEY_AUTO_ARC_PIP_CH_CHANGE
        '    KEY_AUTO_ARC_AUTOCOLOR_SUCCESS
        '    KEY_AUTO_ARC_AUTOCOLOR_FAIL
        '    KEY_AUTO_ARC_C_FORCE_AGING
        '    KEY_AUTO_ARC_USBJACK_INSPECT
        '    KEY_AUTO_ARC_JACK_IDENT
        '    KEY_NINE_SEPERATE
        '    KEY_ZOOM_IN
        '    KEY_ZOOM_OUT
        '    KEY_MIC

        '    KEY_AUTO_ARC_CAPTION_KOR
        '    KEY_AUTO_ARC_CAPTION_ENG
        '    KEY_AUTO_ARC_PIP_SOURCE_CHANGE
    End Sub

    Private Sub CreateSamsungRemoteIniFileInfo()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateSamsungRemoteIniFileInfo called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = gRemoteControlPath
        Dim SamsungRemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        If SamsungRemoteType = "Samsungiapp" Then
            WriteStringIniFile(MyUDN, "20", "Power" & ":;:-:" & "KEY_POWER" & ":;:-:1:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "21", "PowerOn" & ":;:-:" & "KEY_POWERON" & ":;:-:1:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "22", "PowerOff" & ":;:-:" & "KEY_POWEROFF" & ":;:-:1:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "23", "Source" & ":;:-:" & "KEY_SOURCE" & ":;:-:2:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "24", "Mute" & ":;:-:" & "KEY_MUTE" & ":;:-:10:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "25", "1" & ":;:-:" & "KEY_1" & ":;:-:5:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "26", "2" & ":;:-:" & "KEY_2" & ":;:-:5:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "27", "3" & ":;:-:" & "KEY_3" & ":;:-:5:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "28", "4" & ":;:-:" & "KEY_4" & ":;:-:6:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "29", "5" & ":;:-:" & "KEY_5" & ":;:-:6:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "30", "6" & ":;:-:" & "KEY_6" & ":;:-:6:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "31", "7" & ":;:-:" & "KEY_7" & ":;:-:7:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "32", "8" & ":;:-:" & "KEY_8" & ":;:-:7:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "33", "9" & ":;:-:" & "KEY_9" & ":;:-:7:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "34", "0" & ":;:-:" & "KEY_0" & ":;:-:8:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "35", "11" & ":;:-:" & "KEY_11" & ":;:-:18:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "36", "12" & ":;:-:" & "KEY_12" & ":;:-:18:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "37", "-" & ":;:-:" & "KEY_PLUS100" & ":;:-:8:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "38", "PRE-CH" & ":;:-:" & "KEY_PRECH" & ":;:-:8:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "39", "TV" & ":;:-:" & "KEY_TV" & ":;:-:2:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "40", "Vol +" & ":;:-:" & "KEY_VOLUP" & ":;:-:10:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "41", "Vol -" & ":;:-:" & "KEY_VOLDOWN" & ":;:-:10:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "42", "Channel +" & ":;:-:" & "KEY_CHUP" & ":;:-:9:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "43", "Channel -" & ":;:-:" & "KEY_CHDOWN" & ":;:-:9:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "44", "Return" & ":;:-:" & "KEY_RETURN" & ":;:-:12:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "45", "Up" & ":;:-:" & "KEY_UP" & ":;:-:11:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "46", "Down" & ":;:-:" & "KEY_DOWN" & ":;:-:11:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "47", "Left" & ":;:-:" & "KEY_LEFT" & ":;:-:11:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "48", "Right" & ":;:-:" & "KEY_RIGHT" & ":;:-:11:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "49", "Enter" & ":;:-:" & "KEY_ENTER" & ":;:-:11:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "50", "Info" & ":;:-:" & "KEY_INFO" & ":;:-:12:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "51", "Exit" & ":;:-:" & "KEY_EXIT" & ":;:-:12:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "52", "Tools" & ":;:-:" & "KEY_TOOLS" & ":;:-:12:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "53", "CH List" & ":;:-:" & "KEY_CHLIST" & ":;:-:9:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "54", "Red" & ":;:-:" & "KEY_RED" & ":;:-:13:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "55", "Green" & ":;:-:" & "KEY_GREEN" & ":;:-:13:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "56", "Blue" & ":;:-:" & "KEY_BLUE" & ":;:-:13:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "57", "Yellow" & ":;:-:" & "KEY_YELLOW" & ":;:-:13:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "58", "REC" & ":;:-:" & "KEY_REC" & ":;:-:14:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "59", "RW" & ":;:-:" & "KEY_RW" & ":;:-:14:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "60", "FF" & ":;:-:" & "KEY_FF" & ":;:-:14:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "61", "InfoLink" & ":;:-:" & "KEY_INFOLINK" & ":;:-:15:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "62", "FavCH" & ":;:-:" & "KEY_FAVCH" & ":;:-:15:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "63", "Content" & ":;:-:" & "KEY_CONTENT" & ":;:-:15:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "64", "Menu" & ":;:-:" & "KEY_MENU" & ":;:-:15:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "65", "WLink" & ":;:-:" & "KEY_WLINK" & ":;:-:15:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "66", "CC" & ":;:-:" & "KEY_CC" & ":;:-:15:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "67", "Guide" & ":;:-:" & "KEY_GUIDE" & ":;:-:16:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "68", "Subtitle" & ":;:-:" & "KEY_SUBTITLE" & ":;:-:16:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "69", "Aspect" & ":;:-:" & "KEY_ASPECT" & ":;:-:16:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "70", "PIP On/Off" & ":;:-:" & "KEY_PIP_ONOFF" & ":;:-:16:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "71", "PIP CH Up" & ":;:-:" & "KEY_PIP_CHUP" & ":;:-:16:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "72", "PIP CH Down" & ":;:-:" & "KEY_PIP_CHDOWN" & ":;:-:16:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "73", "3D" & ":;:-:" & "KEY_3D" & ":;:-:17:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "74", "AD" & ":;:-:" & "KEY_AD" & ":;:-:17:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "75", "PMode" & ":;:-:" & "KEY_PMODE" & ":;:-:17:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "76", "SMode" & ":;:-:" & "KEY_SMODE" & ":;:-:17:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "77", "Step" & ":;:-:" & "KEY_STEP" & ":;:-:14:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "78", "Sleep" & ":;:-:" & "KEY_SLEEP" & ":;:-:1:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "79", "Dolby SRR" & ":;:-:" & "KEY_DOLBY_SRR" & ":;:-:18:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "80", "MTS" & ":;:-:" & "KEY_MTS" & ":;:-:18:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "81", "Disc Menu" & ":;:-:" & "KEY_DISC_MENU" & ":;:-:18:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "82", "EMode" & ":;:-:" & "KEY_EMODE" & ":;:-:18:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "83", "AV1" & ":;:-:" & "KEY_AV1" & ":;:-:3:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "84", "AV2" & ":;:-:" & "KEY_AV2" & ":;:-:3:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "85", "AV3" & ":;:-:" & "KEY_AV3" & ":;:-:3:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "86", "SVIDEO1" & ":;:-:" & "KEY_SVIDEO1" & ":;:-:4:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "87", "SVIDEO2" & ":;:-:" & "SVIDEO2" & ":;:-:4:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "88", "SVIDEO3" & ":;:-:" & "KEY_SVIDEO3" & ":;:-:4:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "89", "COMPONENT1" & ":;:-:" & "KEY_COMPONENT1" & ":;:-:3:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "90", "COMPONENT2" & ":;:-:" & "KEY_COMPONENT2" & ":;:-:3:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "91", "DVI" & ":;:-:" & "KEY_DVI" & ":;:-:4:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "92", "HDMI" & ":;:-:" & "KEY_HDMI" & ":;:-:2:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "93", "HDMI1" & ":;:-:" & "KEY_HDMI1" & ":;:-:2:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "94", "HDMI2" & ":;:-:" & "KEY_HDMI2" & ":;:-:2:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "95", "HDMI3" & ":;:-:" & "KEY_HDMI3" & ":;:-:2:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "96", "HDMI4" & ":;:-:" & "KEY_HDMI4" & ":;:-:2:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "97", "Internet" & ":;:-:" & "KEY_RSS" & ":;:-:4:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "98", "Play" & ":;:-:" & "KEY_PLAY" & ":;:-:14:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "99", "Pause" & ":;:-:" & "KEY_PAUSE" & ":;:-:14:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "100", "Stop" & ":;:-:" & "KEY_STOP" & ":;:-:14:;:-:3", objRemoteFile)

        ElseIf SamsungRemoteType = "SamsungWebSocket" Then
            WriteStringIniFile(MyUDN, "20", "Power" & ":;:-:" & "KEY_POWER" & ":;:-:1:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "21", "PowerOn" & ":;:-:" & "KEY_POWERON" & ":;:-:1:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "22", "PowerOff" & ":;:-:" & "KEY_POWEROFF" & ":;:-:1:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "23", "Source" & ":;:-:" & "KEY_SOURCE" & ":;:-:2:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "24", "Mute" & ":;:-:" & "KEY_MUTE" & ":;:-:10:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "25", "1" & ":;:-:" & "KEY_1" & ":;:-:5:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "26", "2" & ":;:-:" & "KEY_2" & ":;:-:5:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "27", "3" & ":;:-:" & "KEY_3" & ":;:-:5:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "28", "4" & ":;:-:" & "KEY_4" & ":;:-:6:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "29", "5" & ":;:-:" & "KEY_5" & ":;:-:6:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "30", "6" & ":;:-:" & "KEY_6" & ":;:-:6:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "31", "7" & ":;:-:" & "KEY_7" & ":;:-:7:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "32", "8" & ":;:-:" & "KEY_8" & ":;:-:7:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "33", "9" & ":;:-:" & "KEY_9" & ":;:-:7:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "34", "0" & ":;:-:" & "KEY_0" & ":;:-:8:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "35", "11" & ":;:-:" & "KEY_11" & ":;:-:18:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "36", "12" & ":;:-:" & "KEY_12" & ":;:-:18:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "37", "-" & ":;:-:" & "KEY_PLUS100" & ":;:-:8:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "38", "PRE-CH" & ":;:-:" & "KEY_PRECH" & ":;:-:8:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "39", "TV" & ":;:-:" & "KEY_TV" & ":;:-:2:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "40", "Vol +" & ":;:-:" & "KEY_VOLUP" & ":;:-:10:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "41", "Vol -" & ":;:-:" & "KEY_VOLDOWN" & ":;:-:10:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "42", "Channel +" & ":;:-:" & "KEY_CHUP" & ":;:-:9:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "43", "Channel -" & ":;:-:" & "KEY_CHDOWN" & ":;:-:9:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "44", "Return" & ":;:-:" & "KEY_RETURN" & ":;:-:12:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "45", "Up" & ":;:-:" & "KEY_UP" & ":;:-:11:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "46", "Down" & ":;:-:" & "KEY_DOWN" & ":;:-:11:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "47", "Left" & ":;:-:" & "KEY_LEFT" & ":;:-:11:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "48", "Right" & ":;:-:" & "KEY_RIGHT" & ":;:-:11:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "49", "Enter" & ":;:-:" & "KEY_ENTER" & ":;:-:11:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "50", "Info" & ":;:-:" & "KEY_INFO" & ":;:-:12:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "51", "Exit" & ":;:-:" & "KEY_EXIT" & ":;:-:12:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "52", "Tools" & ":;:-:" & "KEY_TOOLS" & ":;:-:12:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "53", "CH List" & ":;:-:" & "KEY_CH_LIST" & ":;:-:9:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "54", "Red" & ":;:-:" & "KEY_RED" & ":;:-:13:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "55", "Green" & ":;:-:" & "KEY_GREEN" & ":;:-:13:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "56", "Blue" & ":;:-:" & "KEY_BLUE" & ":;:-:13:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "57", "Yellow" & ":;:-:" & "KEY_YELLOW" & ":;:-:13:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "58", "REC" & ":;:-:" & "KEY_REC" & ":;:-:14:;:-:6", objRemoteFile)
            'WriteStringIniFile(MyUDN, "59", "RW" & ":;:-:" & "KEY_RW" & ":;:-:14:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "60", "FF" & ":;:-:" & "KEY_FF" & ":;:-:14:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "61", "InfoLink" & ":;:-:" & "KEY_INFOLINK" & ":;:-:15:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "62", "FavCH" & ":;:-:" & "KEY_FAVCH" & ":;:-:15:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "63", "Content" & ":;:-:" & "KEY_CONTENT" & ":;:-:15:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "64", "Menu" & ":;:-:" & "KEY_MENU" & ":;:-:15:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "65", "WLink" & ":;:-:" & "KEY_WLINK" & ":;:-:15:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "66", "CC" & ":;:-:" & "KEY_CC" & ":;:-:15:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "67", "Guide" & ":;:-:" & "KEY_GUIDE" & ":;:-:16:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "68", "Subtitle" & ":;:-:" & "KEY_SUBTITLE" & ":;:-:16:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "69", "Aspect" & ":;:-:" & "KEY_ASPECT" & ":;:-:16:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "70", "PIP On/Off" & ":;:-:" & "KEY_PIP_ONOFF" & ":;:-:16:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "71", "PIP CH Up" & ":;:-:" & "KEY_PIP_CHUP" & ":;:-:16:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "72", "PIP CH Down" & ":;:-:" & "KEY_PIP_CHDOWN" & ":;:-:16:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "73", "3D" & ":;:-:" & "KEY_PANNEL_CHDOWN" & ":;:-:17:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "74", "AD" & ":;:-:" & "KEY_AD" & ":;:-:17:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "75", "PMode" & ":;:-:" & "KEY_PMODE" & ":;:-:17:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "76", "SMode" & ":;:-:" & "KEY_SMODE" & ":;:-:17:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "77", "Step" & ":;:-:" & "KEY_STEP" & ":;:-:14:;:-:7", objRemoteFile)
            'WriteStringIniFile(MyUDN, "78", "Sleep" & ":;:-:" & "KEY_SLEEP" & ":;:-:1:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "79", "Dolby SRR" & ":;:-:" & "KEY_DOLBY_SRR" & ":;:-:18:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "80", "MTS" & ":;:-:" & "KEY_MTS" & ":;:-:18:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "81", "Disc Menu" & ":;:-:" & "KEY_DISC_MENU" & ":;:-:18:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "82", "EMode" & ":;:-:" & "KEY_EMODE" & ":;:-:18:;:-:6", objRemoteFile)
            'WriteStringIniFile(MyUDN, "83", "AV1" & ":;:-:" & "KEY_AV1" & ":;:-:3:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "84", "AV2" & ":;:-:" & "KEY_AV2" & ":;:-:3:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "85", "AV3" & ":;:-:" & "KEY_AV3" & ":;:-:3:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "86", "SVIDEO1" & ":;:-:" & "KEY_SVIDEO1" & ":;:-:4:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "87", "SVIDEO2" & ":;:-:" & "SVIDEO2" & ":;:-:4:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "88", "SVIDEO3" & ":;:-:" & "KEY_SVIDEO3" & ":;:-:4:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "89", "COMPONENT1" & ":;:-:" & "KEY_COMPONENT1" & ":;:-:3:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "90", "COMPONENT2" & ":;:-:" & "KEY_COMPONENT2" & ":;:-:3:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "91", "DVI" & ":;:-:" & "KEY_DVI" & ":;:-:4:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "91", "DTV Source" & ":;:-:" & "KEY_DTV" & ":;:-:4:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "92", "HDMI Source" & ":;:-:" & "KEY_HDMI" & ":;:-:2:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "93", "HDMI1" & ":;:-:" & "KEY_HDMI1" & ":;:-:2:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "94", "HDMI2" & ":;:-:" & "KEY_HDMI2" & ":;:-:2:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "95", "HDMI3" & ":;:-:" & "KEY_HDMI3" & ":;:-:2:;:-:6", objRemoteFile)
            'WriteStringIniFile(MyUDN, "96", "HDMI4" & ":;:-:" & "KEY_HDMI4" & ":;:-:2:;:-:7", objRemoteFile)
            'WriteStringIniFile(MyUDN, "97", "Internet" & ":;:-:" & "KEY_RSS" & ":;:-:4:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "98", "Play" & ":;:-:" & "KEY_PLAY" & ":;:-:14:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "99", "Pause" & ":;:-:" & "KEY_PAUSE" & ":;:-:14:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "100", "Stop" & ":;:-:" & "KEY_STOP" & ":;:-:14:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "107", "SmartHub" & ":;:-:" & "KEY_CONTENTS" & ":;:-:14:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "108", "Text" & ":;:-:" & "SamsungEnterText" & ":;:-:15:;:-:1:;:-:" & "www.cnnfn.com", objRemoteFile)    'dcortizen
        ElseIf SamsungRemoteType = "SamsungWebSocketPIN" Then
            WriteStringIniFile(MyUDN, "20", "Power" & ":;:-:" & "KEY_POWER" & ":;:-:1:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "21", "PowerOn" & ":;:-:" & "KEY_POWERON" & ":;:-:1:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "22", "PowerOff" & ":;:-:" & "KEY_POWEROFF" & ":;:-:1:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "23", "Source" & ":;:-:" & "KEY_SOURCE" & ":;:-:2:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "24", "Mute" & ":;:-:" & "KEY_MUTE" & ":;:-:10:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "25", "1" & ":;:-:" & "KEY_1" & ":;:-:5:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "26", "2" & ":;:-:" & "KEY_2" & ":;:-:5:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "27", "3" & ":;:-:" & "KEY_3" & ":;:-:5:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "28", "4" & ":;:-:" & "KEY_4" & ":;:-:6:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "29", "5" & ":;:-:" & "KEY_5" & ":;:-:6:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "30", "6" & ":;:-:" & "KEY_6" & ":;:-:6:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "31", "7" & ":;:-:" & "KEY_7" & ":;:-:7:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "32", "8" & ":;:-:" & "KEY_8" & ":;:-:7:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "33", "9" & ":;:-:" & "KEY_9" & ":;:-:7:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "34", "0" & ":;:-:" & "KEY_0" & ":;:-:8:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "35", "11" & ":;:-:" & "KEY_11" & ":;:-:18:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "36", "12" & ":;:-:" & "KEY_12" & ":;:-:18:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "37", "-" & ":;:-:" & "KEY_PLUS100" & ":;:-:8:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "38", "PRE-CH" & ":;:-:" & "KEY_PRECH" & ":;:-:8:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "39", "LiveTV" & ":;:-:" & "KEY_TV" & ":;:-:2:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "40", "Vol +" & ":;:-:" & "KEY_VOLUP" & ":;:-:10:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "41", "Vol -" & ":;:-:" & "KEY_VOLDOWN" & ":;:-:10:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "42", "Channel +" & ":;:-:" & "KEY_CHUP" & ":;:-:9:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "43", "Channel -" & ":;:-:" & "KEY_CHDOWN" & ":;:-:9:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "44", "Return" & ":;:-:" & "KEY_RETURN" & ":;:-:12:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "45", "Up" & ":;:-:" & "KEY_UP" & ":;:-:11:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "46", "Down" & ":;:-:" & "KEY_DOWN" & ":;:-:11:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "47", "Left" & ":;:-:" & "KEY_LEFT" & ":;:-:11:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "48", "Right" & ":;:-:" & "KEY_RIGHT" & ":;:-:11:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "49", "Select" & ":;:-:" & "KEY_ENTER" & ":;:-:11:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "50", "Info" & ":;:-:" & "KEY_INFO" & ":;:-:12:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "51", "Exit" & ":;:-:" & "KEY_EXIT" & ":;:-:12:;:-:3", objRemoteFile) ' add
            WriteStringIniFile(MyUDN, "52", "Tools" & ":;:-:" & "KEY_TOOLS" & ":;:-:12:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "53", "CH List" & ":;:-:" & "KEY_CH_LIST" & ":;:-:9:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "54", "Red" & ":;:-:" & "KEY_RED" & ":;:-:13:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "55", "Green" & ":;:-:" & "KEY_GREEN" & ":;:-:13:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "56", "Blue" & ":;:-:" & "KEY_CYAN" & ":;:-:13:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "57", "Yellow" & ":;:-:" & "KEY_YELLOW" & ":;:-:13:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "58", "Record" & ":;:-:" & "KEY_REC" & ":;:-:14:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "59", "Rewind" & ":;:-:" & "KEY_REWIND" & ":;:-:14:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "60", "Fast_Forward" & ":;:-:" & "KEY_FF" & ":;:-:14:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "61", "Info" & ":;:-:" & "KEY_INFO" & ":;:-:15:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "62", "FavCH" & ":;:-:" & "KEY_FAVCH" & ":;:-:15:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "63", "Home" & ":;:-:" & "KEY_CONTENT" & ":;:-:15:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "64", "Menu" & ":;:-:" & "KEY_MENU" & ":;:-:15:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "65", "WLink" & ":;:-:" & "KEY_WLINK" & ":;:-:15:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "66", "CC" & ":;:-:" & "KEY_CC" & ":;:-:15:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "67", "Guide" & ":;:-:" & "KEY_GUIDE" & ":;:-:16:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "68", "Subtitle" & ":;:-:" & "KEY_SUBTITLE" & ":;:-:16:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "69", "Aspect" & ":;:-:" & "KEY_ASPECT" & ":;:-:16:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "70", "PIP On/Off" & ":;:-:" & "KEY_PIP_ONOFF" & ":;:-:16:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "71", "PIP CH Up" & ":;:-:" & "KEY_PIP_CHUP" & ":;:-:16:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "72", "PIP CH Down" & ":;:-:" & "KEY_PIP_CHDOWN" & ":;:-:16:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "73", "3D" & ":;:-:" & "KEY_PANNEL_CHDOWN" & ":;:-:17:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "74", "AD" & ":;:-:" & "KEY_AD" & ":;:-:17:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "75", "PMode" & ":;:-:" & "KEY_PMODE" & ":;:-:17:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "76", "SMode" & ":;:-:" & "KEY_SMODE" & ":;:-:17:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "77", "Step" & ":;:-:" & "KEY_STEP" & ":;:-:14:;:-:7", objRemoteFile)
            'WriteStringIniFile(MyUDN, "78", "Sleep" & ":;:-:" & "KEY_SLEEP" & ":;:-:1:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "79", "Dolby SRR" & ":;:-:" & "KEY_DOLBY_SRR" & ":;:-:18:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "80", "MTS" & ":;:-:" & "KEY_MTS" & ":;:-:18:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "81", "Disc Menu" & ":;:-:" & "KEY_DISC_MENU" & ":;:-:18:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "82", "EMode" & ":;:-:" & "KEY_EMODE" & ":;:-:18:;:-:6", objRemoteFile)
            WriteStringIniFile(MyUDN, "83", "AV1" & ":;:-:" & "KEY_AV1" & ":;:-:3:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "84", "AV2" & ":;:-:" & "KEY_AV2" & ":;:-:3:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "85", "AV3" & ":;:-:" & "KEY_AV3" & ":;:-:3:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "86", "SVIDEO1" & ":;:-:" & "KEY_SVIDEO1" & ":;:-:4:;:-:1", objRemoteFile)
            'WriteStringIniFile(MyUDN, "87", "SVIDEO2" & ":;:-:" & "SVIDEO2" & ":;:-:4:;:-:2", objRemoteFile)
            'WriteStringIniFile(MyUDN, "88", "SVIDEO3" & ":;:-:" & "KEY_SVIDEO3" & ":;:-:4:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "89", "COMPONENT1" & ":;:-:" & "KEY_COMPONENT1" & ":;:-:3:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "90", "COMPONENT2" & ":;:-:" & "KEY_COMPONENT2" & ":;:-:3:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "91", "DVI" & ":;:-:" & "KEY_DVI" & ":;:-:4:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "91", "DTV Source" & ":;:-:" & "KEY_DTV" & ":;:-:4:;:-:4", objRemoteFile)
            WriteStringIniFile(MyUDN, "92", "HDMI Source" & ":;:-:" & "KEY_HDMI" & ":;:-:2:;:-:3", objRemoteFile)
            'WriteStringIniFile(MyUDN, "93", "HDMI1" & ":;:-:" & "KEY_HDMI1" & ":;:-:2:;:-:4", objRemoteFile)
            'WriteStringIniFile(MyUDN, "94", "HDMI2" & ":;:-:" & "KEY_HDMI2" & ":;:-:2:;:-:5", objRemoteFile)
            'WriteStringIniFile(MyUDN, "95", "HDMI3" & ":;:-:" & "KEY_HDMI3" & ":;:-:2:;:-:6", objRemoteFile)
            'WriteStringIniFile(MyUDN, "96", "HDMI4" & ":;:-:" & "KEY_HDMI4" & ":;:-:2:;:-:7", objRemoteFile)
            'WriteStringIniFile(MyUDN, "97", "Internet" & ":;:-:" & "KEY_RSS" & ":;:-:4:;:-:5", objRemoteFile)
            WriteStringIniFile(MyUDN, "98", "Play" & ":;:-:" & "KEY_PLAY" & ":;:-:14:;:-:1", objRemoteFile)
            WriteStringIniFile(MyUDN, "99", "Pause" & ":;:-:" & "KEY_PAUSE" & ":;:-:14:;:-:2", objRemoteFile)
            WriteStringIniFile(MyUDN, "100", "Stop" & ":;:-:" & "KEY_STOP" & ":;:-:14:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "101", "Skip_Back" & ":;:-:" & "KEY_PREV" & ":;:-:14:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "102", "Skip_Forward" & ":;:-:" & "KEY_NEXT" & ":;:-:14:;:-:3", objRemoteFile)
            WriteStringIniFile(MyUDN, "103", "DVR" & ":;:-:" & "KEY_DVR" & ":;:-:2:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "104", "Tools" & ":;:-:" & "KEY_TOOLS" & ":;:-:2:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "105", "Anynet" & ":;:-:" & "KEY_ANYNET" & ":;:-:2:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "106", "Antenna" & ":;:-:" & "KEY_ANTENA" & ":;:-:2:;:-:7", objRemoteFile)
            WriteStringIniFile(MyUDN, "107", "SmartHub" & ":;:-:" & "KEY_CONTENTS" & ":;:-:14:;:-:4", objRemoteFile)
            ' 
        End If



    End Sub


    Private Sub CreateHSSamsungRemoteButtons(ReCreate As Boolean)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSSamsungRemoteButtons called for UPnPDevice = " & MyUPnPDeviceName & " and Recreate = " & ReCreate.ToString, LogType.LOG_TYPE_INFO)
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

        If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
            Pair.PairType = VSVGPairType.SingleValue
            Pair.Value = psCreateRemoteAppButtons
            Pair.Status = "Retrieve App Info"
            Pair.Render = Enums.CAPIControlType.Button
            Pair.Render_Location.Row = 1
            Pair.Render_Location.Column = Column + 3
            hs.DeviceVSP_AddPair(HSRefRemote, Pair)
        End If

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
        'CreateRemoteButtons(HSRefRemote) ' changed in V24
    End Sub

    Private Sub CreateHSSamsungRemoteServices()

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateHSSamsungRemoteServices called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        ' First Create the Status Device
        HSRefServiceRemote = GetIntegerIniFile(MyUDN, "di" & HSDevices.Status.ToString & "HSCode", -1)
        If HSRefServiceRemote = -1 Then
            HSRefServiceRemote = CreateHSServiceDevice(HSRefRemote, HSDevices.Status.ToString)
        End If
        If HSRefServiceRemote = -1 Then Exit Sub

    End Sub

    Private Sub CreateSamsungRemoteServiceButtons(HSRef As Integer)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateSamsungRemoteServiceButtons called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Dim objRemoteFile As String = gRemoteControlPath
        Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
        Dim AppOffset As Integer = (GetIntegerIniFile("Samsung App Info" & MyUDN, "AddAppButtons", 20, objRemoteFile) / MaxButtonCollumns) + 1
        Dim InputOffset As Integer = (GetIntegerIniFile("LG Offset Info" & MyUDN, "AddInputButtons", 40, objRemoteFile) / MaxButtonCollumns) + 1 + AppOffset
        ' Dim ChannelOffset As Integer = (GetIntegerIniFile("LG Offset Info" & myudn, "AddChannelButtons", 60, objRemoteFile) / MaxButtonCollumns) + 1 + InputOffset

        Try
            RemoteButtons = GetIniSection(MyUDN, objRemoteFile) '  As Dictionary(Of String, String)
            If RemoteButtons Is Nothing Then
                Log("Error in CreateSamsungRemoteServiceButtons for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
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
                            If RemoteButtonInfos(1).IndexOf("SamsungApp") = 0 Then
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
                            End If
                        End If
                    End If
                Next
                RemoteButtons = Nothing
            End If
        Catch ex As Exception
            Log("Error in CreateSamsungRemoteServiceButtons parsing the .ini file for remote button info for device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub SamsungAddAppButtons(Payload As Object)

        ' this is the list of installed apps in following structure
        ' if non eden
        ' {"data":{"data":[{"appId":"org.tizen.browser","app_type":4,"icon":"/opt/share/webappservice/apps_icon/FirstScreen/webbrowser/250x250.png","is_lock":0,"name":"Internet"}]},"event":"ed.installedApp.get","from":"host"}
        ' if eden
        ' {"data":{"data":[{"accelerators":[],"action_type":null,"appId":"com.samsung.tv.store","appType":"volt_app","icon":"/usr/apps/com.samsung.tv.csfs.res.tizen30/shared/res/Resource/apps/apps/sysAppsNromal.png","id":"APPS","isLock":false,"launcherType":"system","mbrIndex":null,"mbrSource":null,"name":"APPS","position":0,"sourceTypeNum":null},{"accelerators":[],"action_type":null,"appId":"org.tizen.browser","appType":"web_app","icon":"/opt/share/webappservice/apps_icon/FirstScreen/webbrowser/245x138.png","id":"org.tizen.browser","isLock":false,"launcherType":"launcher","mbrIndex":null,"mbrSource":null,"name":"Internet","position":1,"sourceTypeNum":null}]},"event":"ed.edenApp.get","from":"host"} 

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAddAppButtons called for Device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim SeqNbr As Integer = 0
            Dim ButtonIndex As Integer = 200
            Dim RowIndex As Integer = 1
            Dim ColumnIndex As Integer = 1
            Dim objRemoteFile As String = gRemoteControlPath
            Dim DataL1 As Object = FindPairInJSONString(Payload, "data")
            Dim DataL2 As Object = FindPairInJSONString(DataL1, "data")
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(DataL2)
            WriteStringIniFile("Samsung App Info" & MyUDN, "AddAppButtons", JSONdataLevel1.length, objRemoteFile)

            For Each Entries As Object In JSONdataLevel1
                Dim appId As String = ""
                Dim app_type As String = ""
                Dim name As String = ""
                Dim Icon As String = ""
                Dim AppImage As Image = Nothing
                Dim id As String = ""
                Dim launcherType As String = ""
                Dim accelerators As Object = Nothing
                Dim position As String = ""
                Dim sourceTypeNum As String = ""
                Dim mbrIndex As String = ""
                Dim mbrSource As String = ""
                Dim ActionType As String = ""
                Dim isLock As String = ""
                Dim Value As String = ""
                For Each Entry As Object In Entries
                    If Entry.key IsNot Nothing Then
                        If Entry.value IsNot Nothing Then
                            Value = Entry.value.ToString
                        End If
                        If Entry.Key.ToString = "appId" Then
                            appId = Value
                        ElseIf (Entry.Key.ToString) = "app_type" Or (Entry.key.ToString = "appType") Then
                            app_type = Value
                        ElseIf Entry.Key.ToString = "name" Then
                            name = Value
                        ElseIf Entry.Key.ToString = "action_type" Then
                            ActionType = Value
                        ElseIf Entry.Key.ToString = "id" Then
                            id = Value
                        ElseIf Entry.Key.ToString = "isLock" Then
                            isLock = Value
                        ElseIf Entry.Key.ToString = "launcherType" Then
                            launcherType = Value
                        ElseIf Entry.Key.ToString = "mbrIndex" Then
                            mbrIndex = Value
                        ElseIf Entry.Key.ToString = "mbrSource" Then
                            mbrSource = Value
                        ElseIf Entry.Key.ToString = "accelerators" Then
                            accelerators = Entry.value
                        ElseIf Entry.Key.ToString = "sourceTypeNum" Then
                            sourceTypeNum = Value
                        ElseIf Entry.Key.ToString = "position" Then
                            position = Value
                        ElseIf Entry.Key.ToString = "icon" Then
                            Icon = Value
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " retrieving icon info =" & Icon, LogType.LOG_TYPE_INFO)
                            ' 6/4/2020 going to remove this for now, it really isn't useful
                            'If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""iconPath"":""" & Icon & """,""event"": ""ed.apps.icon"", ""to"":""host""}}"), True) Then
                            'If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " retrieving icon info =" & Icon, LogType.LOG_TYPE_INFO)
                            'End If
                        End If
                    End If
                Next
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " found appID = " & appId & " with type =" & app_type & ", name=" & name & " and ICON=" & Icon, LogType.LOG_TYPE_INFO)
                WriteStringIniFile(MyUDN, ButtonIndex.ToString, name & ":;:-:" & "SamsungApp" & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString & ":;:-:" & appId & ":;:-:" & "SamsungAppImage_" & MyUDN & "_" & SeqNbr.ToString & ":;:-:" & app_type.ToString, objRemoteFile)
                If Icon <> "" And appId <> "" And 1 = 2 Then ' this does not work anyway !!!
                    Try
                        AppImage = GetPicture(Icon.ToString)
                        If Not AppImage Is Nothing Then
                            Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                            Dim SuccesfullSave As Boolean = False
                            SuccesfullSave = hs.WriteHTMLImage(AppImage, FileArtWorkPath & "SamsungAppImage_" & MyUDN & "_" & SeqNbr.ToString & ".png", True)
                            If Not SuccesfullSave Then
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " had error storing Image at " & FileArtWorkPath & "AppImage_" & MyUDN & "_" & SeqNbr.ToString.ToString & ".png", LogType.LOG_TYPE_ERROR)
                            Else
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " stored App Image at " & FileArtWorkPath & "AppImage_" & MyUDN & "_" & SeqNbr.ToString.ToString & ".png", LogType.LOG_TYPE_INFO)
                            End If
                            AppImage.Dispose()
                            SeqNbr += 1
                        End If
                    Catch ex As Exception
                        Log("Error in SamsungAddAppButtons retrieving the LG App Image with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAddAppButtons for Device = " & MyUPnPDeviceName & " while processing response with error = " & ex.Message & " with Payload = " & Payload.ToString, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TreatSetIOExSamsung(ButtonValue As Integer)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSetIOExSamsung called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Select Case ButtonValue
            Case psRemoteOff  ' remote off
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, False)
                Try
                    SamsungCloseTCPConnection(False)
                Catch ex As Exception
                    Log("Error in TreatSetIOExSamsung for UPnPDevice = " & MyUPnPDeviceName & "  setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                SetAdministrativeStateRemote(False)
            Case psRemoteOn  ' remote on
                WriteBooleanIniFile("Remote Service by UDN", MyUDN, True)
                If SamsungActivateRemote() Then
                    SetAdministrativeStateRemote(True)
                End If
            Case psCreateRemoteButtons
                CreateSamsungRemoteIniFileInfo()
                CreateHSSamsungRemoteButtons(True)
                CreateRemoteButtons(HSRefRemote)
                CreateSamsungRemoteServiceButtons(HSRefServiceRemote)
                If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then CreateLGRemoteServiceButtons(HSRefServiceRemote)
            Case psRegister  ' Register
                SamsungResetIdentityParms()
                If Not SamsungGetIdentifyParms() Then
                    SamsungOpenPinPage()
                End If
            Case psWOL
                SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), PlugInIPAddress)
                SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, ""), PlugInIPAddress)
                Dim deviceIpAddress As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diIPAddress.ToString, "")
                If deviceIpAddress <> "" Then
                    SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, ""), deviceIpAddress)
                    SendMagicPacket(GetStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, ""), deviceIpAddress)
                End If
            Case psCreateRemoteAppButtons
                'CreateSamsungAppButtons(HSRefRemote)
            Case Else
                If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) And UCase(DeviceStatus) = "ONLINE" Then
                    Dim objRemoteFile As String = gRemoteControlPath
                    Dim ButtonInfoString As String = GetStringIniFile(MyUDN, ButtonValue.ToString, "", objRemoteFile)
                    SamsungSendKeyCode(ButtonInfoString)
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in TreatSetIOExSamsung for UPnPDevice = " & MyUPnPDeviceName & ". The remote is off line", LogType.LOG_TYPE_WARNING)
                End If
        End Select

    End Sub

    Private Function SamsungActivateRemote() As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungActivateRemote called for UPnPDevice = " & MyUPnPDeviceName & " has ServiceStateActive = " & MyRemoteServiceActive.ToString & " and admin state = " & MyAdminStateActive.ToString, LogType.LOG_TYPE_INFO)
        SamsungActivateRemote = False
        If Not MyRemoteServiceActive And MyAdminStateActive Then
            If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
                ' Newer models 2016
                SamsungOpenUnEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, ""), GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, ""))
            ElseIf GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocketPIN" Then
                ' check if we have a pin. Models Y2014 and Y2015
                If Not SamsungGetIdentifyParms() Then
                    'SamsungSessionID = SamsungAuthenticateUsePIN()
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungActivateRemote for UPnPDevice = " & MyUPnPDeviceName & " trying to turn remote on but no authentication credentials available. Go to conf screen and authenticate with PIN", LogType.LOG_TYPE_ERROR)
                    Return False
                End If
                SamsungSessionID = SamsungAuthenticateUsePIN(GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, ""), True)

                If SamsungSessionID = "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungActivateRemote for UPnPDevice = " & MyUPnPDeviceName & " no sessionID", LogType.LOG_TYPE_ERROR)
                    Return False
                End If
                SamsungOpenEncryptedWebSocket(SamsungSessionID)
            Else
                Try
                    ' Legacy way older models <2014
                    SamsungEstablishTCPConnection()
                Catch ex As Exception
                    Log("Error in SamsungActivateRemote for UPnPDevice = " & MyUPnPDeviceName & "  setting Remote Service flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Return False
                End Try
            End If
        End If
        SetHSRemoteState()
    End Function

    Private Function ProcessReturnString(InputString As String) As String
        ProcessReturnString = InputString
        If InputString = "" Then Exit Function
        '<Byte - unknown byte>
        '    0x00 = ???
        '<Byte - Length of string1>
        '0:      x00()
        '<x-Bytes - string1>
        '   "iapp.samsung"
        '<Byte - Length of resultbytes>
        '0:      x00()
        '      <resultbytes>
        '           0x00 0x00 0x00 0x00
        Dim Bytes As Byte()
        Dim OutputString As String = ""
        Dim TempInteger As Integer
        Dim HexAsciiString As String = ""
        Bytes = System.Text.Encoding.UTF8.GetBytes(InputString)
        For Each ByteElement In Bytes
            ' convert from byte to hex in ascii
            If ByteElement >= &H20 And ByteElement <= &H7A Then
                HexAsciiString = HexAsciiString & Chr(ByteElement)
            Else
                TempInteger = ByteElement / 16
                HexAsciiString = HexAsciiString & "0x" & ConvertBase16(TempInteger)
                TempInteger = ByteElement Mod 16
                HexAsciiString = HexAsciiString & ConvertBase16(TempInteger) & " "
            End If

        Next
        ProcessReturnString = HexAsciiString

    End Function

    Private Function ExtractSamsungInfoFromDeviceXML(PageHTML As String) As Boolean
        ExtractSamsungInfoFromDeviceXML = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' added on 9/23/2020 to cover the cases where you have no more info on remotes. Version .68
        If GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then ExtractSamsungInfoFromDeviceXML = True
        If PageHTML = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ExtractSamsungInfoFromDeviceXML called for device = " & MyUPnPDeviceName & ". Empty HTML", LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        Dim xmlDoc As New XmlDocument With {.XmlResolver = Nothing}

        PageHTML = RemoveControlCharacters(PageHTML)

        Try
            xmlDoc.LoadXml(PageHTML)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " while retieving document with URL = " & MyDocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Dim SamsungIPRemote As String = ""

        Try
            SamsungIPRemote = xmlDoc.GetElementsByTagName("sec:X_IPControl").Item(0).InnerXml
            '<sec:X_IPControl>wlanMacAddress:cc:6e:a4:d9:ab:c8, eth0MacAddress:c0:48:e6:c3:3b:8d,supportDMR:1,DMR_UDN:uuid : 395b84db-ae68-452d-b241-f947ade672c7,listenFrequency:2412,wlanFrequency:0,bssid:00:00:00:00:00:00,vdProductType:TV</sec:X_IPControl>
            If SamsungIPRemote <> "" Then
                Dim IPParts As String() = SamsungIPRemote.Split(",")
                Try
                    For Each IPPart As String In IPParts
                        If IPPart <> "" Then
                            Dim TagParts As String() = IPPart.Split(":")
                            If TagParts.Length > 1 Then
                                If TagParts(0) = "wlanMacAddress" Then
                                    Dim Macaddress As String = ""
                                    IPPart = IPPart.Remove(0, 14)
                                    'IPPart.CopyTo(15, Macaddress, 0, IPParts.Length - 15)
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, Macaddress.Replace("-", ":"))
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " stored wifiMac Address = " & Macaddress, LogType.LOG_TYPE_INFO)
                                ElseIf TagParts(0) = "eth0MacAddress" Then
                                    Dim Macaddress As String = ""
                                    IPPart = IPPart.Remove(0, 14)
                                    'IPPart.CopyTo(15, Macaddress, 0, IPParts.Length - 15)
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diMACAddress.ToString, Macaddress.Replace("-", ":"))
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " stored Mac Address = " & Macaddress, LogType.LOG_TYPE_INFO)
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " while parsing sec:X_IPControl with info = " & SamsungIPRemote & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found X_IPControl = " & SamsungIPRemote.ToString, LogType.LOG_TYPE_INFO)
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "SamsungWebSocket")
            'WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketName.ToString, CapabilityName)
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "8001")
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "api/v2/channels/samsung.remote.control?name=" & ToBase64("MediaController"))
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "u5Y00EnwvXkhM4CqnAGJVQ==")
            Return True
        Catch ex As Exception
            'Log("Error in ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " while parsing sec:X_IPControl (1) with info = " & SamsungIPRemote & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            Dim SamsungCapabilities As String = xmlDoc.GetElementsByTagName("sec:Capabilities").Item(0).InnerXml
            ' if not found it will go to the exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found SamsungTag sec:Capabilities with info = " & SamsungCapabilities, LogType.LOG_TYPE_INFO)
            If SamsungCapabilities <> "" Then
                Dim SamsungCapabilitiesxmlDoc As New XmlDocument
                SamsungCapabilitiesxmlDoc.LoadXml(SamsungCapabilities)

                Try
                    Dim CapabilityName As String = SamsungCapabilitiesxmlDoc.GetElementsByTagName("sec:Capability").Item(0).Attributes("name").InnerText
                    Dim CapabilityPort As String = SamsungCapabilitiesxmlDoc.GetElementsByTagName("sec:Capability").Item(0).Attributes("port").InnerText
                    Dim CapabilityLocation As String = SamsungCapabilitiesxmlDoc.GetElementsByTagName("sec:Capability").Item(0).Attributes("location").InnerText
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found Capability with Name = " & CapabilityName & ", Port = " & CapabilityPort & ", Location = " & CapabilityLocation, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "SamsungWebSocket")
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketName.ToString, CapabilityName)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "8001")
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "api/v2/channels/samsung.remote.control?name=" & ToBase64("MediaController"))
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "u5Y00EnwvXkhM4CqnAGJVQ==")
                    ExtractSamsungInfoFromDeviceXML = True
                Catch ex As Exception
                End Try
                Try ' sec:deviceID
                    Dim DeviceID As String = xmlDoc.GetElementsByTagName("sec:deviceID").Item(0).InnerXml
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found deviceID = " & DeviceID, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungDeviceID.ToString, DeviceID)
                Catch ex As Exception
                End Try
                Try
                    ' sec:ProductCap
                    Dim ProductCap As String = xmlDoc.GetElementsByTagName("sec:ProductCap").Item(0).InnerXml
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found ProductCap = " & ProductCap, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungProductCap.ToString, ProductCap)
                Catch ex As Exception
                End Try
            End If

            Try
                Dim ProductCap As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungProductCap.ToString, "") 'xmlDoc.GetElementsByTagName("sec:ProductCap").Item(0).InnerText
                If ProductCap <> "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found SamsungTag sec:ProductCap with info = " & ProductCap, LogType.LOG_TYPE_INFO)
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketProductCap.ToString, ProductCap)
                    Dim test As String = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "")
                    If GetStringIniFile(DeviceUDN, DeviceInfoIndex.diRemoteType.ToString, "") = "SamsungWebSocket" Then
                        If ProductCap <> "" Then
                            Dim ProductCapParts As String() = Split(ProductCap, ",")
                            For Each Capability As String In ProductCapParts
                                If Capability.ToUpper = "Y2014" Or Capability.ToUpper = "Y2015" Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found a Samsung TV with PIN remote control. Year = " & Capability, LogType.LOG_TYPE_INFO)
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "SamsungWebSocketPIN")
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "8000")
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungPairingPort.ToString, "8080")
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "socket.io/1/")
                                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungAppId.ToString, "654321")
                                ElseIf Capability.ToUpper = "Y2017" Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found a Qseries Samsung TV. Year = " & Capability, LogType.LOG_TYPE_INFO)
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ExtractSamsungInfoFromDeviceXML for device = " & MyUPnPDeviceName & " found a Samsung TV Capability = " & Capability, LogType.LOG_TYPE_INFO)
                                End If
                            Next
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try
        Catch ex As Exception
        End Try
        xmlDoc = Nothing
    End Function


    Private Sub SamsungEstablishTCPConnection()

        ' My C8000 TV uses TCP Port 55000
        'my $myip = "192.168.100.2"; #Doesn't seem to be really used
        'my $mymac = "01-23-45-67-89-ab"; #Used for the access control/validation, but not after that AFAIK
        'my $appstring = "iphone..iapp.samsung"; #What the iPhone app reports
        'my $tvappstring = "iphone..iapp.samsung"; #Might need changing to match your TV type
        'my $remotename = "Perl Samsung Remote"; #What gets reported when it asks for permission/also shows in General->Wireless Remote Control menu

        'my $messagepart1 = chr(0x64) . chr(0x00) . chr(length(encode_base64($myip, ""))) . chr(0x00) . encode_base64($myip, "") . chr(length(encode_base64($mymac, ""))) . chr(0x00) . encode_base64($mymac, "") . chr

        '(length(encode_base64($remotename, ""))) . chr(0x00) . encode_base64($remotename, "");
        'my $part1 = chr(0x00) . chr(length($appstring)) . chr(0x00) . $appstring . chr(length($messagepart1)) . chr(0x00) . $messagepart1;

        'print $sock $part1;
        'print $part1;
        'print "\n";

        'my $messagepart2 = chr(0xc8) . chr(0x00);
        'my $part2 = chr(0x00) . chr(length($appstring)) . chr(0x00) . $appstring . chr(length($messagepart2)) . chr(0x00) . $messagepart2;
        'print $sock $part2;
        'print $part2;
        'print "\n";

        ' little code to convert some captured keys
        'Dim Charar As Byte()
        'Charar = FromBase64("S0VZX1BMVVMxMDA=") - this is KEY_PLUS100
        'Dim ascii As Encoding = Encoding.ASCII
        'Dim ReturnStr As String = ascii.GetChars(Charar)
        'log( "Our converted string = " & ReturnStr)

        ' Need to find out my MAC Address

        'Dim MyIP As String = "192.168.1.117"
        'Dim MyMac As String = "00-21-9B-23-AA-F7"
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("EstablishTCPConnection called for UPnPDevice = " & MyUPnPDeviceName & " and MyRemoteServiceActive = " & MyRemoteServiceActive, LogType.LOG_TYPE_INFO)
        Dim MyPort As String = "55000"
        Dim MyAppString As String = "iphone..iapp.samsung"
        Dim MyTVAppString As String = "iphone..iapp.samsung"
        Dim MyRemoteName As String = "HomeSeer Samsung Remote"
        Dim MessagePart1 As String = ""
        Dim MessagePart2 As String = ""
        Dim MessagePart3 As String = ""
        MessagePart1 = Chr(&H64) & Chr(0) & Chr(ToBase64(MyIPAddress).Length) & Chr(0) & ToBase64(MyIPAddress) & Chr(ToBase64(MyMacAddress).Length) & Chr(0) & ToBase64(MyMacAddress) &
            Chr(ToBase64(MyRemoteName).Length) & Chr(0) & ToBase64(MyRemoteName)
        'Dim Part1 As String = ""
        Dim Part1 As Byte() = System.Text.Encoding.ASCII.GetBytes(Chr(0) & Chr(MyAppString.Length) & Chr(0) & MyAppString & Chr(MessagePart1.Length) & Chr(0) & MessagePart1)
        'Part1 = Chr(0) & Chr(MyAppString.Length) & Chr(0) & MyAppString & Chr(MessagePart1.Length) & Chr(0) & MessagePart1


        MessagePart2 = Chr(&HC8) & Chr(0)
        'Dim Part2 As String = ""
        'Part2 = Chr(0) & Chr(MyAppString.Length) & Chr(0) & MyAppString & Chr(MessagePart2.Length) & Chr(0) & MessagePart2
        Dim Part2 As Byte() = System.Text.Encoding.ASCII.GetBytes(Chr(0) & Chr(MyAppString.Length) & Chr(0) & MyAppString & Chr(MessagePart2.Length) & Chr(0) & MessagePart2)

        'Dim MySocket As New GetSocket
        If MySamsungAsyncSocket Is Nothing Then
            Try
                MySamsungAsyncSocket = New AsynchronousClient
            Catch ex As Exception
                Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to open Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyRemoteServiceActive = False
                Exit Sub
            End Try
        End If

        AddHandler MySamsungAsyncSocket.DataReceived, AddressOf HandleSamsungTCPDataReceived

        Try
            MySamsungClient = MySamsungAsyncSocket.ConnectSocket(MyIPAddress, MyPort)
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to open Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyRemoteServiceActive = False
            Exit Sub
        End Try

        If MySamsungClient Is Nothing Then
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " . Unable to open Socket", LogType.LOG_TYPE_ERROR)
            MyRemoteServiceActive = False
            Exit Sub
        End If

        Dim WaitForConnection As Integer = 0
        Do While WaitForConnection < 10
            If MySamsungAsyncSocket.MySocketIsClosed Then
                wait(1)
                WaitForConnection = WaitForConnection + 1
            Else
                Exit Do
            End If
        Loop

        If WaitForConnection >= 10 Then
            ' unsuccesfull connection
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " . Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
            MySamsungAsyncSocket.CloseSocket()
            Exit Sub
        End If

        'Dim ReturnString As String = ""
        If PIDebuglevel > DebugLevel.dlEvents Then Log("EstablishTCPConnection for device - " & MyUPnPDeviceName & " will send Part 1 = " & Encoding.UTF8.GetString(Part1, 0, Part1.Length), LogType.LOG_TYPE_INFO)

        'If MyRemoteControlHSCode <> "" Then hs.setdeviceValue(MyRemoteControlHSCode, 1)
        MyRemoteServiceActive = True
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then MySamsungAsyncSocket.Receive()
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MySamsungAsyncSocket.response = False
        Try
            If Not MySamsungAsyncSocket.Send(Part1) Then
                Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                SamsungCloseTCPConnection(True)
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            SamsungCloseTCPConnection(True)
            Exit Sub
        End Try
        Try
            ' Send test data to the remote device.
            MySamsungAsyncSocket.sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            SamsungCloseTCPConnection(True)
            Exit Sub
        End Try

        Dim WaitResponse As Integer = 0
        Do While WaitResponse < 10
            If Not MySamsungAsyncSocket.response Then
                wait(1)
                WaitResponse = WaitResponse + 1
            Else
                MySamsungAsyncSocket.response = False
                Exit Do
            End If
        Loop

        ' will be 0x02 0x2F 0x00 unknown.livingroom.iapp.samsung0x02 0x00 e0x00  if refused
        ' not sure what this meant 0x02 0x2F 0x00 unknown.livingroom.iapp.samsung0x04 0x00 d0x00 0x01 0x00 0x02 0x2F 0x00 unknown.livingroom.iapp.samsung0x04 0x00 ,0x01 0x02 0x00 
        If PIDebuglevel > DebugLevel.dlEvents Then Log("EstablishTCPConnection  for device - " & MyUPnPDeviceName & " will send Part 2 = " & Encoding.UTF8.GetString(Part2, 0, Part2.Length), LogType.LOG_TYPE_INFO)
        Try
            If Not MySamsungAsyncSocket.Send(Part2) Then
                Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                SamsungCloseTCPConnection(True)
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            SamsungCloseTCPConnection(True)
            Exit Sub
        End Try
        Try
            MySamsungAsyncSocket.sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in EstablishTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            SamsungCloseTCPConnection(True)
            Exit Sub
        End Try

    End Sub

    Private Sub SamsungCloseTCPConnection(Force As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungCloseTCPConnection called for UPnPDevice = " & MyUPnPDeviceName & ", Force = " & Force.ToString & " and RemoteServiceActive = " & MyRemoteServiceActive.ToString, LogType.LOG_TYPE_INFO)
        'If Not (MyRemoteServiceActive Or Force) Then Exit Sub
        Dim SamsungRemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        If SamsungRemoteType = "Samsungiapp" Then
            If MySamsungAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler MySamsungAsyncSocket.DataReceived, AddressOf HandleSamsungTCPDataReceived
                Catch ex As Exception
                End Try
                Try
                    MySamsungAsyncSocket.CloseSocket()
                Catch ex As Exception
                    Log("Error in CloseTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                MySamsungAsyncSocket = Nothing
            End If
            MySamsungClient = Nothing

        Else
            If MySamsungWebSocket IsNot Nothing Then
                Try
                    RemoveHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
                Catch ex As Exception
                End Try
                Try
                    RemoveHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
                Catch ex As Exception
                End Try
                Try
                    MySamsungWebSocket.CloseSocket()
                Catch ex As Exception
                    Log("Error in CloseTCPConnection for UPnPDevice = " & MyUPnPDeviceName & " closing the websocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                MySamsungWebSocket = Nothing
            End If
        End If
        Try
            If HSRefRemote <> -1 Then hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
            MyRemoteServiceActive = False
        Catch ex As Exception
            Log("Error in CloseTCPConnection 2 for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        SetHSRemoteState()
    End Sub


    Private Sub SamsungResetIdentityParms()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungResetIdentityParms called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        SamsungAESKey = Nothing
        SamsungSessionID = ""
        WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungAesKey.ToString, "")
        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungSessionID.ToString, "")
        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, "")
        ' need to  update HS status!
        SetHSRemoteState()
    End Sub

    Private Function SamsungGetIdentifyParms() As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungGetIdentifyParms called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        SamsungSessionID = ""
        SamsungAESKey = Nothing
        SamsungAESKey = Convert.FromBase64String(GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungAesKey.ToString, ""))
        SamsungSessionID = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungSessionID.ToString, "")
        SamsungPIN = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, "")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungGetIdentifyParms called for device - " & MyUPnPDeviceName & " retrieved PIN=" & SamsungPIN & ", SessionID=" & SamsungSessionID & ",AESKey=" & ByteArrayToHexString(SamsungAESKey), LogType.LOG_TYPE_INFO)
        If SamsungAESKey IsNot Nothing And SamsungSessionID <> "" And SamsungPIN <> "" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SamsungOpenPinPage() As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenPinPage called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim pairingport As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungPairingPort.ToString, "")
        Dim RequestURL As String = "http://" & MyIPAddress & ":" & pairingport & "/ws/apps/CloudPINPage"
        Dim data = System.Text.ASCIIEncoding.ASCII.GetBytes("pin4")

        Dim ReturnHeader As String = ""
        Dim ReturnBody As String = ""

        If Not SendWebRequest(RequestURL, "POST", False, "pin4", ReturnHeader, ReturnBody) Then
            Return False
        End If
        Return True
    End Function

    Public Function SamsungAuthenticateUsePIN(Pin As String, WeAlreadyHaveAPIN As Boolean) As String

        ' this is some pre work to get all the ports correct. The URL comes from the capabilities

        'GET /ms/1.0/ HTTP/1.1
        'Host: 192.168.1.145:8001
        'Connection: Keep-Alive

        'HTTP/ 1.1 200 OK
        'X-Powered - By: Express
        'Access-Control - Allow - Origin:  *
        'Access-Control - Allow - Credentials:  true
        'Access-Control - Allow - Methods:  GET,PUT,POST,DELETE
        'Access-Control - Allow - Headers: Origin, X - Requested -With, Content - Type, Accept, SilentLaunch
        'Content-Type: application/ json; charset=utf-8
        'Content-Length:  713
        'Date : Sun, 8 Apr 2018 05: 02:33 GMT
        'Connection: keep-alive

        '{
        ' "DUID": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '  "Model": "15_HAWKM_UHD",
        '  "ModelName": "UN55JS8000",
        '  "ModelDescription": "Samsung DTV RCR",
        '  "NetworkType": "wired",
        '  "SSID": "",
        '  "IP": "192.168.1.145",
        '  "FirmwareVersion": "Unknown",
        '  "DeviceName": "[TV] UN55JS8500",
        '  "DeviceID": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '  "UDN": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '  "Resolution": "1920x1080",
        '  "CountryCode": "US",
        '  "SmartHubAgreement": "true",
        '  "ServiceURI": "http://192.168.1.145:8001/ms/1.0/",
        '  "DialURI": "http://192.168.1.145:8001/ws/apps/",
        '  "Capabilities": [
        '    {
        '      "name": "samsung:multiscreen:1",
        '      "port": "8001",
        '      "location": "/ms/1.0/"
        '    }
        '  ]
        '}


        'Someone posted a working version of a command line tool for windows that pairs with the newer tv's and is able to send some keycodes to tizen tv's (at least it worked for me). 
        'You can find the link here. It seems Like he used the smartview dll's from samsung to get this to work so it's unfortunately not directly visible how the pairing is implemented there. 
        'However I sniffed the communication when the tool Is pairing to my tv And when it just sends a command after being paired already And the pattern Is the following
        '
        'For pairing: 
        '
        'GET http://<ip>:8080/ws/apps/CloudPINPage this actually ...this triggers the pin screen on the tv
        'GET http://<ip>:8080/ws/pairing?step=0&app_id=<some_app_id>&device_id=<some_device_id>&type=1 ...this triggers the pin screen on the tv
        'POST http: //<ip>:8080/ws/pairing?step=1&app_id=<some_app_id>&device_id=<some_device_id> with {"auth_Data": {"auth_type": "SPC", "GeneratorServerHello": <pin_something_hash>}} where <pin_something_hash> has a composition As described here
        'POST http: //<ip>:8080/ws/pairing?step=2&app_id=<some_app_id>&device_id=<some_device_id> with {"auth_Data": {"auth_type": "SPC", "request_id": <some_number>, "ServerAckMsg": <some_ack_msg>}}
        'DELETE http: //<ip>:8080/ws/apps/CloudPINPage/run
        '
        'Afterwards A GET request ist made to http: //<ip>:8000/socket.io/1/?t=<timestamp> to receive a websocket connection url Like http://<ip>:8000/socket.io/1/websocket/1AQhMAUMI-mJNiAFpADu?t=<timestamp>. 
        '(When I try to get a websocket connection Like that without pairing the tv socket just hangs up...)
        '
        'When the TV Is paired then, to send a command only step 1 Is made before the websocket connection will be established as described. The problem now Is that I have no idea how the hash Is composed.


        ' Not sure what this does, maybe to verify whether the TV supports this CLOUDPIN thing
        'GET /ws/apps/CloudPINPage HTTP/1.1
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '
        'HTTP/ 1.1 200 OK
        'API-Version: v1.00
        'Content-Type: Text/ html
        'Transfer-Encoding: chunked
        'Date : Wed, 23 Aug 2017 04: 36:30 GMT
        'Server: WebServer
        '
        'fb
        '<?xml version="1.0" encoding="UTF-8"?><service xmlns="urn:dial-multiscreen-org:schemas:dial" xmlns:atom = "http://www.w3.org/2005/Atom" <> Name > CloudPINPage</name><options allowStop="true"/><state>stopped</state><atom:link rel = "run" href="run"/></service>
        '0

        ' https://github.com/eclair4151/samsung_encrypted_POC/blob/master/main.py

        '### Step 0 START
        '       device_id = '12345'
        'step0_pin_url = 'http://' + tv_address + ':8080/ws/apps/CloudPINPage'
        'requests.post(step0_pin_url, Data ='pin4')
        'step0_url = 'http://' + tv_address + ':8080/ws/pairing?step=0&app_id=com.samsung.companion&device_id=12345&type=1'
        'r = requests.get(step0_url) #we can prob ignore this response
        '### Step 0 START

        ' this opens up the PIN page with a PIN
        'POST /ws/apps/CloudPINPage HTTP/1.1
        'CONTENT-Type: application/ json
        'CONTENT-LENGTH:  4
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '
        'pin4


        ' response
        'HTTP/ 1.1 201 Created
        'API-Version: v1.00
        'Content-Type: Text/ html
        'LOCATION: http : ///ws/apps/CloudPINPage/run
        'Transfer-Encoding: chunked
        'Date : Wed, 23 Aug 2017 04: 36:30 GMT
        'Server: WebServer
        '
        '20
        'http:///ws/apps/CloudPINPage/run
        '0


        '### Step 1 START
        '        pin = Input("Enter TV Pin: ")
        '        payload = {'pin': pin, 'payload': '', 'deviceId': device_id}
        'r = requests.post(external_server + '/step1', headers=external_headers, data=json.dumps(payload), verify=False)
        'step1_url = 'http://' + tv_address + ':8080/ws/pairing?step=1&app_id=com.samsung.companion&device_id=12345&type=1'
        'step1_response = requests.post(step1_url, Data = r.text)
        '#### Step 1 End


        'GET /ws/pairing?step=0&app_id=12345&device_id=608e1d92-f42e-4d89-8c25-569da61f2e81&type=1 HTTP/1.1
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '

        'HTTP/ 1.1 200 OK
        'Content-Type: application/ x - javascript; charset=utf-8
        'Content-Length:  18
        'Cache-Control: no-Cache
        'Secure-Mode:  true
        'Date : Wed, 23 Aug 2017 04: 37:14 GMT
        'Server: WebServer
        '
        '{"auth_data"""}
        '
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN called for device - " & MyUPnPDeviceName & " with PIN = " & Pin, LogType.LOG_TYPE_INFO)
        If Pin = "" Then Return ""

        If MySamsungWebSocket IsNot Nothing Then
            If MySamsungWebSocket.WebSocketActive Then
                ' already active
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " already has active websocket", LogType.LOG_TYPE_INFO)
                Return GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungSessionID.ToString, "")
            End If
        End If

        Dim RequestURL As String = ""
        Dim ReturnHeader As String = ""
        Dim ReturnBody As String = ""
        Dim PairingPort As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungPairingPort.ToString, "")
        Dim SamsungAppID As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungAppId.ToString, "")
        Dim GeneratorClientHello As String = ""
        Dim RequestID As String = "0"

        If WeAlreadyHaveAPIN Then GoTo step1
step0:

        RequestURL = "http://" & MyIPAddress & ":" & PairingPort & "/ws/pairing?step=0&app_id=" & SamsungAppID & "&device_id=" & SamsungDeviceID & MyMacAddress & "&type=1"
        ReturnHeader = ""
        ReturnBody = ""

        If Not SendWebRequest(RequestURL, "GET", True, "", ReturnHeader, ReturnBody) Then
            Return ""
        End If

        'In the next message, the GeneratorServerHello is an encryption between the app_id and the PIN
        'POST /ws/pairing?step=1&app_id=12345&device_id=608e1d92-f42e-4d89-8c25-569da61f2e81 HTTP/1.1
        'CONTENT-Type: application/ json
        'CONTENT-LENGTH:  367
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '
        '{"auth_Data"{"auth_type":"SPC","GeneratorServerHello":"010200000000000000008A00000006363534333231C90927F38F66B0633E0B9391A8330914D2FA1B8542A159BF46281C2867FF0734764E60BA10940B6FC325AFA85576CB78D25DE4C88667B2DCEC162CF41886A1DCA83CF28838B12D64A9D288D5A8362380A20BF567B732C233ED1130B73FBC4FC21BAF8B35975EE183781257FB4D3E8979FA8F69B23984F007BB0BF3FAD5AD47C20000000000"}}

step1:

        Dim PayLoad As String = ""
        Dim HelloString As String = ""

        'Dim Response As SPCApiBridge = Nothing 'SpcApiWrapper ' Object = Nothing 'SPCApiBridge = Nothing
        Dim aes_key As Byte() = Nothing
        Dim data_hash As Byte() = Nothing
        Try
            HelloString = GenerateServerHello(SamsungAppID, Pin, aes_key, data_hash)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " trying to Generate a ServerHello with Error = " & ex.ToString, LogType.LOG_TYPE_ERROR)
            Return ""
        End Try

        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN  for device - " & MyUPnPDeviceName & " GeneratorServerHello = " & HelloString, LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN  for device - " & MyUPnPDeviceName & " GeneratorServerHello = " & HelloString, LogType.LOG_TYPE_INFO)  


        RequestURL = "http://" & MyIPAddress & ":" & PairingPort & "/ws/pairing?step=1&app_id=" & SamsungAppID & "&device_id=" & SamsungDeviceID & MyMacAddress  ' & "&type=1"
        ReturnHeader = ""
        ReturnBody = ""
        PayLoad = "{""auth_Data"":{""auth_type"":""SPC"",""GeneratorServerHello"":""" & HelloString & """}}"
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN  for device - " & MyUPnPDeviceName & "  Sending GeneratorServerHello with payload = " & PayLoad, LogType.LOG_TYPE_INFO)

        If Not SendWebRequest(RequestURL, "POST", True, PayLoad, ReturnHeader, ReturnBody) Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Unsuccessful GeneratorServerHello", LogType.LOG_TYPE_INFO)
            Return ""
        End If

        ' Parse the response, lift out the GeneratorClientHello
        ' Expected Response
        'HTTP/ 1.1 200 OK
        'Content-Type: application/ x - javascript; charset=utf-8
        'Content-Length:  440
        'Cache-Control: no-Cache
        'Secure-Mode:  true
        'Date : Wed, 23 Aug 2017 04: 37:15 GMT
        'Server: WebServer
        '
        '{"auth_data""{\"auth_type\":\"SPC\",\"request_id\":\"0\",\"GeneratorClientHello\":\"010100000000000000009E00000006363534333231B47329EB3F22BA588116C73CD6937FCA2B4989A6A1FBB9D24A9669AAD8F4982901BA566CE5DDA943BC07EE4A0815C7107147979A6B64E27A716822E0772ECD1181E398D003EE0350F50D7B8D2A14877E2ED439CFBCC1787B7DA69C1C23F079375A2D4A9ABF3474B758BABAE2AEE1E42E43EE1634B6C11A89636FC99BBD6F08B372FD8BF4278C1B94022004B4AAC95A58A7142B0E0000000000\"}"}


        Dim AuthData As String = FindPairInJSONString(ReturnBody, "auth_data").ToString
        If AuthData <> "" Then
            GeneratorClientHello = FindPairInJSONString(AuthData, "GeneratorClientHello").ToString.Trim("""")
            RequestID = FindPairInJSONString(AuthData, "request_id").ToString.Trim("""")
        End If


        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " GeneratorClientHello = " & GeneratorClientHello, LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Pin = " & Pin, LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " GeneratorClientHello = " & GeneratorClientHello, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Pin = " & Pin, LogType.LOG_TYPE_INFO)

        If GeneratorClientHello = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " GeneratorClientHello is empty. Response was = " & ReturnBody, LogType.LOG_TYPE_INFO)
            ' I'm going to interpret this as a bad PIN and reset the PIN
            SamsungResetIdentityParms()
            Return ""
        End If


        If Pin = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Pin is empty", LogType.LOG_TYPE_INFO)
            Return ""
        End If

        Dim skPrime As Byte() = Nothing

        Dim AESKey As Byte() = Nothing
        Try
            AESKey = ParseClientHello(GeneratorClientHello, data_hash, aes_key, SamsungAppID, skPrime)
        Catch ex As Exception
            Return ""
        End Try

        '### Step 2 START
        'payload = {'pin': pin, 'payload': codecs.decode(step1_response.text, 'unicode_escape'), 'deviceId': device_id}
        'r = requests.post(external_server + '/step2', data=json.dumps(payload), headers=external_headers, verify=False)
        'step2_url = 'http://' + tv_address + ':8080/ws/pairing?step=2&app_id=com.samsung.companion&device_id=12345&type=1&request_id=0'
        'step2_response = requests.post(step2_url, Data = r.text)
        '### Step 2 End

        'POST /ws/pairing?step=2&app_id=12345&device_id=608e1d92-f42e-4d89-8c25-569da61f2e81 HTTP/1.1
        'CONTENT-Type: application/ json
        'CONTENT-LENGTH:  140
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '
        '{"auth_Data"{"auth_type":"SPC","request_id":"0","ServerAckMsg":"0103000000000000000014802137319B9785A3FE78EF19E1997F86C96DCE8E0000000000"}}


        Dim ServerAckMsg As String = ""
        ServerAckMsg = generateServerAcknowledge(skPrime)
        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungServerAckMsg.ToString, ServerAckMsg)

Step2:

        RequestURL = "http://" & MyIPAddress & ":" & PairingPort & "/ws/pairing?step=2&app_id=" & SamsungAppID & "&device_id=" & SamsungDeviceID & MyMacAddress

        ReturnHeader = ""
        ReturnBody = ""

        ServerAckMsg = GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungServerAckMsg.ToString, "")
        Try
            PayLoad = "{""auth_Data"":{""auth_type"":""SPC"",""request_id"":""" & RequestID & """,""ServerAckMsg"":""" & ServerAckMsg & """}}"
        Catch ex As Exception
            Return ""
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " is sending ServerAckMsg with content = " & PayLoad, LogType.LOG_TYPE_INFO)

        If Not SendWebRequest(RequestURL, "POST", True, PayLoad, ReturnHeader, ReturnBody) Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " unsuccessful sending ServerAckMsg", LogType.LOG_TYPE_INFO)
            Return ""
        End If

        ' Expected Response

        'HTTP/ 1.1 200 OK
        'Content-Type: application/ x - javascript; charset=utf-8
        'Content-Length:  177
        'Cache-Control: no-Cache
        'Secure-Mode:  true
        'Date : Wed, 23 Aug 2017 04: 37:16 GMT
        'Server: WebServer
        '
        '{"auth_data""{\"auth_type\":\"SPC\",\"request_id\":\"0\",\"ClientAckMsg\":\"010400000000000000001454D210EC50DFB5F189043D0DCD3BFEAE4A16BCBE0000000000\",\"session_id\":\"1\"}"}

        ' Parse the response, lift out the ClientAckMsg
        Dim ClientAckMsg As String = ""
        Dim SessionID As String = ""
        Dim SessionKey As String = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " received authentication data for ClientAck with content = " & ReturnBody, LogType.LOG_TYPE_INFO)

        AuthData = FindPairInJSONString(ReturnBody, "auth_data").ToString
        If AuthData IsNot Nothing AndAlso AuthData <> "" Then
            ClientAckMsg = FindPairInJSONString(AuthData, "ClientAckMsg").ToString.Trim("""")
            SessionID = FindPairInJSONString(AuthData, "session_id").ToString.Trim("""")
            SessionKey = FindPairInJSONString(AuthData, "session_key").ToString.Trim("""")
        End If

        If SessionID IsNot Nothing Then SessionID = ""

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " ClientAckMsg = " & ClientAckMsg, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " SessionID = " & SessionID, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " SessionKey = " & SessionKey, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " AESKey = " & ByteArrayToHexString(AESKey), LogType.LOG_TYPE_INFO)

        If ClientAckMsg Is Nothing Or ClientAckMsg = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Unsuccessful ParseClientAck", LogType.LOG_TYPE_INFO)
            Return ""
        End If

        Try
            parseClientAcknowledge(ClientAckMsg, skPrime)
        Catch ex As Exception
            Return ""
        End Try


        If AESKey Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " Unsuccessful retrieval of Key", LogType.LOG_TYPE_INFO)
            Return ""
        End If
        SamsungAESKey = AESKey

        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " AESKey = " & System.Text.ASCIIEncoding.ASCII.GetChars(AESKey), LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " AESKey Length = " & AESKey.Length.ToString, LogType.LOG_TYPE_INFO)

        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungAesKey.ToString, Convert.ToBase64String(AESKey))
        WriteStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungSessionID.ToString, SessionID)

        'GET /common/1.0.0/service/startService?appID=com.samsung.companion HTTP/1.1
        'HOST: 192.168.1.165:8000
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '

        'RequestURL = "http://192.168.1.145:8080/common/1.0.0/service/startService?appID=com.samsung.companion"
        ' Maybe this is where we use the Capabilityname etc rather than com.samsung.companion
        '            http://192.168.1.145:8001/ms/1.0/
        'RequestURL = "http://192.168.1.145:8001/ms/1.0/service/startService?appID=" & AppID 
        'ReturnHeader = ""
        'ReturnBody = ""

        'If Not SendWebRequest(RequestURL, "POST", True, PayLoad, ReturnHeader, ReturnBody) Then
        'If Response IsNot Nothing Then
        'Try
        'Response.Dispose()
        'Catch ex As Exception
        'End Try
        'End If
        'Console.Write("Unsuccessful Post start service. Enter to Exit - ")
        ''Console.ReadLine()
        ''Exit Sub
        'End If

        'HTTP/ 1.1 404 Not Found
        'X-Powered - By: Express
        'Access-Control - Allow - Origin:  *
        'Access-Control - Allow - Methods:  GET,PUT,POST,DELETE,OPTIONS
        'Access-Control - Allow - Headers: Content-Type, Authorization
        'Content-Type: Text/ plain
        'Date : Wed, 23 Aug 2017 04: 37:16 GMT
        'Connection: keep-alive
        'Transfer-Encoding: chunked
        '
        '49
        'Cannot GET /common/1.0.0/service/startService?appID=com.samsung.companion
        '0



        '### Step 3 START
        'payload = {'pin': pin, 'payload': codecs.decode(step2_response.text, 'unicode_escape'), 'deviceId': device_id}
        'r = requests.post(external_server + '/step3', data=json.dumps(payload), headers=external_headers, verify=False)
        'enc_key = r.json()['session_key']
        'session = r.json()['session_id']
        'Print('session_key: ' + enc_key)
        'Print('session_id: ' + session)
        'step3_url = 'http://' + tv_address + ':8080/ws/apps/CloudPINPage/run'
        'requests.delete(step3_url)
        '### Step 3 End

        'DELETE /ws/apps/CloudPINPage/run HTTP/1.1
        'HOST: 192.168.1.165:8080
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1

        If Not WeAlreadyHaveAPIN Then
            ' Next step close the PIN page

            'DELETE /ws/apps/CloudPINPage/run HTTP/1.1
            'HOST: 192.168.1.165:8080
            'CONNECTION: Keep-Alive
            'USER-AGENT: httpclient v0.1
            '
            'HTTP/ 1.1 200 OK
            'API-Version: v1.00
            'Content-Length:  0
            'Date : Wed, 23 Aug 2017 04: 37:15 GMT
            'Server: WebServer

            RequestURL = "http://" & MyIPAddress & ":" & PairingPort & "/ws/apps/CloudPINPage/run"
            ReturnHeader = ""
            ReturnBody = ""
            PayLoad = ""

            If Not SendWebRequest(RequestURL, "DELETE", True, PayLoad, ReturnHeader, ReturnBody) Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungAuthenticateUsePIN for device - " & MyUPnPDeviceName & " unsuccessful Delete", LogType.LOG_TYPE_INFO)
                Return ""
            End If

            ' Expected response
            '
            'HTTP/ 1.1 200 OK
            'API-Version: v1.00
            'Content-Length:  0
            'Date : Wed, 23 Aug 2017 04: 37:16 GMT
            'Server: WebServer

        End If

        If SessionID IsNot Nothing AndAlso SessionID <> "" Then
            SamsungSessionID = SessionID
            If SamsungAESKey IsNot Nothing Then
                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)
                SetHSRemoteState()
            End If
        End If
        Return SessionID
    End Function

    Public Function SamsungOpenEncryptedWebSocket(SessionID As String) As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & " called with SessionId = " & SessionID, LogType.LOG_TYPE_INFO)
        If SamsungAESKey Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & " has no AESKey", LogType.LOG_TYPE_INFO)
            Return False
        Else
            If SamsungAESKey.Length = 0 Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & " has no AESKey", LogType.LOG_TYPE_INFO)
                Return False
            End If
        End If
        If MySamsungWebSocket IsNot Nothing Then
            If MySamsungWebSocket.WebSocketActive Then
                ' already active
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & " already has active websocket", LogType.LOG_TYPE_INFO)
                Return True
            End If
        End If

        '
        'Afterwards A GET request ist made to http: //<ip>:8000/socket.io/1/?t=<timestamp> to receive a websocket connection url Like http://<ip>:8000/socket.io/1/websocket/1AQhMAUMI-mJNiAFpADu?t=<timestamp>. 
        '(When I try to get a websocket connection Like that without pairing the tv socket just hangs up...)
        '
        'When the TV Is paired then, to send a command only step 1 Is made before the websocket connection will be established as described. The problem now Is that I have no idea how the hash Is composed.

        ' https://github.com/eclair4151/samsung_encrypted_POC/blob/master/main.py
        ' https://github.com/eclair4151/SmartCrypto

        'GET /common/1.0.0/service/startService?appID=com.samsung.companion HTTP/1.1
        'HOST: 192.168.1.165:8000
        'CONNECTION: Keep-Alive
        'USER-AGENT: httpclient v0.1
        '

        'RequestURL = "http://192.168.1.145:8080/common/1.0.0/service/startService?appID=com.samsung.companion"
        ' Maybe this is where we use the Capabilityname etc rather than com.samsung.companion
        '            http://192.168.1.145:8001/ms/1.0/
        'RequestURL = "http://192.168.1.145:8001/ms/1.0/service/startService?appID=" & AppID 
        'ReturnHeader = ""
        'ReturnBody = ""


        '### Step 3 START
        'payload = {'pin': pin, 'payload': codecs.decode(step2_response.text, 'unicode_escape'), 'deviceId': device_id}
        'r = requests.post(external_server + '/step3', data=json.dumps(payload), headers=external_headers, verify=False)
        'enc_key = r.json()['session_key']
        'session = r.json()['session_id']
        'Print('session_key: ' + enc_key)
        'Print('session_id: ' + session)
        'step3_url = 'http://' + tv_address + ':8080/ws/apps/CloudPINPage/run'
        'requests.delete(step3_url)
        '### Step 3 End


        ' OK we need a Handshake token
        ' do a get on :8000/socket.io/1
        'millis = Int(round(DateTime.t() * 1000))
        'step4_url = 'http://' + tv_address + ':8000/socket.io/1/?t=' + str(millis)
        Dim WebSocketPort As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "")
        Dim WebSocketURL As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "")
        If WebSocketPort = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & "; no WebSocketPort info", LogType.LOG_TYPE_INFO)
            Return False
        End If
        Dim RequestURL As String = "http://" & MyIPAddress & ":" & WebSocketPort & "/" & WebSocketURL
        Dim ReturnHeader As String = ""
        Dim ReturnBody As String = ""
        Dim PayLoad As String = ""

        If Not SendWebRequest(RequestURL, "GET", True, PayLoad, ReturnHeader, ReturnBody) Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for device - " & MyUPnPDeviceName & " unsuccessful get handshakeToken", LogType.LOG_TYPE_INFO)
            Return False
        End If

        ' the result is something like this
        ' w2cibNduEVG4rwLuAH7P:60:60:websocket,htmlfile,xhr-polling,jsonp-polling

        Dim ReturnParts As String() = Split(ReturnBody, ":")
        Dim HandshakeToken As String = ""
        If ReturnParts(0) IsNot Nothing Then
            HandshakeToken = ReturnParts(0)
        End If

        'GET /socket.io/1/websocket/6aWAnGkG4SEayQH6bdxN HTTP/1.1
        'Host: 192.168.1.217:8000
        'Sec-WebSocket-Version: 13
        'Upgrade: websocket
        'Sec-WebSocket-Key: d7hgP7MnHNFRbmS9tAztsA==
        'Connection: Upgrade
        'Origin: Http://192.168.1.217:8000

        'Http/ 1.1 101 Switching Protocols
        'Upgrade: websocket
        'Connection: Upgrade
        'Sec-WebSocket-Accept: XScOKR/zVCww0GzlE9AfkwUj9LM=



        '## Step 4 START   WEBSOCKETS
        'millis = Int(round(time.time() * 1000))
        'step4_url = 'http://' + tv_address + ':8000/socket.io/1/?t=' + str(millis)
        'websocket_response = requests.get(step4_url)
        'websocket_url = 'ws://' + tv_address + ':8000/socket.io/1/websocket/' + websocket_response.text.split(':')[0]


        'time.sleep(1)
        'Print('sending command!')
        'aesLib = aes_lib.AESCipher(enc_key, session)
        'Connection = websocket.create_connection(websocket_url)
        'time.sleep(0.35)
        'Connection.send('1::/com.samsung.companion')
        'time.sleep(0.35)
        'r = Connection.send(aesLib.generate_command('KEY_VOLDOWN'))
        'time.sleep(0.35)
        'Connection.Close()
        'Print('sent')

        '## Step 4 End

        If MySamsungWebSocket Is Nothing Then
            Try
                MySamsungWebSocket = New WebSocketClient(False)
            Catch ex As Exception
                Log("Error in SamsungOpenEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " unable to open WebSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyRemoteServiceActive = False
                Return False
            End Try
        End If

        If Not MySamsungWebSocket.ConnectSocket(MyIPAddress, GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "")) Then
            MySamsungWebSocket = Nothing
            Return False
        End If

        ' wait until connected. Important for SSL as it takes longer
        Dim WaitLoopCounter As Integer = 0
        While MySamsungWebSocket.MySocketIsClosed
            wait(1)
            WaitLoopCounter += 1
            If WaitLoopCounter > 10 Then Exit While
        End While

        If Not MySamsungWebSocket.Receive() Then
            Try
                MySamsungWebSocket.CloseSocket()
            Catch ex As Exception
            End Try
            MySamsungWebSocket = Nothing
            Return False
        End If

        If Not MySamsungWebSocket.UpgradeWebSocket(WebSocketURL & "websocket/" & HandshakeToken, GetStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, ""), 0, False) Then
            Try
                MySamsungWebSocket.CloseSocket()
            Catch ex As Exception
            End Try
            MySamsungWebSocket = Nothing
            Return False
        End If

        AddHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
        AddHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " added handlers", LogType.LOG_TYPE_INFO)


        Dim WaitForConnect As Integer
        While WaitForConnect < 10
            If Not MySamsungWebSocket.WebSocketActive Then
                wait(1)
            Else
                Exit While
            End If
            WaitForConnect += 1
        End While

        If MySamsungWebSocket Is Nothing Then
            Return False
        End If

        If Not MySamsungWebSocket.WebSocketActive Then
            Return False
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " has now an active remote", LogType.LOG_TYPE_INFO)

        MyRemoteServiceActive = True
        Return True
    End Function

    Private Sub SamsungOpenUnEncryptedWebSocket(WebSocketPort As String, Location As String)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket called for UPnPDevice = " & MyUPnPDeviceName & ", WebSocketPort = " & WebSocketPort & ", location = " & Location & " and MyRemoteServiceActive = " & MyRemoteServiceActive, LogType.LOG_TYPE_INFO)


        'key in http://192.168.1.165:8001/api/v2/ and get
        '{
        ' "id": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        ' "name": "[TV] UN55JS8500",
        ' "version": "2.0.25",
        ' "device": {
        '   "type": "Samsung SmartTV",
        '   "duid": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '   "model": "15_HAWKM_UHD",
        '   "modelName": "UN55JS8000",
        '   "description": "Samsung DTV RCR",
        '   "networkType": "wired",
        '   "ssid": "",
        '   "ip": "192.168.1.165",
        '   "firmwareVersion": "Unknown",
        '   "name": "[TV] UN55JS8500",
        '   "id": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '   "udn": "uuid:3b76f941-bbdd-493a-b837-b351deb09555",
        '   "resolution": "1920x1080",
        '   "countryCode": "US",
        '   "msfVersion": "2.0.25",
        '   "smartHubAgreement": "true",
        '   "wifiMac": "fc:f1:36:dc:a8:d9",
        '   "developerMode": "0",
        '   "developerIP": ""
        '},
        '"type": "Samsung SmartTV",
        '"uri": "http://192.168.1.165:8001/api/v2/"
        '}
        ' {"device":{"FrameTVSupport":"false","GamePadSupport":"true","ImeSyncedSupport":"true","OS":"Tizen","TokenAuthSupport":"true","VoiceSupport":"true","countryCode":"NO","description":"Samsung DTV RCR","developerIP":"0.0.0.0","developerMode":"0","duid":"uuid:c15fc058-0ab1-4c8d-80ca-b3f11d81e291","firmwareVersion":"Unknown","id":"uuid:c15fc058-0ab1-4c8d-80ca-b3f11d81e291","ip":"192.168.2.247","model":"18_KANTM2_QTV","modelName":"QE65Q9FNA","name":"[TV] Samsung Q9 Series (65)","networkType":"wired","resolution":"3840x2160","smartHubAgreement":"true","type":"Samsung SmartTV","udn":"uuid:c15fc058-0ab1-4c8d-80ca-b3f11d81e291","wifiMac":"c0:48:e6:c3:3b:8d"},"id":"uuid:c15fc058-0ab1-4c8d-80ca-b3f11d81e291","isSupport":"{"DMP_DRM_PLAYREADY":"false","DMP_DRM_WIDEVINE":"false","DMP_available":"true","EDEN_available":"true","FrameTVSupport":"false","ImeSyncedSupport":"true","TokenAuthSupport":"true","remote_available":"true","remote_fourDirections":"true","remote_touchPad":"true","remote_voiceControl":"true"}","name":"[TV] Samsung Q9 Series (65)","remote":"1.0","type":"Samsung SmartTV","uri":"http://192.168.2.247:8001/api/v2/","version":"2.0.25"}

        Dim URL As String = "http://" & MyIPAddress & ":" & WebSocketPort & "/api/v2/"
        Dim RequestUri = New Uri(URL)
        Dim StreamText As String = ""
        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = "GET"
            wRequest.KeepAlive = False
            wRequest.ProtocolVersion = HttpVersion.Version11
            Dim webResponse As WebResponse = Nothing
            webResponse = wRequest.GetResponse
            Dim reader As New StreamReader(webResponse.GetResponseStream())
            StreamText = reader.ReadToEnd()
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & " from URL = " & URL & " retrieving info = " & StreamText, LogType.LOG_TYPE_INFO)
            reader.Close()
            webResponse.Close()
        Catch ex As Exception
            Log("Error in SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & " retrieving info at URL = " & URL.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If StreamText <> "" Then
                Dim DeviceData As Object = FindPairInJSONString(StreamText, "device")
                If DeviceData IsNot Nothing Then
                    Dim TokenAuthSupport As Object = FindPairInJSONString(DeviceData, "TokenAuthSupport")
                    ' not sure this will be a boolean or text representing a boolean
                    If TokenAuthSupport IsNot Nothing Then
                        If TokenAuthSupport.ToString.ToLower = "true" Then
                            ' set token to be found. Maybe change port to 8002.
                            WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diSamsungTokenAuthSupport.ToString, True)
                            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "8002")  ' this will force the SSL port to be used
                        End If
                    End If
                    ' retrieve wifiMac
                    Dim WifiMac As Object = FindPairInJSONString(DeviceData, "wifiMac")
                    If WifiMac IsNot Nothing Then
                        WriteStringIniFile(MyUDN, DeviceInfoIndex.diWifiMacAddress.ToString, WifiMac.ToString)
                    End If
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & " parsing /api/v2/ response with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If StreamText <> "" Then
                Dim SupportData As Object = FindPairInJSONString(StreamText, "isSupport")
                If SupportData IsNot Nothing Then
                    WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungisSupportInfo.ToString, SupportData.ToString)
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & " parsing isSupport info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try


        ' TEST dcortizen
        ' Smartthings and smartview seem to be hunting for this ExtraService app.
        ' when not found, Smartview hunts for IDs = 3201503001633, 3201506003502, 3201411000404 -> the latter which results in Enhancement App
        'URL = "http://" & MyIPAddress & ":" & WebSocketPort & "/api/v2/applications/ExtraService"

        'RequestUri = New Uri(URL)
        'StreamText = ""
        'Try
        'Dim p = ServicePointManager.FindServicePoint(RequestUri)
        'p.Expect100Continue = False
        'Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
        'wRequest.Method = "GET"
        'wRequest.KeepAlive = False
        'wRequest.ProtocolVersion = HttpVersion.Version11
        'Dim webResponse As WebResponse = Nothing
        'WebResponse = wRequest.GetResponse
        'Dim reader As New StreamReader(webResponse.GetResponseStream())
        'StreamText = reader.ReadToEnd()
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket (TEST) for device - " & MyUPnPDeviceName & " retrieving info = " & StreamText, LogType.LOG_TYPE_INFO)
        'reader.Close()
        'WebResponse.Close()
        'Catch ex As Exception
        'Log("Error in SamsungOpenUnEncryptedWebSocket (TEST) for device - " & MyUPnPDeviceName & " retrieving info at URL = " & URL.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        'End Try

        ' Send key  URL_FORMAT = "ws://{}:{}/api/v2/channels/samsung.remote.control?name={}"

        ' send something:        
        'payload = json.dumps({
        '    "method" "ms.remote.control",
        '    "params": {
        '        "Cmd": "Click",
        '        "DataOfCmd": key,
        '        "Option":  "false",
        '        "TypeOfRemote": "SendRemoteKey"
        '    }
        '})

        ' also this returns info http://192.168.1.165:8001/ws/apps/
        ' I need to start hunting for this, it is part of the device XML
        '   <sec:Capabilities<> sec: Capability Name = "samsung:multiscreen:1" port="8001" location="/ms/1.0/"/></sec:Capabilities>
        '   <sec:ProductCap>Tizen,Y2015,WebURIPlayable,NavigateInPause,ScreenMirroringP2PMAC=7a:bd:bc:71:33:76</sec:ProductCap>


        ' this could be the JSON way of registering
        ' GET /api/v2/channels/samsung.remote.control?name=cmVtb3Rl
        ' read this posting https://github.com/Ape/samsungctl/issues/38


        Dim WebSocketKey As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, "")
        Dim WebSocketURL As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, "")
        If WebSocketPort = "" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & "; no WebSocketPort info", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        Dim SSLRequired As Boolean = GetBooleanIniFile(MyUDN, DeviceInfoIndex.diSamsungTokenAuthSupport.ToString, False)
        Dim SSLAuthToken As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungToken.ToString, "")
        If MySamsungWebSocket Is Nothing Then
            Try
                MySamsungWebSocket = New WebSocketClient(SSLRequired)
            Catch ex As Exception
                Log("Error in SamsungOpenUnEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " unable to open WebSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyRemoteServiceActive = False
                Exit Sub
            End Try
        End If

        'moved here in v32 from 2 blocks down
        AddHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
        AddHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " added handlers", LogType.LOG_TYPE_INFO)


        If Not MySamsungWebSocket.ConnectSocket(MyIPAddress, GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, "")) Then
            Try
                RemoveHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
            Catch ex As Exception
            End Try
            MySamsungWebSocket = Nothing
            Exit Sub
        End If

        ' wait until connected. Important for SSL as it takes longer
        Dim WaitLoopCounter As Integer = 0
        While MySamsungWebSocket.MySocketIsClosed
            wait(1)
            WaitLoopCounter += 1
            If WaitLoopCounter > 10 Then Exit While
        End While

        If Not MySamsungWebSocket.Receive() Then
            Try
                MySamsungWebSocket.CloseSocket()
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
            Catch ex As Exception
            End Try
            MySamsungWebSocket = Nothing
            Exit Sub
        End If

        If SSLRequired And (SSLAuthToken <> "") Then
            WebSocketURL = WebSocketURL & "&token=" & SSLAuthToken
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket for device - " & MyUPnPDeviceName & " added token = " & WebSocketURL, LogType.LOG_TYPE_INFO)
        End If

        If Not MySamsungWebSocket.UpgradeWebSocket(WebSocketURL, GetStringIniFile(MyUDN, DeviceInfoIndex.diSecWebSocketKey.ToString, ""), 0, False) Then
            MySamsungWebSocket.CloseSocket()
            Try
                RemoveHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
            Catch ex As Exception
            End Try
            MySamsungWebSocket = Nothing
            Exit Sub
        End If

        MyRemoteServiceActive = True

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungOpenUnEncryptedWebSocket for UPnPDevice = " & MyUPnPDeviceName & " has now an active remote", LogType.LOG_TYPE_INFO)

        ' This is what comes back
        ' HTTP/1.1 101 Switching Protocols
        ' Upgrade: websocket
        ' Connection: Upgrade
        ' Sec-WebSocket-Accept: G/cEt4HtsYEnP0MnSVkKRk459gM=
        ' Sec-WebSocket-Protocol: dumb-increment-protocol

        ' get all apps
        ' {"method":"ms.channel.emit","params":{"event": "ed.installedApp.get", "to":"host"}}
        ' Launch them
        ' {"method":"ms.channel.emit","params":{"event": "ed.apps.launch", "to":"host", "data":{"appId":"org.tizen.browser","action_type":"NATIVE_LAUNCH","metaTag":"http:\/\/hackaday.com"}}}
        ' ws.onopen = function() {ws.send('{"method":"ms.channel.emit","params":{"event": "ed.apps.launch", "to":"host", "data":{"appId":"bstjKvX6LM.molotov","action_type":"DEEP_LINK"}}}')};

    End Sub

    Public Function SamsungSendEncryptedKeyCode(AESKey As Byte(), SessionID As String, KeyCode As String) As Boolean

        '
        'Afterwards A GET request ist made to http: //<ip>:8000/socket.io/1/?t=<timestamp> to receive a websocket connection url Like http://<ip>:8000/socket.io/1/websocket/1AQhMAUMI-mJNiAFpADu?t=<timestamp>. 
        '(When I try to get a websocket connection Like that without pairing the tv socket just hangs up...)
        '
        'When the TV Is paired then, to send a command only step 1 Is made before the websocket connection will be established as described. The problem now Is that I have no idea how the hash Is composed.

        ' https://github.com/eclair4151/samsung_encrypted_POC/blob/master/main.py

        Dim RequestURL As String = ""
        Dim ReturnHeader As String = ""
        Dim ReturnBody As String = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendEncryptedKeyCode called for UPnPDevice = " & MyUPnPDeviceName & " with KeyCode = " & KeyCode, LogType.LOG_TYPE_INFO)
        If MySamsungWebSocket Is Nothing Then
            ' either open or go back
            Return False
        End If

        Dim KeyCodeString As String = "{""method"":""POST"",""body"":{""plugin"":""RemoteControl"",""param1"":""uuid:" & SamsungDeviceID & MyMacAddress & """,""param2"":""Click"",""param3"":""" & KeyCode & """,""param4"":false,""api"":""SendRemoteKey"",""version"":""1.000""}}"
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungSendEncryptedKeyCode for device - " & MyUPnPDeviceName & " is composing KeyCodeString = " & KeyCodeString, LogType.LOG_TYPE_INFO)

        Dim EncryptedKeyCodeBytes As Byte() = AESE(System.Text.ASCIIEncoding.ASCII.GetBytes(KeyCodeString), AESKey)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendEncryptedKeyCode for device - " & MyUPnPDeviceName & " is sending EncryptedKeyCodeBytes Length = " & EncryptedKeyCodeBytes.Length.ToString, LogType.LOG_TYPE_INFO)

        Dim Body As String = "["
        For i = 0 To EncryptedKeyCodeBytes.Length - 1
            Body += EncryptedKeyCodeBytes(i).ToString + ","
        Next

        Body = Body.Remove(Body.Length - 1, 1) + "]"
        Dim SocketData As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes("5::/com.samsung.companion:{""name"":""callCommon"",""args"":[{""Session_Id"":" & SessionID & ",""body"":""" & Body & """}]}")

        If PIDebuglevel > DebugLevel.dlEvents Then
            Log("SamsungSendEncryptedKeyCode for device - " & MyUPnPDeviceName & " is sending CommandInfo = " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)
            Log("SamsungSendEncryptedKeyCode for device - " & MyUPnPDeviceName & " is sending CommandInfo Length = " & SocketData.Length, LogType.LOG_TYPE_INFO)
        End If
        If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendEncryptedKeyCode for device - " & MyUPnPDeviceName & " unsuccessful send Key = " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)
            Return False
        End If

        Return True

    End Function

    Public Function SendWebRequest(RequestURL As String, Method As String, KeepAlive As Boolean, Payload As String, ByRef ReturnHeader As String, ByRef ReturnBody As String) As Boolean

        Dim wRequest As HttpWebRequest = Nothing
        Dim xmlDoc As New XmlDocument
        Dim Data As Byte() = Nothing
        SendWebRequest = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendWebRequest called for Device = " & MyUPnPDeviceName & " with Method = " & Method & " and RequestURL = " & RequestURL, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendWebRequest for device - " & MyUPnPDeviceName & " called with URL = " & RequestURL & ", Method = " & Method & ", Payload =  " & Payload, LogType.LOG_TYPE_INFO)

        If Payload.Length > 0 Then
            Data = System.Text.ASCIIEncoding.ASCII.GetBytes(Payload)
        End If

        Try
            Dim RequestUri = New Uri(RequestURL)
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest)
            wRequest.Method = Method
            wRequest.KeepAlive = KeepAlive
            wRequest.AllowAutoRedirect = True
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.ContentType = "application/json"
            If Payload.Length > 0 And Data IsNot Nothing Then
                wRequest.ContentLength = Data.Length
            Else
                wRequest.ContentLength = 0
            End If
            wRequest.UserAgent = "httpclient v0.1"
            If Payload.Length > 0 Then
                Dim stream = wRequest.GetRequestStream()
                stream.Write(Data, 0, Data.Length)
                stream.Close()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendWebRequest for device - " & MyUPnPDeviceName & " creating a webrequest with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try


        Dim webResponse As HttpWebResponse = Nothing
        Dim webStream As Stream = Nothing

        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ReturnBody = strmRdr.ReadToEnd()
                ReturnHeader = webResponse.Headers.ToString()
                strmRdr.Dispose()
                webStream.Dispose()
                Try
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendWebRequest for device - " & MyUPnPDeviceName & " has Response Header = " & ReturnHeader, LogType.LOG_TYPE_INFO)
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendWebRequest for device - " & MyUPnPDeviceName & " has response body = " & ReturnBody, LogType.LOG_TYPE_INFO)
                Catch ex1 As Exception
                End Try
            End If
            Log("Error in SendWebRequest for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message & " and Response = " & ReturnBody, LogType.LOG_TYPE_ERROR)
            Return False
        Catch ex As Exception
            Log("Error in SendWebRequest for device - " & MyUPnPDeviceName & " doing a GetResponse with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Try
            ReturnHeader = webResponse.Headers.ToString()
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SendWebRequest for device - " & MyUPnPDeviceName & " has Response Header = " & ReturnHeader, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try

        Try
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ReturnBody = strmRdr.ReadToEnd()
                strmRdr.Dispose()
                webStream.Dispose()
                Try
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("SendWebRequest for device - " & MyUPnPDeviceName & " has Response Header = " & ReturnBody, LogType.LOG_TYPE_INFO)
                Catch ex1 As Exception
                End Try
            End If
        Catch ex As Exception

        End Try

        Return True

    End Function

    Private Function InterpretResult(ReturnString As String) As String
        InterpretResult = ""
        Dim ascii As Encoding = Encoding.UTF8
        Dim ReturnBytes As [Byte]() = ascii.GetBytes(ReturnString)
        '    ---
        '<Byte - unknown byte>
        '     0x02 = ???
        '<Byte - Length of string1>
        '0x00
        '<x-Bytes - string1>
        '     "iapp.samsung"
        '<Byte - Length of resultbytes>
        '0x00
        '<resultbytes>
        '     0xC8 0x00 0x02 0x00 0x31 0x00  =  ??
        '---
        '<Byte - unknown byte>
        '      0x02 = ???
        '<Byte - Length of string1>
        '0x00
        '<x-Bytes - string1>
        '    "iapp.samsung"
        '<Byte - length of resultbytes>
        '0x00
        '<resultbytes>
        '     0x03 0x00 0x04 0x00 0x41 0x41 0x3D 0x3D  =  ??
        '---
    End Function


    Public Sub SamsungSendMessage(MessageID As String, Message As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMessage called with MessageID = " & MessageID & " and Message = " & Message.ToString & " while service active = " & MyMessageServiceActive.ToString, LogType.LOG_TYPE_INFO)
        If Not MyMessageServiceActive Then
            Exit Sub
        End If
        '    Message XML structure is parsed as following:

        '<Category>Body</Category> 
        'is checked to contain "SMS", "Incoming Call", "Schedule Reminder" and "MMS". However, "MMS" category is not implemented in TV.

        '<DisplayType>Body</DisplayType>
        'is checked if contains "Maximum". If yes, detailed message is displayed on the TV, otherwise only short info is displayed.

        'According to the Category type, different tags are used to retrieve content:

        '"SMS" Category
        '==============
        ' <ReceiveTime>
        '   <Date>YYYY-MM-DD</Date>		(YYYY - year, MM - month, DD - day)
        '   <Time>HH:MM:SS</Time>		(HH - hour, MM - minute, SS - second)
        ' </ReceiveTime>
        ' <Receiver>
        '   <Number>Receiver Phone Number</Number>
        '   <Name>Receiver Name</Name>
        ' </Receiver>
        ' <Sender>
        '   <Number>Sender Phone Number</Number>
        '   <Name>Sender Name</Name>
        ' </Sender>
        ' <Body>SMS Body</Body>
        '
        'Sender and Body is displayed only in case DisplayType is set to "Maximum".


        '"Incoming Call" Category
        '========================
        ' <CallTime>
        '   <Date>YYYY-MM-DD</Date>		(YYYY - year, MM - month, DD - day)
        '   <Time>HH:MM:SS</Time>		(HH - hour, MM - minute, SS - second)
        ' </CallTime>
        ' <Callee>
        '   <Number>Callee Phone Number</Number>
        '   <Name>Callee Name</Name>
        ' </Callee>
        ' <Caller>
        '   <Number>Caller Phone Number</Number>
        '   <Name>Caller Name</Name>
        ' </Caller>

        ' <Category>Incoming Call</Category><DisplayType>Maximum</DisplayType><CallTime><Date>YYYY-MM-DD</Date><Time>HH:MM:SS</Time></CallTime><Callee><Number>Callee Phone Number</Number><Name>Callee Name</Name></Callee><Caller><Number>Caller Phone Number</Number><Name>Caller Name</Name></Caller>

        'Caller is displayed only in case DisplayType is set to "Maximum".

        '"Schedule Reminder" Category
        '============================
        ' <StartTime>
        '   <Date>YYYY-MM-DD</Date>		(YYYY - year, MM - month, DD - day)
        '   <Time>HH:MM:SS</Time>		(HH - hour, MM - minute, SS - second)
        ' </StartTime>
        ' <Owner>
        '   <Number>Owner Phone Number</Number>
        '   <Name>Owner Name</Name>
        ' </Owner>
        ' <Subject>Schedule Reminder Subject</Subject>
        ' <EndTime>
        '   <Date>YYYY-MM-DD</Date>		(YYYY - year, MM - month, DD - day)
        '   <Time>HH:MM:SS</Time>		(HH - hour, MM - minute, SS - second)
        ' </EndTime>
        ' <Location>Location Name</Location>
        ' <Body>Schedule Reminder Body</Body>
        ' 


        ' <Category>Schedule Reminder</Category><DisplayType>Maximum</DisplayType><StartTime><Date>YYYY-MM-DD</Date><Time>HH:MM:SS</Time></StartTime><Owner><Number>Owner Phone Number</Number><Name>Owner Name</Name></Owner><Subject>Schedule Reminder Subject</Subject><EndTime><Date>YYYY-MM-DD</Date><Time>HH:MM:SS</Time></EndTime><Location>Location Name</Location><Body>Schedule Reminder Body</Body>

        'EndTime, Location and Body is displayed only in case DisplayType is set to "Maximum".
        '    Sample(Request)

        'Sample SOAP request to display SMS message (without HTTP headers):
        '==================================================================
        '<?xml version="1.0" encoding="utf-8"?>
        '  <s:Envelope s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
        '    <s:Body>
        '      <u:AddMessage xmlns:u="urn:samsung.com:service:MessageBoxService:1\">
        '        <MessageType>text/xml</MessageType>
        '        <MessageID>can be anything</MessageID>
        '        <Message>
        '          &lt;Category&gt;SMS&lt;/Category&gt;
        '          &lt;DisplayType&gt;Maximum&lt;/DisplayType&gt;
        '          &lt;ReceiveTime&gt;
        '          &lt;Date&gt;2010-05-04&lt;/Date&gt;
        '          &lt;Time&gt;01:02:03&lt;/Time&gt;
        '          &lt;/ReceiveTime&gt;
        '          &lt;Receiver&gt;
        '          &lt;Number&gt;12345678&lt;/Number&gt;
        '          &lt;Name&gt;Receiver&lt;/Name&gt;
        '          &lt;/Receiver&gt;
        '          &lt;Sender&gt;
        '          &lt;Number&gt;11111&lt;/Number&gt;
        '          &lt;Name&gt;Sender&lt;/Name&gt;
        '          &lt;/Sender&gt;
        '          &lt;Body&gt;Hello World!!!&lt;/Body&gt;
        '        </Message>
        '      </u:AddMessage>
        '    </s:Body>
        '  </s:Envelope>


        '<Category>SMS</Category>
        '<DisplayType>Maximum</DisplayType>
        '<ReceiveTime>
        '<Date>2010-05-04</Date>
        '<Time>01:02:03</Time>
        '</ReceiveTime>
        '<Receiver>
        '<Number>12345678</Number>
        '<Name>Receiver</Name>
        '</Receiver>
        '<Sender>
        '<Number>11111</Number>
        '<Name>Sender</Name>
        '</Sender>
        '<Body>Hello World!!!</Body>
        ' <Category>SMS</Category><DisplayType>Maximum</DisplayType><ReceiveTime><Date>2010-05-04</Date><Time>01:02:03</Time></ReceiveTime><Receiver><Number>12345678</Number><Name>Receiver</Name></Receiver><Sender><Number>11111</Number><Name>Sender</Name></Sender><Body>Hello Honey</Body>
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMessage called for UPnP Device = " & MyUPnPDeviceName & " with MessageId = " & MessageID.ToString & " and Message = " & Message.ToString, LogType.LOG_TYPE_INFO)
        Dim MessageXML As String = ""
        Dim time As DateTime = DateTime.Now
        Dim Dateformat As String = "yyyy-MM-dd"
        Dim Timeformat As String = "H:mm:ss"
        MessageXML = "<Category>SMS</Category><DisplayType>Maximum</DisplayType><ReceiveTime><Date>" & time.ToString(Dateformat) & "</Date><Time>" & time.ToString(Timeformat) & "</Time></ReceiveTime><Receiver><Number></Number><Name></Name></Receiver><Sender><Number></Number><Name>Homeseer Plugin</Name></Sender><Body>" & Message & "</Body>"
        PMBAddMessage("1", "text/xml; charset=""utf-8""", MessageXML)
    End Sub

    Public Sub SamsungSendKeyCode(KeyCodeString As String)

        Dim SamsungRemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        Dim ButtonInfos As String()
        Dim KeyCode As String = ""
        ButtonInfos = Split(KeyCodeString, ":;:-:")
        If UBound(ButtonInfos, 1) > 2 Then
            KeyCode = ButtonInfos(1)
        Else
            Exit Sub
        End If

        If SamsungRemoteType = "Samsungiapp" Then
            Dim MyAppString As String = "iphone..iapp.samsung"
            Dim MyTVAppString As String = "iphone..iapp.samsung"
            Dim MyRemoteName As String = "HomeSeer Samsung Remote"
            Dim MessagePart As String = ""
            Dim ReturnString As String = ""
            MessagePart = Chr(0) & Chr(0) & Chr(0) & Chr(ToBase64(KeyCode).Length) & Chr(0) & ToBase64(KeyCode)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendKeyCode was called for UPnPDevice = " & MyUPnPDeviceName & " with key = " & KeyCode.ToString & " and has created code = " & ToBase64(KeyCode), LogType.LOG_TYPE_INFO)
            'Dim FullPart As String = ""
            'FullPart = Chr(1) & Chr(MyTVAppString.Length) & Chr(0) & MyTVAppString & Chr(MessagePart.Length) & Chr(0) & MessagePart
            Dim FullPart As Byte() = System.Text.Encoding.ASCII.GetBytes(Chr(1) & Chr(MyTVAppString.Length) & Chr(0) & MyTVAppString & Chr(MessagePart.Length) & Chr(0) & MessagePart)
            If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                If MySamsungClient Is Nothing Then
                    Try
                        SamsungEstablishTCPConnection()
                    Catch ex As Exception
                        Log("Error in SamsungSendKeyCode for UPnPDevice = " & MyUPnPDeviceName & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Else
                Exit Sub
            End If
            If MySamsungClient Is Nothing Then
                Log("Error in SamsungSendKeyCode. There is no socket for UPnPDevice =  " & MyUPnPDeviceName, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            Try
                MySamsungAsyncSocket.response = False
                If Not MySamsungAsyncSocket.Send(FullPart) Then
                    Log("Error in SamsungSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code", LogType.LOG_TYPE_ERROR)
                    SamsungCloseTCPConnection(True)
                    FullPart = Nothing
                    Exit Sub
                End If
            Catch ex As Exception
                Log("Error in SamsungSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                SamsungCloseTCPConnection(True)
                FullPart = Nothing
                Exit Sub
            End Try
            FullPart = Nothing
            Try
                MySamsungAsyncSocket.sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in SamsungSendKeyCode for UPnPDevice =  " & MyUPnPDeviceName & " while sending a code with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                SamsungCloseTCPConnection(True)
                Exit Sub
            End Try

        ElseIf SamsungRemoteType = "SamsungWebSocket" Then
            ' this comes from https://forum.samygo.tv/viewtopic.php?f=104&t=12300 who claimes it works on Q type TVs
            ' "{"method":"ms.remote.control","params":{"Cmd":"Click","DataOfCmd":"%s","Option":"false","TypeOfRemote":"SendRemoteKey"}}", keyToSend);
            'Dim WidgetPart As String = "{""method"":""ms.channel.emit"",""params"":{""event"":""ed.apps.launch"",""to"":""host"",""data"":{""appId"":""" & KeyCode & """,""action_type"": ""NATIVE_LAUNCH""}}}"
            ' RN1MCdNq8t.Netflix
            ' gDhibXvFya.HBOGO
            ' Ahw07WXIjx.Dailymotion
            'tisT7SVUug.tunein
            'cexr1qp97S.Deezer
            'xqqJ00GGlC.okidoki
            '4ovn894vo9.Facebook
            'vbUQClczfR.Wuakitv
            'QizQxC7CUf.PlayMovies
            'QBA3qXl8rv.Kick
            'DJ8grEH6Hu.arte
            'JtPoChZbf4.Vimeo
            'hIWwRyZjcD.GameFlyStreaming
            'sHi2hDJGmf.nolim
            'guMmq95nKK.CanalPlusLauncher
            'RN1MCdNq8t.Netflix / org.tizen.netflix - app
            'evKhCgZelL.AmazonIgnitionLauncher2 / org.tizen.ignition
            '9Ur5IzDKqV.TizenYouTube
            'gDhibXvFya.HBOGO
            'EmCpcvhukH.ElevenSports
            'ASUvdWVqRb.FilmBoxLive
            'rJeHak5zRg.Spotify
            'ABor2M9vjb.acc(AccuWeather)
            'EkzyZtmneG.My5
            'yFo6bAK50v.Dennexpres
            'gdEZI5lLXr.Europa2FHD
            'bm9PqdAwjv.TvSme
            'dH3Ztod7bU.IDNES
            'wsFJCxteqc.OnetVodEden
            'rZyaXW5csM.TubaFM
            '4bjaTLNMia.curzon
            'RVvpJ8SIU6.ocs
            'bstjKvX6LM.molotov
            'RffagId0eC.SfrSport
            'phm0eEdRZ4.ExtraTweetIM2
            'VAarU8iUtx.samsungTizen(Vevo)
            'g0ScrkpO1l.SmartIPTV
            'kIciSQlYEM.plex

            Dim MessagePart As String = "{""method"":""ms.remote.control"",""params"":{""Cmd"":""Click"", ""Option"":""false"", ""TypeOfRemote"":""SendRemoteKey"", ""DataOfCmd"":""" & KeyCode & """}}"
            'Dim MessagePart As String = "{""eventType"":""EMP"",""plugin"":""SecondTV""}"
            'KeyCodeString = "{""eventType"":""EMP"",""plugin"":""RemoteControl""}"
            'KeyCodeString = "{""method"":""POST"",""body"":{""plugin"":""NNavi"",""api"":""GetDUID"",""version"":""1.000""}}"

            ' the CMD can be Click or Press or Release. The two latter used to enter a repeating key stroke.
            ' TypeOfRemote can be SendRemoteKey or ProcessMouseDevice or SendInputString

            Try

                If ButtonInfos(1) = "SamsungApp" Then
                    Dim SamsungAppID As String = ButtonInfos(4)
                    Dim SamsungAppType As String = ButtonInfos(6)
                    ' {"method":"ms.channel.emit","params":{"event": "ed.apps.launch", "to":"host", "data":{"appId": "org.tizen.browser", "action_type": "NATIVE_LAUNCH"}}
                    ' "app_type": 2 uses DEEP_LINK and I believe "app_type": 4 is NATIVE_LAUNCH
                    If SamsungAppType = "2" Then
                        ' use DEEP_LINK
                        MessagePart = "{""method"":""ms.channel.emit"",""params"":{""event"":""ed.apps.launch"",""to"":""host"",""data"":{""appId"":""" & SamsungAppID & """,""action_type"": ""DEEP_LINK""}}}"
                    ElseIf SamsungAppType = "4" Then
                        ' use NATIVE_LAUNCH
                        MessagePart = "{""method"":""ms.channel.emit"",""params"":{""event"":""ed.apps.launch"",""to"":""host"",""data"":{""appId"":""" & SamsungAppID & """,""action_type"": ""NATIVE_LAUNCH""}}}"
                    End If
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendKeyCode was called for UPnPDevice = " & MyUPnPDeviceName & " with AppID = " & SamsungAppID & " and has created code = " & MessagePart, LogType.LOG_TYPE_INFO)
                ElseIf ButtonInfos(1) = "SamsungWidget" Then
                    Dim SamsungWidget As String = ButtonInfos(2)
                    MessagePart = "{""method"":""ms.channel.emit"",""params"":{""event"":""ed.apps.launch"",""to"":""host"",""data"":{""appId"":""" & SamsungWidget & """,""action_type"": ""NATIVE_LAUNCH""}}}"
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendKeyCode was called for UPnPDevice = " & MyUPnPDeviceName & " with key = " & KeyCode.ToString & " with WidgetID = " & SamsungWidget & " and has created code = " & MessagePart, LogType.LOG_TYPE_INFO)
                ElseIf ButtonInfos(1) = "SamsungEnterText" Then
                    ' should I add sending text ie {"method":"ms.remote.control","params":{"Cmd":"$BASE64ENCODEDSTRING$","TypeOfRemote":"SendInputString","DataOfCmd":"base64"}}
                    Dim SamsungText As String = ButtonInfos(4)
                    MessagePart = "{""method"":""ms.remote.control"",""params"":{""Cmd"":""" & ToBase64(SamsungText) & """,""TypeOfRemote"":""SendInputString"",""DataOfCmd"":""base64""}}"
                Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungSendKeyCode was called for UPnPDevice = " & MyUPnPDeviceName & " with key = " & KeyCode.ToString & " and has created code = " & MessagePart, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
                Log("Error in SamsungSendKeyCode trying to form the message to be sent for UPnPDevice =  " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try

            If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                If MySamsungWebSocket Is Nothing Then
                    SamsungOpenUnEncryptedWebSocket(GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketPort.ToString, ""), GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungWebSocketLocation.ToString, ""))
                End If
            Else
                Exit Sub
            End If
            If MySamsungWebSocket Is Nothing Then
                Log("Error in SamsungSendKeyCode. There is no socket for UPnPDevice =  " & MyUPnPDeviceName, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes(MessagePart), True) Then
                Exit Sub
            End If
        ElseIf SamsungRemoteType = "SamsungWebSocketPIN" Then
            If GetBooleanIniFile("Remote Service by UDN", MyUDN, False) Then
                If MySamsungWebSocket Is Nothing Then
                    SamsungSessionID = SamsungAuthenticateUsePIN(GetStringIniFile(DeviceUDN, DeviceInfoIndex.diSamsungRemotePIN.ToString, ""), True)
                    If SamsungSessionID <> "" And SamsungAESKey IsNot Nothing Then
                        If SamsungAESKey.Length > 0 Then
                            SamsungOpenEncryptedWebSocket(SamsungSessionID)
                        Else
                            Log("Warning in SamsungSendKeyCode for Device = " & MyUPnPDeviceName & ". Autenticate and get PIN first. Go to the config page and click on Get PIN, enter PIN in test box", LogType.LOG_TYPE_WARNING)
                            Exit Sub
                        End If
                    Else
                        Log("Warning in SamsungSendKeyCode for Device = " & MyUPnPDeviceName & ". Autenticate and get PIN first. Go to the config page and click on Get PIN, enter PIN in test box", LogType.LOG_TYPE_WARNING)
                        Exit Sub
                    End If
                End If
            Else
                Exit Sub
            End If
            If MySamsungWebSocket Is Nothing Then
                Log("Error in SamsungSendKeyCode. There is no socket for UPnPDevice =  " & MyUPnPDeviceName, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            If Not SamsungSendEncryptedKeyCode(SamsungAESKey, SamsungSessionID, KeyCode) Then
                Exit Sub
            End If
        End If
    End Sub

    Private Sub SamsungTreatTextResponse(Input As String)
        ' Parse the response, lift out the ClientAckMsg
        Dim Args As String = ""
        Dim Name As String = ""

        Try
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1
            JSONdataLevel1 = json.DeserializeObject(Input)
            For Each Entry As Object In JSONdataLevel1
                If Entry.Key = "name" Then
                    Name = Entry.Value
                ElseIf Entry.Key = "args" Then
                    Args = Entry.Value
                End If
            Next
            JSONdataLevel1 = Nothing
            json = Nothing
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " processing response with error = " & ex.Message & " and Input = " & Input, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " found name = " & Name, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " found args = " & Args, LogType.LOG_TYPE_INFO)
        If Args = "" Then Exit Sub
        If Args.IndexOf("[") <> 0 Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " no [ found, Args = " & Args, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Else
            Args = Args.Remove(0, 1)
        End If

        If Args.IndexOf("]") <> Args.Length - 1 Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " no ] found, Args = " & Args, LogType.LOG_TYPE_ERROR)
            Exit Sub
        Else
            Args = Args.Remove(Args.Length - 1, 1)
        End If

        Dim ArgsSplit As String() = Args.Split(",")
        Dim ArgsBytes As Byte()
        ReDim ArgsBytes(UBound(ArgsSplit))
        For i = 0 To UBound(ArgsSplit)
            ArgsBytes(i) = Val(ArgsSplit(i))
        Next
        Dim ReturnVal As Byte() = AESD(ArgsBytes, SamsungAESKey)

        Dim ReturnString As String = System.Text.ASCIIEncoding.ASCII.GetChars(ReturnVal)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatTextResponse for device - " & MyUPnPDeviceName & " received decoded Args = " & ReturnString, LogType.LOG_TYPE_INFO)

    End Sub

    Public Sub HandleSamsungWebSocketDataReceived(sender As Object, e As Byte())
        If e Is Nothing Then Exit Sub

        If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived called for Device = " & MyUPnPDeviceName & " and Line = " & Encoding.UTF8.GetString(e, 0, e.Length), LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungWebSocketDataReceived called for Device = " & MyUPnPDeviceName & " Datasize = " & e.Length.ToString, LogType.LOG_TYPE_INFO)

        ' This is either text or binary, not sure it makes a difference

        If UBound(e) = 0 Then Exit Sub
        Dim SamsungRemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")

        Dim TextInfo As String = ASCIIEncoding.ASCII.GetChars(e)
        If SamsungRemoteType = "SamsungWebSocketPIN" Then
            Try
                If TextInfo = "1::" Then
                    Dim SocketData As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes("1::/com.samsung.companion")
                    If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " unsuccessful send '1::/com.samsung.companion'", LogType.LOG_TYPE_INFO)
                        Exit Sub
                    End If
                    Exit Sub
                    ' not sure about the purpose of the rest, but leave it in. 
                    Dim KeyCodeString As String = "{""eventType"":""EMP"",""plugin"":""SecondTV""}"
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is composing KeyCodeString = " & KeyCodeString, LogType.LOG_TYPE_INFO)

                    Dim EncryptedKeyCodeBytes As Byte() = AESE(System.Text.ASCIIEncoding.ASCII.GetBytes(KeyCodeString), SamsungAESKey)

                    Dim Body As String = "["
                    For i = 0 To EncryptedKeyCodeBytes.Length - 1
                        Body += EncryptedKeyCodeBytes(i).ToString + ","
                    Next

                    Body = Body.Remove(Body.Length - 1, 1) + "]"
                    SocketData = System.Text.ASCIIEncoding.ASCII.GetBytes("5::/com.samsung.companion:{""name"":""registerPush"",""args"":[{""Session_Id"":" & SamsungSessionID & ",""body"":""" & Body & """}]}")

                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is CommandInfo= " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)

                    If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " unsuccessful send Key = " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)
                    End If
                    wait(0.5)
                    KeyCodeString = "{""eventType"":""EMP"",""plugin"":""RemoteControl""}"
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is composing KeyCodeString = " & KeyCodeString, LogType.LOG_TYPE_INFO)

                    EncryptedKeyCodeBytes = AESE(System.Text.ASCIIEncoding.ASCII.GetBytes(KeyCodeString), SamsungAESKey)

                    Body = "["
                    For i = 0 To EncryptedKeyCodeBytes.Length - 1
                        Body += EncryptedKeyCodeBytes(i).ToString + ","
                    Next

                    Body = Body.Remove(Body.Length - 1, 1) + "]"
                    SocketData = System.Text.ASCIIEncoding.ASCII.GetBytes("5::/com.samsung.companion:{""name"":""registerPush"",""args"":[{""Session_Id"":" & SamsungSessionID & ",""body"":""" & Body & """}]}")
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is CommandInfo= " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)

                    If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " unsuccessful send Key = " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)
                    End If

                    wait(0.5)
                    KeyCodeString = "{""method"":""POST"",""body"":{""plugin"":""NNavi"",""api"":""GetDUID"",""version"":""1.000""}}"
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is composing KeyCodeString = " & KeyCodeString, LogType.LOG_TYPE_INFO)

                    EncryptedKeyCodeBytes = AESE(System.Text.ASCIIEncoding.ASCII.GetBytes(KeyCodeString), SamsungAESKey)

                    Body = "["
                    For i = 0 To EncryptedKeyCodeBytes.Length - 1
                        Body += EncryptedKeyCodeBytes(i).ToString + ","
                    Next

                    Body = Body.Remove(Body.Length - 1, 1) + "]"
                    SocketData = System.Text.ASCIIEncoding.ASCII.GetBytes("5::/com.samsung.companion:{""name"":""callCommon"",""args"":[{""Session_Id"":" & SamsungSessionID & ",""body"":""" & Body & """}]}")
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " is CommandInfo= " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)

                    If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, SocketData, True) Then
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungWebSocketDataReceived for device - " & MyUPnPDeviceName & " unsuccessful send Key = " & System.Text.ASCIIEncoding.ASCII.GetChars(SocketData), LogType.LOG_TYPE_INFO)
                    End If
                ElseIf TextInfo = "2::" Then
                    MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("2::"), True) 'False)
                ElseIf TextInfo.IndexOf("5::/com.samsung.companion:") <> -1 Then
                    TextInfo = TextInfo.Remove(0, 26)
                    SamsungTreatTextResponse(TextInfo)
                End If
            Catch ex As Exception
                Log("Error in HandleSamsungWebSocketDataReceived for UPnPDevice =  " & MyUPnPDeviceName & " while processing response with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            ' this is the newer models 
            SamsungTreatJSONResponse(TextInfo)
        End If
    End Sub

    Private Sub SamsungTreatJSONResponse(Input As String)
        ' too much logging ...
        If PIDebuglevel > DebugLevel.dlEvents Then
            Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " processing JSON with Input = " & Input, LogType.LOG_TYPE_INFO)
        ElseIf PIDebuglevel > DebugLevel.dlErrorsOnly Then
            If Input.Length > 100 Then
                Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " processing JSON with InputLength = " & Input.Length.ToString, LogType.LOG_TYPE_INFO)
            Else
                Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " processing JSON with Input = " & Input, LogType.LOG_TYPE_INFO)
            End If
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then
            Try
                Dim json As New JavaScriptSerializer
                Dim JSONdataLevel1 As Object
                JSONdataLevel1 = json.DeserializeObject(Input)
                For Each Entry As Object In JSONdataLevel1
                    Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " found Key = " & Entry.key.ToString, LogType.LOG_TYPE_INFO) ' too much logging & " and Value = " & Entry.value.ToString, LogType.LOG_TYPE_INFO)
                Next
                JSONdataLevel1 = Nothing
                json = Nothing
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " processing response with error = " & ex.Message & " and Input = " & Input, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        Try
            Dim EventInfo As Object = FindPairInJSONString(Input, "event")
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " found EventInfo = " & EventInfo.ToString, LogType.LOG_TYPE_INFO)
            If EventInfo.ToString.ToLower = "ms.channel.timeout" Then
                Log("warning SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received response = " & Input & " probably indicating no authentication. Retry while in front of TV and allow remote control access", LogType.LOG_TYPE_WARNING)
                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                SamsungCloseTCPConnection(True)
            ElseIf EventInfo.ToString.ToLower = "ms.channel.clientdisconnect" Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("warning SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received response = " & Input & " which will set remote to deactivate. New command will try to reopen or use activate", LogType.LOG_TYPE_WARNING)
                ' WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                SamsungCloseTCPConnection(True)
            ElseIf EventInfo.ToString.ToLower = "ms.channel.unauthorized" Then
                Log("warning SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received response = " & Input & " probably indicating previously DENIED authentication. Go to TV settings and deleted the denied device and retry, this time ALLOW authentication", LogType.LOG_TYPE_WARNING)
                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, False)
                SamsungCloseTCPConnection(True)

            ElseIf EventInfo.ToString.ToLower = "ms.error" Then
                Log("warning SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received response = " & Input & " probably indicating no authentication. Retry while in front of TV and allow remote control access", LogType.LOG_TYPE_WARNING)
            ElseIf EventInfo.ToString.ToLower = "ms.channel.connect" Then
                '
                ' {"data": {"clients"[{"attributes":{"name":"cmNjbGk="},"connectTime":1537329503141,"deviceName":"cmNjbGk=","id":"1c2591a4-694b-4223-8c2b-bee61687b82","isHost":false}],"id":"1c2591a4-694b-4223-8c2b-bee61687b82"},"event":"ms.channel.connect"}
                ' appears SSL encryption will need to use port 8002 and turn info will be something like this
                ' {"data": {"clients": [{"attributes": {"name": "fooBase64=="},"connectTime": 1541354167097,"deviceName": "fooBase64==","id": "xy123","isHost": false}],"id": "xy123","token": "65811577"},"event": "ms.channel.connect"}
                '   note the token is something to be used to keep authentication going ie.  wss://@ip:8002/api/v2/channels/samsung.remote.control?name=base64Name&token=THETOKEN
                ' {"data":{"clients":[{"attributes":{"name":"TWVkaWFDb250cm9sbGVy"},"connectTime":1543990640045,"deviceName":"TWVkaWFDb250cm9sbGVy","id":"1a95c2a7-c4d-47e2-829-22c4b313e52","isHost":false}],"id":"1a95c2a7-c4d-47e2-829-22c4b313e52"},"event":"ms.channel.connect"} 


                ' I seem to be getting all these multiple connect messages, how do I avoid doing the same thing over and over?

                WriteBooleanIniFile(MyUDN, DeviceInfoIndex.diRegistered.ToString, True)
                Try
                    Dim Data As Object = FindPairInJSONString(Input, "data")
                    ' appears no SSL encryption will need to use port 8002 and return info will be something like this
                    ' {"data": {"clients": [{"attributes": {"name": "fooBase64=="},"connectTime": 1541354167097,"deviceName": "fooBase64==","id": "xy123","isHost": false}],"id": "xy123","token": "65811577"},"event": "ms.channel.connect"}
                    If Data IsNot Nothing Then
                        Dim ID As String = FindPairInJSONString(Data, "id").ToString
                        If ID <> "" Then WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungClientID.ToString, ID)
                        Dim Token As String = FindPairInJSONString(Data, "token").ToString
                        If Token <> "" Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received and stored token = " & Token, LogType.LOG_TYPE_INFO)
                            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSamsungToken.ToString, Token)
                        End If
                        'Dim Clients As Object = FindPairInJSONString(Data, "clients")
                        'If Clients IsNot Nothing Then
                        'End If
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " no token found error was = " & ex.Message, LogType.LOG_TYPE_INFO)
                End Try

                ' I believe when EDEN_available is set in the isSupport info, we should do a "event":"ed.edenApp.get" else a "ed.installedApp.get"

                Dim isSupportInfo As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungisSupportInfo.ToString, "")
                Dim EdenSupport As Boolean = False
                Try
                    If isSupportInfo <> "" Then
                        Dim EdenAvail As String = FindPairInJSONString(isSupportInfo, "EDEN_available")
                        If EdenAvail <> "" Then
                            If EdenAvail.ToLower = "true" Then EdenSupport = True
                        End If
                    End If
                Catch ex As Exception
                End Try

                myRetrieveInitInfoTimer = New Timers.Timer With {
                    .AutoReset = False, .Interval = 1000, .Enabled = True}





            ElseIf EventInfo.ToString = "ed.installedApp.get" Then
                ' this is the list of installed apps in following structure
                ' {"data":{"data":[{"appId":"org.tizen.browser","app_type":4,"icon":"/opt/share/webappservice/apps_icon/FirstScreen/webbrowser/250x250.png","is_lock":0,"name":"Internet"}]},"event":"ed.installedApp.get","from":"host"}
                SamsungAddAppButtons(Input)
            ElseIf EventInfo.ToString = "ed.edenApp.get" Then
                ' { method 'ms.channel.emit', params: {Event:  'ed.edenApp.get',to 'host',data: { },}})
                ' {"data":{"data":[{"accelerators":[],"action_type":null,"appId":"com.samsung.tv.store","appType":"volt_app","icon":"/usr/apps/com.samsung.tv.csfs.res.tizen30/shared/res/Resource/apps/apps/sysAppsNromal.png","id":"APPS","isLock":false,"launcherType":"system","mbrIndex":null,"mbrSource":null,"name":"APPS","position":0,"sourceTypeNum":null},{"accelerators":[],"action_type":null,"appId":"org.tizen.browser","appType":"web_app","icon":"/opt/share/webappservice/apps_icon/FirstScreen/webbrowser/245x138.png","id":"org.tizen.browser","isLock":false,"launcherType":"launcher","mbrIndex":null,"mbrSource":null,"name":"Internet","position":1,"sourceTypeNum":null}]},"event":"ed.edenApp.get","from":"host"} 
                SamsungAddAppButtons(Input)
            ElseIf EventInfo.ToString = "ed.apps.launch" Then
                ' {"method":"ms.channel.emit","params":{"event": "ed.apps.launch", "to":"host", "data":{"appId":"org.tizen.browser","action_type":"NATIVE_LAUNCH","metaTag":"http:\/\/hackaday.com"}}}
            ElseIf EventInfo.ToString = "ms.remote.imeStart" Then
                '{"data":"input","entrylimit":-1,"event":"ms.remote.imeStart"} 
            ElseIf EventInfo.ToString = "ms.remote.imeUpdate" Then
                ' {"data":"aHR0cDovL3d3dy5lc3BuLmNvbS9uYmEvc3RvcnkvXy9pZC8yNTUzMDk2Ni9kaXJrLW5vd2l0emtpLWRhbGxhcy1tYXZlcmlja3MtbWFrZXMtZGVidXQtcmVjb3JkLTIxc3Qtc2Vhc29u","entrylimit":-1,"event":"ms.remote.imeUpdate"} 
                Try
                    Dim Data As Object = FindPairInJSONString(Input, "data")
                    If Data IsNot Nothing Then
                        Dim Data64Base As Byte() = Convert.FromBase64String(CType(Data, String))
                        If Data64Base IsNot Nothing Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received and imeData = " & System.Text.Encoding.Unicode.GetString(Data64Base), LogType.LOG_TYPE_INFO)
                        End If
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " no token found error was = " & ex.Message, LogType.LOG_TYPE_INFO)
                End Try
            ElseIf EventInfo.ToString = "ms.remote.imeEnd" Then
                '{"event":"ms.remote.imeEnd"} 
            ElseIf EventInfo.ToString = "ed.edenTV.update" Then
                '{"event":"ed.edenTV.update","data":{"update_type":"ed.edenApp.update"}}
                Try
                    Dim Data As Object = FindPairInJSONString(Input, "data")
                    If Data IsNot Nothing Then
                        Dim UpdateType As String = CType(FindPairInJSONString(Data, "update_type"), String)
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " received and edenTV update with update_type = " & UpdateType.ToString, LogType.LOG_TYPE_INFO)
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " no update_type found with error = " & ex.Message, LogType.LOG_TYPE_INFO)
                End Try
            ElseIf EventInfo.ToString = "ms.channel.read" Then
                'ms.channel.read"
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SamsungTreatJSONResponse for device - " & MyUPnPDeviceName & " couldn't find event info with error = " & ex.Message & " and Input = " & Input, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub myRetrieveInitInfoTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles myRetrieveInitInfoTimer.Elapsed
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("myRetrieveInitInfoTimer_Elapsed called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Try
            Dim isSupportInfo As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diSamsungisSupportInfo.ToString, "")
            Dim EdenSupport As Boolean = False
            Try
                If isSupportInfo <> "" Then
                    Dim EdenAvail As String = FindPairInJSONString(isSupportInfo, "EDEN_available")
                    If EdenAvail <> "" Then
                        If EdenAvail.ToLower = "true" Then EdenSupport = True
                    End If
                End If
            Catch ex As Exception
            End Try
            ' https://review.tizen.org/git/?p=platform/core/convergence/app-comm-svc.git;a=blob;f=MSF-Node/org.tizen.multiscreen/server/plugins/plugin-api-v2/channels/index.js;h=8d548dcdda4e9fd12a9d280e7ee4cc3199e8e967;hb=refs/heads/tizen_3.0#l374

            If EdenSupport Then
                If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.edenApp.get"", ""to"":""host""}}"), True) Then
                    '{"data":{"data":[{"accelerators":[],"action_type":null, "appId":  "com.samsung.tv.store","appType":"volt_app","icon":"/usr/apps/com.samsung.tv.csfs.res.tizen30/shared/res/Resource/apps/apps/sysAppsNromal.png","id":"APPS","isLock":false,"launcherType":"system","mbrIndex":null, "mbrSource": null, "name":  "APPS","position":0,"sourceTypeNum":null}, {"accelerators":  [],"action_type":null, "appId":  "org.tizen.browser","appType":"web_app","icon":"/opt/share/webappservice/apps_icon/FirstScreen/webbrowser/245x138.png","id":"org.tizen.browser","isLock":false,"launcherType":"launcher","mbrIndex":null, "mbrSource": null, "name":  "Internet","position":1,"sourceTypeNum":null}]}, "event":  "ed.edenApp.get","from":"host"} 
                End If
            Else
                If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.installedApp.get"", ""to"":""host""}}"), True) Then
                End If
            End If

            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.getChannel.get"", ""to"":""host""}}"), True) Then
            End If

            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""say"", ""data"":""Hello World!"",""to"":""all""}}"), True) Then
            End If

            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.edenTV.update"",""to"":""host""}}"), True) Then
            End If

            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.edenTV.get"",""to"":""host""}}"), True) Then
            End If

            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.emit"",""params"":{""event"": ""ed.edenApp.update"",""to"":""host""}}"), True) Then
            End If

            ' {"event":"ed.edenTV.update","data":{"update_type":"ed.edenApp.update"}}
            ' 
            If Not MySamsungWebSocket.SendDataOverWebSocket(OpcodeText, ASCIIEncoding.ASCII.GetBytes("{""method"":""ms.channel.read"",""params"":{""event"": ""ed.appStateRequest.get"", ""to"":""host""}}"), True) Then
            End If
            ' ms.channel.anything
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlOff Then Log("Error in myRetrieveInitInfoTimer_Elapsed with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub


    Public Sub HandleSamsungTCPDataReceived(sender As Object, e As Byte())
        If e Is Nothing Then Exit Sub
        If PIDebuglevel > DebugLevel.dlEvents Then Log("HandleSamsungTCPDataReceived called for Device = " & MyUPnPDeviceName & " and Line = " & Encoding.UTF8.GetString(e, 0, e.Length), LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungTCPDataReceived called for Device = " & MyUPnPDeviceName & " Datasize = " & e.Length.ToString, LogType.LOG_TYPE_INFO)

    End Sub

    Public Sub HandleSamsungSocketClosed(sender As Object)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleSamsungSocketClosed called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        MyRemoteServiceActive = False
        If MySamsungWebSocket IsNot Nothing Then
            RemoveHandler MySamsungWebSocket.DataReceived, AddressOf HandleSamsungWebSocketDataReceived
            RemoveHandler MySamsungWebSocket.WebSocketClosed, AddressOf HandleSamsungSocketClosed
        End If
        MySamsungWebSocket = Nothing
        ' maybe some more actions are needed here, like updating the HS status of the remote
        Try
            If HSRefRemote <> -1 Then hs.SetDeviceValueByRef(HSRefRemote, dsDeactivated, True)
            MyRemoteServiceActive = False
        Catch ex As Exception
            Log("Error in HandleSamsungSocketClosed for UPnPDevice = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Function AESE(ByVal input As Byte(), ByVal key As Byte()) As Byte()
        Try
            Dim AES As New Security.Cryptography.RijndaelManaged
            AES.Padding = Security.Cryptography.PaddingMode.PKCS7
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESE for Device = " & MyUPnPDeviceName & " has Key Length =" + key.Length.ToString, LogType.LOG_TYPE_INFO)
            AES.Key = key
            AES.Mode = Security.Cryptography.CipherMode.ECB
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESE for Device = " & MyUPnPDeviceName & " has Mode       =" + AES.Mode.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESE for Device = " & MyUPnPDeviceName & " has Padding is =" + AES.Padding.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESE for Device = " & MyUPnPDeviceName & " has Key Size   =" + AES.KeySize.ToString, LogType.LOG_TYPE_INFO)
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = input
            Return DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AESE for UPnPDevice =  " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Private Function AESD(ByVal input As Byte(), ByVal key As Byte()) As Byte()
        Try
            Dim AES As New System.Security.Cryptography.RijndaelManaged
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESD for Device = " & MyUPnPDeviceName & " has Input lenght =" + input.Length.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESD for Device = " & MyUPnPDeviceName & " has Key lenght =" + key.Length.ToString, LogType.LOG_TYPE_INFO)
            AES.Padding = System.Security.Cryptography.PaddingMode.None
            AES.Mode = Security.Cryptography.CipherMode.ECB
            AES.Key = key
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESD for Device = " & MyUPnPDeviceName & " has Mode       =" + AES.Mode.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESD for Device = " & MyUPnPDeviceName & " has Padding is =" + AES.Padding.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AESD for Device = " & MyUPnPDeviceName & " has Key Size   =" + AES.KeySize.ToString, LogType.LOG_TYPE_INFO)
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = input
            Return DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AESD for UPnPDevice =  " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Private Function GenerateServerHello(UserID As String, Pin As String, ByRef aes_key As Byte(), ByRef data_hash As Byte()) As String
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenerateServerHello called for UPnPDevice =  " & MyUPnPDeviceName & " with UserID = " & UserID & " and PIN = " & Pin, LogType.LOG_TYPE_INFO)

            ' sha1 = hashlib.sha1()
            Dim sha1 As New System.Security.Cryptography.SHA1Managed

            ' sha1.update(pin.encode('utf-8'))
            ' pin_hash = sha1.digest()
            Dim pin_hash As Byte() = sha1.ComputeHash(Encoding.UTF8.GetBytes(Pin)) ' this returns 20 bytes
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated pin_hash = " & ByteArrayToHexString(pin_hash), LogType.LOG_TYPE_INFO)

            'aes_key = pin_hash[:16] this takes the first 16 characters
            'Dim aes_key As Byte() = Nothing
            ReDim aes_key(15)
            Array.Copy(pin_hash, 0, aes_key, 0, 16)

            'logger.debug('crypto: aes: ', aes_key)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated aes_key = " & ByteArrayToHexString(aes_key), LogType.LOG_TYPE_INFO)

            'iv = b"\x00" * BLOCK_SIZE
            Dim iv As Byte() = New Byte(BLOCK_SIZE - 1) {}
            For i = 0 To BLOCK_SIZE - 1
                iv(i) = 0
            Next

            'cipher = AES.new(aes_key, AES.MODE_CBC, iv)
            Dim cipher As New System.Security.Cryptography.RijndaelManaged
            cipher.Padding = System.Security.Cryptography.PaddingMode.None
            cipher.Mode = Security.Cryptography.CipherMode.CBC
            cipher.Key = aes_key
            cipher.IV = iv

            'encrypted = cipher.encrypt(bytes(bytearray.fromhex(keys.publicKey)))
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = cipher.CreateEncryptor
            Dim buffer As Byte() = HexStringToByteArray(publicKey)
            Dim encryped As Byte() = DESEncrypter.TransformFinalBlock(buffer, 0, buffer.Length)

            'logger.debug('crypto: aes encrypted: ', encrypted.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated encrypted = " & ByteArrayToHexString(encryped), LogType.LOG_TYPE_INFO)

            'swapped = encrypt_parameter_data_with_aes(encrypted)
            Dim swapped As Byte() = encrypt_parameter_data_with_aes(encryped)
            If swapped Is Nothing Then Return Nothing

            'logger.debug('crypto: aes swapped: ', swapped)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated swapped = " & ByteArrayToHexString(swapped), LogType.LOG_TYPE_INFO)

            'Data = struct.pack(">I", Len(user_id)) + user_id.encode('utf-8') + swapped   pack big endian unsigned integers = 4 bytes
            Dim Data As Byte() = BitConverter.GetBytes(CType(UserID.Length, UInt32))
            If (BitConverter.IsLittleEndian) Then
                Array.Reverse(Data)
            End If
            Data = Data.Concat(ASCIIEncoding.UTF8.GetBytes(UserID)).ToArray()
            Data = Data.Concat(swapped).ToArray()

            'logger.debug('crypto: data buffer: ', data)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated data buffer = " & ByteArrayToHexString(Data), LogType.LOG_TYPE_INFO)

            'sha1 = hashlib.sha1()
            'sha1.update(Data)
            'data_hash = sha1.digest()
            data_hash = sha1.ComputeHash(Data)

            'logger.debug('crypto: data hash: ', data_hash)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated data_hash  = " & ByteArrayToHexString(data_hash), LogType.LOG_TYPE_INFO)

            'server_hello = (
            'b"\x01\x02" +
            '(b"\x00" * 5) +
            'struct.pack(">I", Len(user_id) + 132) +
            'Data +
            '(b"\x00" * 5)
            ')
            Dim server_hello As Byte() = {1, 2, 0, 0, 0, 0, 0}
            Dim UIdLen As Byte() = BitConverter.GetBytes(CType(UserID.Length + 132, UInt32))
            If (BitConverter.IsLittleEndian) Then
                Array.Reverse(UIdLen)
            End If
            server_hello = server_hello.Concat(UIdLen).ToArray()
            server_hello = server_hello.Concat(Data).ToArray()
            server_hello = server_hello.Concat({0, 0, 0, 0, 0}).ToArray()

            'Return server_hello, data_hash, aes_key
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " generated ServerHello = " & ByteArrayToHexString(server_hello), LogType.LOG_TYPE_INFO)

            Return ByteArrayToHexString(server_hello)

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GenerateServerHello for UPnPDevice =  " & MyUPnPDeviceName & " with UserID = " & UserID & ", Pin = " & Pin & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Function ParseClientHello(clientHello As String, dataHash As Byte(), aesKey As Byte(), gUserId As String, ByRef SKPrime As Byte()) As Byte()
        Try
            Const USER_ID_POS As Integer = 15
            Const USER_ID_LEN_POS As Integer = 11
            Const GX_SIZE As Byte = 128 ' 0x80

            '   GX_SIZE = 0x80
            '    data = bytes(bytearray.fromhex(client_hello))
            Dim data As Byte() = HexStringToByteArray(clientHello)

            '    firstLen = struct.unpack(">I", data[7:11])[0] ' this get a 4 byte integer and only uses the first byte of it, it is big Endean
            Dim TempArray As Byte()

            ReDim TempArray(3)
            Array.Copy(data, 7, TempArray, 0, 4)
            If (BitConverter.IsLittleEndian) Then
                Array.Reverse(TempArray)
            End If
            Dim firstLen As Integer = BitConverter.ToInt32(TempArray, 0)

            '    userIdLen = struct.unpack(">I", data[11:15])[0]
            ReDim TempArray(3)
            Array.Copy(data, 11, TempArray, 0, 4)
            If (BitConverter.IsLittleEndian) Then
                Array.Reverse(TempArray)
            End If
            Dim userIdLen As Integer = BitConverter.ToInt32(TempArray, 0)

            '    destLen = user_id_len + 132 + SHA_DIGEST_LENGTH
            Dim destLen As Integer = userIdLen + 132 + SHA_DIGEST_LENGTH

            '    thirdLen = user_id_len + 132
            Dim thirdLen As Integer = userIdLen + 132

            '    print("thirdLen: "+str(thirdLen))
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " thirdLen: " & thirdLen, LogType.LOG_TYPE_INFO)

            '    Print("hello: " + data.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " hello: " & ByteArrayToHexString(data), LogType.LOG_TYPE_INFO)

            '    dest = data[USER_ID_LEN_POS:thirdLen+USER_ID_LEN_POS] + dataHash
            Dim dest As Byte()
            ReDim dest(thirdLen - 1)
            Array.Copy(data, USER_ID_LEN_POS, dest, 0, thirdLen)
            dest = dest.Concat(dataHash).ToArray()

            '    print("dest: "+dest.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " dest: " & ByteArrayToHexString(dest), LogType.LOG_TYPE_INFO)

            '    userId=data[USER_ID_POS:userIdLen+USER_ID_POS]
            Dim userId As Byte()
            ReDim userId(userIdLen - 1)
            Array.Copy(data, USER_ID_POS, userId, 0, userIdLen)

            '    Print("userId: " + userId.decode('utf-8'))
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " userId: " & ASCIIEncoding.UTF8.GetChars(userId), LogType.LOG_TYPE_INFO)


            '    pEncWBGx = data[USER_ID_POS+userIdLen:GX_SIZE+USER_ID_POS+userIdLen]
            Dim pEncWBGx As Byte()
            ReDim pEncWBGx(GX_SIZE - 1)
            Array.Copy(data, USER_ID_POS + userIdLen, pEncWBGx, 0, GX_SIZE)

            '    Print("pEncWBGx: " + pEncWBGx.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " pEncWBGx: " & ByteArrayToHexString(pEncWBGx), LogType.LOG_TYPE_INFO)

            '    pEncGx = decrypt_parameter_data_with_aes(pEncWBGx)
            Dim pEncGx As Byte() = decrypt_parameter_data_with_aes(pEncWBGx)

            '    Print("pEncGx: " + pEncGx.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " pEncGx: " & ByteArrayToHexString(pEncGx), LogType.LOG_TYPE_INFO)

            '    iv = b"\x00" * BLOCK_SIZE
            Dim iv As Byte() = New Byte(BLOCK_SIZE - 1) {}
            For i = 0 To BLOCK_SIZE - 1
                iv(i) = 0
            Next

            '    cipher = AES.new(aesKey, AES.MODE_CBC, iv)
            Dim cipher As New System.Security.Cryptography.RijndaelManaged
            cipher.Padding = System.Security.Cryptography.PaddingMode.None
            cipher.Mode = Security.Cryptography.CipherMode.CBC
            cipher.Key = aesKey
            cipher.IV = iv

            '    pGx = cipher.decrypt(pEncGx)
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = cipher.CreateDecryptor
            Dim pGx As Byte() = DESDecrypter.TransformFinalBlock(pEncGx, 0, pEncGx.Length)

            '    print("pGx: " + pGx.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " pGx: " & ByteArrayToHexString(pGx), LogType.LOG_TYPE_INFO)

            '    bnPGx = int(pGx.hex(),16)
            Dim TemppGX As Byte() = {}
            ReDim TemppGX(pGx.Length - 1)
            pGx.CopyTo(TemppGX, 0)
            If BitConverter.IsLittleEndian Then
                Array.Reverse(TemppGX)
            End If
            TemppGX = TemppGX.Concat({0}).ToArray()   ' take out any negative sign
            Dim bnPGx As BigInteger = New BigInteger(TemppGX)

            '    bnPrime = int(keys.prime,16)
            Dim bmprimeArray As Byte() = HexStringToByteArray(prime)
            If BitConverter.IsLittleEndian Then
                Array.Reverse(bmprimeArray)
            End If
            bmprimeArray = bmprimeArray.Concat({0}).ToArray()   ' take out any negative sign
            Dim bnPrime As BigInteger = New BigInteger(bmprimeArray)

            '    bnPrivateKey = int(keys.privateKey,16)
            Dim bnPrivateKeyArray As Byte() = HexStringToByteArray(privateKey)
            If BitConverter.IsLittleEndian Then
                Array.Reverse(bnPrivateKeyArray)
            End If
            bnPrivateKeyArray = bnPrivateKeyArray.Concat({0}).ToArray()   ' take out any negative sign
            Dim bnPrivateKey As BigInteger = New BigInteger(bnPrivateKeyArray)

            '    secret = bytes.fromhex(hex(pow(bnPGx, bnPrivateKey, bnPrime)).rstrip("L").lstrip("0x"))
            Dim secret As Byte() = BigInteger.ModPow(bnPGx, bnPrivateKey, bnPrime).ToByteArray
            If (secret.Length = 129) Then
                If (secret(128) = 0) Then
                    ReDim Preserve secret(127)
                End If
            End If

            If BitConverter.IsLittleEndian Then
                Array.Reverse(secret)
            End If

            '    print("secret: " + secret.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " secret: " & ByteArrayToHexString(secret), LogType.LOG_TYPE_INFO)

            '    dataHash2 = data[USER_ID_POS+userIdLen+GX_SIZE:USER_ID_POS+userIdLen+GX_SIZE+SHA_DIGEST_LENGTH];
            Dim dataHash2 As Byte()
            ReDim dataHash2(SHA_DIGEST_LENGTH - 1)
            Array.Copy(data, USER_ID_POS + userIdLen + GX_SIZE, dataHash2, 0, SHA_DIGEST_LENGTH)

            '    print("hash2: " + dataHash2.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " hash2: " & ByteArrayToHexString(dataHash2), LogType.LOG_TYPE_INFO)

            '    secret2 = userId + secret
            Dim secret2 As Byte() = {}
            secret2 = secret2.Concat(userId).ToArray()
            secret2 = secret2.Concat(secret).ToArray()

            '    print("secret2: " + secret2.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " secret2: " & ByteArrayToHexString(secret2), LogType.LOG_TYPE_INFO)

            '    sha1 = hashlib.sha1()
            Dim sha1 As New System.Security.Cryptography.SHA1Managed

            '    sha1.update(secret2)
            '    dataHash3 = sha1.digest()
            Dim dataHash3 As Byte() = sha1.ComputeHash(secret2)

            '    print("hash3: " + dataHash3.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " hash3: " & ByteArrayToHexString(dataHash3), LogType.LOG_TYPE_INFO)

            '    if dataHash2 != dataHash3:
            '        print("Pin error!!!")
            '        return False
            Dim different = False
            If dataHash2.Length = dataHash3.Length Then
                For i = 0 To dataHash2.Length - 1
                    If dataHash2(i) <> dataHash3(i) Then
                        different = True
                        Exit For
                    End If
                Next
            Else
                different = True
            End If
            If different Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("warning in ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " Pin error!!! ", LogType.LOG_TYPE_WARNING)
                Return Nothing
            End If

            '    print("Pin OK :)\n")
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " Pin OK!!! ", LogType.LOG_TYPE_INFO)

            '    flagPos = userIdLen + USER_ID_POS + GX_SIZE + SHA_DIGEST_LENGTH
            Dim flagPos As Integer = userIdLen + USER_ID_POS + GX_SIZE + SHA_DIGEST_LENGTH

            '    if ord(data[flagPos:flagPos+1]):
            '        print("First flag error!!!")
            '        return False
            If data(flagPos) <> 0 Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("warning in ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " First flag error!!! ", LogType.LOG_TYPE_WARNING)
                Return Nothing
            End If

            '    flagPos = userIdLen + USER_ID_POS + GX_SIZE + SHA_DIGEST_LENGTH
            flagPos = userIdLen + USER_ID_POS + GX_SIZE + SHA_DIGEST_LENGTH

            '    if struct.unpack(">I",data[flagPos+1:flagPos+5])[0]:
            '        print("Second flag error!!!")
            '        return False
            ReDim TempArray(4 - 1)
            Array.Copy(data, flagPos + 1, TempArray, 0, 4)
            If (BitConverter.IsLittleEndian) Then
                Array.Reverse(TempArray)
            End If
            Dim TempInt As Integer = BitConverter.ToInt32(TempArray, 0)
            If TempInt <> 0 Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("warning in ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " Second flag error!!! ", LogType.LOG_TYPE_WARNING)
                Return Nothing
            End If

            '    sha1 = hashlib.sha1()
            '    sha1.update(dest)
            '    dest_hash = sha1.digest()
            Dim dest_hash As Byte() = sha1.ComputeHash(dest)

            '    print("dest_hash: " + dest_hash.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " dest_hash: " & ByteArrayToHexString(dest_hash), LogType.LOG_TYPE_INFO)

            '    finalBuffer = userId + gUserId.encode('utf-8') + pGx + bytes.fromhex(keys.publicKey) + secret
            Dim finalBuffer As Byte() = {}
            finalBuffer = finalBuffer.Concat(userId).ToArray()
            'finalBuffer = finalBuffer.Concat({6, 5, 4, 3, 2, 1}).ToArray()
            finalBuffer = finalBuffer.Concat(Encoding.UTF8.GetBytes(gUserId)).ToArray()
            finalBuffer = finalBuffer.Concat(pGx).ToArray()
            finalBuffer = finalBuffer.Concat(HexStringToByteArray(publicKey)).ToArray()
            finalBuffer = finalBuffer.Concat(secret).ToArray()

            '    sha1 = hashlib.sha1()
            '    sha1.update(finalBuffer)
            '    SKPrime = sha1.digest()
            SKPrime = sha1.ComputeHash(finalBuffer)

            '    print("SKPrime: " + SKPrime.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " SKPrime: " & ByteArrayToHexString(SKPrime), LogType.LOG_TYPE_INFO)

            '    sha1 = hashlib.sha1()
            '    sha1.update(SKPrime+b"\x00")
            '    SKPrimeHash = sha1.digest()
            Dim SKPrimeHash As Byte() = sha1.ComputeHash(SKPrime.Concat({0}).ToArray)

            '    print("SKPrimeHash: " + SKPrimeHash.hex())
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " SKPrimeHash: " & ByteArrayToHexString(SKPrimeHash), LogType.LOG_TYPE_INFO)

            '    ctx = applySamyGOKeyTransform(SKPrimeHash[:16])

            Dim ctx As Byte()
            ReDim ctx(15)
            Dim PartialHash As Byte()
            ReDim PartialHash(15)
            Array.Copy(SKPrimeHash, 0, PartialHash, 0, 16)
            ' ctx = applySamyGOKeyTransform(PartialHash)

            Dim ctxPointer As IntPtr = Marshal.AllocHGlobal(ctx.Length)
            Dim PartialHashPointer As IntPtr = Marshal.AllocHGlobal(PartialHash.Length)

            Marshal.Copy(ctx, 0, ctxPointer, ctx.Length)
            Marshal.Copy(PartialHash, 0, PartialHashPointer, PartialHash.Length)

            applySamyGOKeyTransform(PartialHashPointer, ctxPointer)

            Marshal.Copy(ctxPointer, ctx, 0, ctx.Length)

            Marshal.FreeHGlobal(ctxPointer)
            Marshal.FreeHGlobal(PartialHashPointer)



            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " ctx: " & ByteArrayToHexString(ctx), LogType.LOG_TYPE_INFO)

            '    return {"ctx": ctx, "SKPrime": SKPrime}

            Return ctx

        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ParseClientHello for UPnPDevice =  " & MyUPnPDeviceName & " with clientHello = " & clientHello & ", dataHash = " & ByteArrayToHexString(dataHash) & ", aesKey = " & ByteArrayToHexString(aesKey) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Private Function generateServerAcknowledge(SKPrime As Byte()) As String
        Try
            ' sha1 = hashlib.sha1()
            Dim sha1 As New System.Security.Cryptography.SHA1Managed

            'sha1.update(SKPrime + b"\x01")
            'SKPrimeHash = sha1.digest()
            Dim SKPrimeHash As Byte() = sha1.ComputeHash(SKPrime.Concat({1}).ToArray)
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("generateServerAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " generated SKPrimeHash = " & ByteArrayToHexString(SKPrimeHash), LogType.LOG_TYPE_INFO)

            'Return "0103000000000000000014" + SKPrimeHash.hex().upper() + "0000000000"
            Dim ReturnString As String = "0103000000000000000014" + ByteArrayToHexString(SKPrimeHash).ToUpper + "0000000000"
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("generateServerAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " generated returnString = " & ReturnString, LogType.LOG_TYPE_INFO)
            Return ReturnString
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in generateServerAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " with SKPrime = " & ByteArrayToHexString(SKPrime) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return ""
        End Try
    End Function

    Private Function parseClientAcknowledge(clientAck As String, SKPrime As Byte()) As Boolean
        Try
            '   sha1 = hashlib.sha1()
            Dim sha1 As New System.Security.Cryptography.SHA1Managed

            '   sha1.update(SKPrime + b"\x02")
            '   SKPrimeHash = sha1.digest()
            Dim SKPrimeHash As Byte() = sha1.ComputeHash(SKPrime.Concat({2}).ToArray)

            '   tmpClientAck = "0104000000000000000014" + SKPrimeHash.hex().upper() + "0000000000"
            Dim tmpClientAck As String = "0104000000000000000014" + ByteArrayToHexString(SKPrimeHash).ToUpper() + "0000000000"
            If clientAck = tmpClientAck Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("parseClientAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " returned true!!", LogType.LOG_TYPE_INFO)
                Return True
            Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in parseClientAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " returned false!!", LogType.LOG_TYPE_WARNING)
                Return False
            End If
            '   Return clientAck == tmpClientAck
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in parseClientAcknowledge for UPnPDevice =  " & MyUPnPDeviceName & " with clientAck = " & clientAck & " and SKPrime = " & ByteArrayToHexString(SKPrime) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
    End Function




    Private Function encrypt_parameter_data_with_aes(input As Byte()) As Byte()
        Try
            If PIDebuglevel > DebugLevel.dlEvents Then Log("encrypt_parameter_data_with_aes called for UPnPDevice =  " & MyUPnPDeviceName & " with input = " & ByteArrayToHexString(input), LogType.LOG_TYPE_INFO)
            'iv = b"\x00" * BLOCK_SIZE
            Dim iv As Byte() = New Byte(BLOCK_SIZE - 1) {}
            For i = 0 To BLOCK_SIZE - 1
                iv(i) = 0
            Next
            'output = b""
            Dim output As Byte() = {}
            'For num in range(0, 128, 16):
            '   cipher = AES.new(
            '       bytes(bytearray.fromhex(keys.wbKey)),AES.MODE_CBC,iv )
            '   output += cipher.encrypt(inpt[num:num+16])
            Dim buffer As Byte()
            ReDim buffer(15)
            Dim Index_ As Integer = 0
            For Index As Integer = 0 To 127 Step 16
                If (Index + 16) <= input.Length Then
                    Dim cipher As New System.Security.Cryptography.RijndaelManaged
                    cipher.Padding = System.Security.Cryptography.PaddingMode.None
                    cipher.Mode = Security.Cryptography.CipherMode.CBC
                    cipher.Key = HexStringToByteArray(wbKey)
                    cipher.IV = iv
                    Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = cipher.CreateEncryptor
                    System.Array.Copy(input, Index, buffer, 0, 16)
                    Dim encryped As Byte() = DESEncrypter.TransformFinalBlock(buffer, 0, buffer.Length)
                    output = output.Concat(encryped).ToArray()
                End If
            Next
            Return output
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in encrypt_parameter_data_with_aes for UPnPDevice =  " & MyUPnPDeviceName & " with input = " & ByteArrayToHexString(input) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Private Function decrypt_parameter_data_with_aes(Input As Byte()) As Byte()
        Try
            If PIDebuglevel > DebugLevel.dlEvents Then Log("encrypt_parameter_data_with_aes called for UPnPDevice =  " & MyUPnPDeviceName & " with input = " & ByteArrayToHexString(Input), LogType.LOG_TYPE_INFO)
            'iv = b"\x00" * BLOCK_SIZE
            Dim iv As Byte() = New Byte(BLOCK_SIZE - 1) {}
            For i = 0 To BLOCK_SIZE - 1
                iv(i) = 0
            Next
            'output = b""
            Dim output As Byte() = {}

            'For num in range(0, 128, 16)
            '   cipher = AES.new(bytes(bytearray.fromhex(keys.wbKey)),AES.MODE_CBC,iv)
            '   output += cipher.decrypt(inpt[num: num+16])
            Dim buffer As Byte()
            ReDim buffer(15)
            Dim Index_ As Integer = 0
            For Index As Integer = 0 To 127 Step 16
                If (Index + 16) <= Input.Length Then
                    Dim cipher As New System.Security.Cryptography.RijndaelManaged
                    cipher.Padding = System.Security.Cryptography.PaddingMode.None
                    cipher.Mode = Security.Cryptography.CipherMode.CBC
                    cipher.Key = HexStringToByteArray(wbKey)
                    cipher.IV = iv
                    Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = cipher.CreateDecryptor
                    System.Array.Copy(Input, Index, buffer, 0, 16)
                    Dim decrypt As Byte() = DESDecrypter.TransformFinalBlock(buffer, 0, buffer.Length)
                    output = output.Concat(decrypt).ToArray()
                End If
            Next
            Return output
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in decrypt_parameter_data_with_aes for UPnPDevice =  " & MyUPnPDeviceName & " with input = " & ByteArrayToHexString(Input) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Private Function applySamyGOKeyTransform__(Input As Byte()) As Byte()
        Dim r As System.Security.Cryptography.Aes = System.Security.Cryptography.Aes.Create()

        r.BlockSize = 128
        r.Key = HexStringToByteArray(transKey)
        Dim transKeyArray As Byte() = HexStringToByteArray(transKey)
        'r.Key = Input

        Dim iv As Byte() = New Byte(BLOCK_SIZE - 1) {}
        For i = 0 To BLOCK_SIZE - 1
            iv(i) = 0
        Next
        r.IV = HexStringToByteArray(transKey)
        ' r.IV = iv
        r.Mode = System.Security.Cryptography.CipherMode.ECB
        r.Padding = System.Security.Cryptography.PaddingMode.None
        'Dim encryptor As System.Security.Cryptography.ICryptoTransform = r.CreateEncryptor(r.Key, r.IV)
        Dim encryptor As System.Security.Cryptography.ICryptoTransform = r.CreateEncryptor()
        Dim Tempresult As Byte() = encryptor.TransformFinalBlock(transKeyArray, 0, transKeyArray.Length)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("applySamyGOKeyTransform for UPnPDevice =  " & MyUPnPDeviceName & " Tempctx: " & ByteArrayToHexString(Tempresult), LogType.LOG_TYPE_INFO)

        Return encryptor.TransformFinalBlock(Input, 0, Input.Length)

    End Function

#Region "Queue handling"
    Private NotificationHandlerReEntryFlag As Boolean = False
    Private MyNotificationQueue As Queue(Of String) = New Queue(Of String)()
    Private MissedNotificationHandlerFlag As Boolean = False
    Friend WithEvents MyNotifyTimer As Timers.Timer

    Private Sub HandleDataReceived(sender As Object, e As String)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungRemote.HandleDataReceived received Line = " & e, LogType.LOG_TYPE_INFO)
        Try
            SyncLock (MyNotificationQueue)
                MyNotificationQueue.Enqueue(e)
            End SyncLock
            MyNotifyTimer.Enabled = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlOff Then Log("Error in SamsungRemote.HandleDataReceived queuing the Notification = " & e.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_INFO)
        End Try
        sender = Nothing
        e = Nothing
    End Sub



    Private Sub TreatNotficationQueue()
        If NotificationHandlerReEntryFlag Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungRemote.TreatNotficationQueue has Re-Entry while processing Notification queue with # elements = " & MyNotificationQueue.Count.ToString, LogType.LOG_TYPE_WARNING)
            MissedNotificationHandlerFlag = True
            Exit Sub
        End If
        NotificationHandlerReEntryFlag = True
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SamsungRemote.TreatNotficationQueue is processing Notification queue with # elements = " & MyNotificationQueue.Count.ToString, LogType.LOG_TYPE_INFO)
        Dim NotificationEvent As String = ""

        Try
            While MyNotificationQueue.Count > 0

                SyncLock (MyNotificationQueue)
                    NotificationEvent = MyNotificationQueue.Dequeue
                End SyncLock
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SamsungRemote.TreatNotficationQueue is processing Notification = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
                Try
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlOff Then Log("Error in SamsungRemote.TreatNotficationQueue for Event = " & NotificationEvent.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End While
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlOff Then Log("Error in SamsungRemote.TreatNotficationQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        NotificationEvent = ""
        If MissedNotificationHandlerFlag Then MyNotifyTimer.Enabled = True ' rearm the timer to prevent events from getting lost
        MissedNotificationHandlerFlag = False
        NotificationHandlerReEntryFlag = False
    End Sub

#End Region

End Class