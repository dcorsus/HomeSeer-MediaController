Imports System.Net
Imports System.Text
Imports System.Net.WebSockets
Imports System.Threading
Imports System.Threading.Tasks

Public Class NetWebSocket

    Protected Friend ws As System.Net.WebSockets.ClientWebSocket = Nothing
    Private _webSocketIsOpened As Boolean = False
    Private cancelToken As CancellationTokenSource
    Protected Friend nbrOfOpenWebAttempts As Integer = 0
    Protected Friend sendSequenceNbr As Integer = 0
    Protected Friend receiveSequenceNbr As Integer = 0
    Friend WithEvents myWebSocketSendTimer As Timers.Timer
    Public Delegate Sub NewMsgEventHandler(msg As String)
    Public Event NewMsgReceived As NewMsgEventHandler
    Public Delegate Sub wsStateChangeEventHandler(isOpen As Boolean)
    Public Event wsStateChange As wsStateChangeEventHandler
    Private sendQueue As New Queue(Of String)
    Private reEntrySendWS As Boolean = False

    Protected Friend Property webSocketIsOpened As Boolean
        Get
            Return _webSocketIsOpened
        End Get
        Set(value As Boolean)
            If _webSocketIsOpened <> value Then
                Try
                    RaiseEvent wsStateChange(value)
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlOff Then Log("Error in webSocketIsOpened raising wsStateChange event with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                _webSocketIsOpened = value
            End If
        End Set
    End Property

    Protected Friend Async Sub OpenWebSocket(webSocketUrl As String)
        If PIDebuglevel > DebugLevel.dlEvents Then
            Log("OpenWebSocket called with Url = " & webSocketUrl, LogType.LOG_TYPE_INFO)
        ElseIf PIDebuglevel > DebugLevel.dlErrorsOnly Then
            Log("OpenWebSocket called", LogType.LOG_TYPE_INFO)
        End If
        ' optional: ignore certificate errors
        nbrOfOpenWebAttempts += 1
        cancelToken = New CancellationTokenSource
        ws = New System.Net.WebSockets.ClientWebSocket
        If ws Is Nothing Then
            If PIDebuglevel > DebugLevel.dlOff Then Log("Error in OpenWebSocket trying to get a ClientWebSocket", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        ServicePointManager.ServerCertificateValidationCallback = Function(s, c, h, d) True
        ws.Options.UseDefaultCredentials = True
        Try
            Await ws.ConnectAsync(New Uri(webSocketUrl), cancelToken.Token)
            If (ws.State = WebSockets.WebSocketState.Open) Then
                ' initialize the websocket queue, used for alarm systems
                sendSequenceNbr = 0
                If myWebSocketSendTimer Is Nothing Then
                    myWebSocketSendTimer = New Timers.Timer With {
                        .Interval = 100,
                        .AutoReset = False,
                        .Enabled = False
                    }
                End If
                webSocketIsOpened = True
                nbrOfOpenWebAttempts = 0    ' reset this counter
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("OpenWebSocket opened with Url = " & webSocketUrl, LogType.LOG_TYPE_INFO)
                While ws.State = WebSockets.WebSocketState.Open
                    Dim bytes(16384) As Byte
                    Dim answ = New ArraySegment(Of Byte)(bytes)
                    Dim receiveTask As Task(Of WebSocketReceiveResult) = ws.ReceiveAsync(answ, cancelToken.Token)
                    Dim recResult As WebSocketReceiveResult = Await receiveTask
                    Dim answString As String = System.Text.Encoding.UTF8.GetString(answ.Array, 0, recResult.Count)
                    While Not recResult.EndOfMessage And recResult.Count > 0
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning OpenWebSocket received fractial part ", LogType.LOG_TYPE_WARNING)
                        receiveTask = ws.ReceiveAsync(answ, cancelToken.Token)
                        recResult = Await receiveTask
                        If recResult.Count > 0 Then answString &= System.Text.Encoding.UTF8.GetString(answ.Array, 0, recResult.Count)
                    End While
                    Try
                        RaiseEvent NewMsgReceived(answString)
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlOff Then Log("Error in TreatWebSocketMessage raising newMsg event with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("OpenWebSocket received = " & System.Text.Encoding.UTF8.GetString(answ.Array), LogType.LOG_TYPE_INFO)
                End While
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in OpenWebSocket. The WS is not open anymore. State = " & ws.State.ToString, LogType.LOG_TYPE_WARNING)
                closeWebSocket()
            Else
                If PIDebuglevel > DebugLevel.dlOff Then Log("Warning OpenWebSocket couldn't be opened with Url = " & webSocketUrl, LogType.LOG_TYPE_WARNING)
                closeWebSocket()
            End If
        Catch ex As Exception
            ' this is called when the socket gets close. I'll change the reporting to a warning 
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning in OpenWebSocket with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            closeWebSocket()
        End Try
    End Sub

    Protected Friend Sub closeWebSocket()
        Try
            If myWebSocketSendTimer IsNot Nothing Then
                myWebSocketSendTimer.Stop()
                myWebSocketSendTimer.Enabled = False
                myWebSocketSendTimer = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If cancelToken IsNot Nothing Then
                cancelToken.Cancel()
                cancelToken = Nothing
            End If
            ws = Nothing
        Catch ex As Exception
        End Try
        webSocketIsOpened = False
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning webSocket closed", LogType.LOG_TYPE_WARNING)
    End Sub

    Protected Friend Async Sub SendWebSocketMessage(msg As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendWebSocketMessage called with Msg = " & msg, LogType.LOG_TYPE_INFO)
        Try
            SyncLock (sendQueue)
                sendQueue.Enqueue(msg)
            End SyncLock
            If reEntrySendWS Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendWebSocketMessage reEntrance with #Msgs in Queue = " & sendQueue.Count, LogType.LOG_TYPE_WARNING)
                Exit Sub
            End If
            reEntrySendWS = True
            Dim queuedMsg As String = ""
            While sendQueue.Count > 0
                SyncLock (sendQueue)
                    queuedMsg = sendQueue.Dequeue()
                End SyncLock
                If queuedMsg <> "" Then
                    Try
                        If ws IsNot Nothing AndAlso ws.State = WebSockets.WebSocketState.Open Then
                            Await ws.SendAsync(New ArraySegment(Of Byte)(Encoding.UTF8.GetBytes(msg)), WebSockets.WebSocketMessageType.Text, True, cancelToken.Token)
                        End If
                    Catch ex As Exception
                        If PIDebuglevel > DebugLevel.dlOff Then Log("Error in SendWebSocketMessage with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        closeWebSocket()
                    End Try
                End If
            End While
        Catch ex As Exception
        End Try
        reEntrySendWS = False
    End Sub


End Class
