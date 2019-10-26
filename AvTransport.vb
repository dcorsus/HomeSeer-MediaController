Imports System.Xml
Imports System.Text
Imports System.Drawing

Partial Public Class HSPI

    Private AVTransport As MyUPnPService = Nothing
    Public WithEvents myAVTransportCallback As New myUPnPControlCallback

    Private MyCurrentArtist As String = ""
    Private MyCurrentAlbum As String = ""
    Private MyNextTrack As String = ""
    Private MyNextArtist As String = ""
    Private MyNextAlbum As String = ""
    Private MyNextAlbumURI As String = ""
    Private MyCurrentPlayerState As player_state_values = player_state_values.stopped
    Private MyCurrentArtworkURL As String = NoArtPath
    Private MyCurrentTransportState As String = ""
    Private MyCurrentTransportStatus As String = ""
    Private MyCurrentNrTracks As Integer = 0
    Private MyCurrentMediaDuration As String = ""
    Private MyCurrentURI As String = ""
    Private MyCurrentURIMetaData As String = ""
    Private MyNextURI As String = ""
    Private MyNextURIMetaData As String = ""
    Private MyCurrentPlayMedium As String = ""
    Private MyCurrentRecordMedium As String = ""
    Private MyCurrentWriteStatus As String = ""
    Private MyCurrentPlayMedia As String = ""
    Private MyCurrentRecMedia As String = ""
    Private MyCurrentRecQualityModes As String = ""
    Private MyCurrentPlayMode As String = ""
    Private MyCurrentRecQualityMode As String = ""
    'Private MyCurrentSpeed As String = ""
    Private MyCurrentTrack As String = ""
    Private MyCurrentTrackNumber As String = ""
    Private MyCurrentTrackDuration As String = ""
    Private MyCurrentTrackMetaData As String = ""
    Private MyCurrentTrackURI As String = ""
    Private MyCurrentRelTime As String = ""
    Private MyCurrentAbsTime As String = ""
    Private MyCurrentRelCount As Integer = 0
    Private MyCurrentAbsCount As Integer = 0
    Private MyActions As String = ""
    Private MyCurrentTrackSize As String = ""
    Private MyCurrentRelByte As String = ""
    Private MyCurrentAbsByte As String = ""
    Private MyCurrentPlayerPosition As Integer = 0
    Private MyPlayStateHSRef As Integer = -1
    Private MyCurrentTransportErrorURI As String = ""
    Private MyCurrentTransportErrorDescription As String = ""
    Private MyCurrentTransportPlaySpeed As Integer = 1
    Private MyCurrentNumberOfTracks As String = ""
    Private MyCurrentTransportActions As String = ""
    Private MyCurrentPossiblePlaybackStorageMedia As String = ""
    Private MyCurrentPossibleRecordStorageMedia As String = ""
    Private MyCurrentPossibleRecordQualityMode As String = ""
    Private MyPossibleRecordQualityModes As String = ""
    Private MyCurrentCreator As String = ""
    Private MyCurrentAuthor As String = ""
    Private MyCurrentGenre As String = ""
    Private MyCurrentTrackDate As String = ""
    Private MyCurrentOriginalTrackNumber As String = ""
    Private MyCurrentClass As String = ""
    Private MyNextCreator As String = ""
    Private MyNextAuthor As String = ""
    Private MyNextGenre As String = ""
    Private MyNextTrackDate As String = ""
    Private MyNextTrackMetaData As String = ""
    Private MyNextOriginalTrackNumber As String = ""
    Private MyNextClass As String = ""
    Private MyTransportStateHasChanged As Boolean = False
    Private MyTrackInfoHasChanged As Boolean = False
    Private MyTittleWasPresent As Boolean = False
    Private MyAlbumArtURIHasChanged As Boolean = False
    Private MyNextTrackInfoHasChanged As Boolean = False
    Private MyNextAlbumArtURIHasChanged As Boolean = False
    Private MyDurationInfoHasChanged As Boolean = False
    Private MyNextArtworkURL As String = NoArtPath
    Private MyQueueLinkedList As New LinkedList(Of myQueueElement)
    Private MyQueuePlayState As player_state_values
    Private MyPlayFromQueue As Boolean = False
    Private MyCurrentQueueIndex As Integer = 1
    Private MyCurrentQueueObjectID As String = ""
    Private MyCurrentQueueServerUDN As String = ""
    Private MyQueueDelay As Integer = 0
    Private MyQueueDelayUserDefined As Integer = 20
    Private MyNbrOfQueueItemsPlayed As Integer = 0
    Private MyNextQueueIndex As Integer = 1
    Private MyNextQueueObjectID As String = ""
    Private MyNextQueueServerUDN As String = ""
    Private MyPreviousQueueIndex As Integer = 1
    Private MyPrefetchNextQueueObject As Boolean = False
    Private NextElementPrefetchState As AVT_PrefetchState = AVT_PrefetchState.PfSNone
    Private MyQueueReEntry As Boolean = False
    Private MyIssueStopInBetween As Boolean = True
    Private NextAvTransportIsAvailable As Boolean = False
    Private UseNextAvTransport As Boolean = True
    Private MySavedTrackPosition As String = ""
    Private MySavedQueueIndex As Integer = 1
    Private strApplySeek As String = ""

    Private AlbumArtURIWasFound As Boolean = False
    Private AlbumArtSizeFound As String = ""
    Private AlbumArtURIFound As String = ""
    Private ProcessAVXMLReEntryFlag As Boolean = False
    Private MyCurrentResolution As String = ""

    Private NewActor As String = ""
    Private NewAuthor As String = ""
    Private NewDesc As String = ""
    Private NewArtist As String = ""
    Private NewFullArtist As String = ""

    Public MyQueueHasChanged As Boolean = False

    ' Info about DLNA
    ' ===============
    ' DLNA.ORG_PS: play speed parameter (integer)
    '       0 invalid play speed
    '       1 normal play speed*/

    ' DLNA.ORG_CI: conversion indicator parameter (integer)
    '       0 not transcoded
    '       1 transcoded

    ' DLNA.ORG_OP: operations parameter (string)
    '       "00" (or "0") neither time seek range nor range supported
    '       "01" range supported
    '       "10" time seek range supported
    '       "11" both time seek range and range supported

    ' DLNA.ORG_FLAGS, padded with 24 trailing 0s
    '     80000000  31  senderPaced
    '     40000000  30  lsopTimeBasedSeekSupported
    '     20000000  29  lsopByteBasedSeekSupported
    '     10000000  28  playcontainerSupported
    '      8000000  27  s0IncreasingSupported
    '      4000000  26  sNIncreasingSupported
    '      2000000  25  rtspPauseSupported
    '      1000000  24  streamingTransferModeSupported
    '       800000  23  interactiveTransferModeSupported
    '       400000  22  backgroundTransferModeSupported
    '       200000  21  connectionStallingSupported
    '       100000  20  dlnaVersion15Supported
    '
    '     Example: (1 << 24) | (1 << 22) | (1 << 21) | (1 << 20)
    '       DLNA.ORG_FLAGS=01700000[000000000000000000000000] // [] show padding

    Public WriteOnly Property ApplySeek As String
        Set(value As String)
            strApplySeek = value
        End Set
    End Property

    Public ReadOnly Property CurrentContentClass As UPnPClassType
        Get
            CurrentContentClass = ProcessClassInfo(MyCurrentClass)
        End Get
    End Property

    Public ReadOnly Property NextContentClass As UPnPClassType
        Get
            NextContentClass = ProcessClassInfo(MyNextClass)
        End Get
    End Property

    Public ReadOnly Property Resolution As String
        Get
            Resolution = MyCurrentResolution
        End Get
    End Property

    Private Sub TransportStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myAVTransportCallback.ControlStateChange 'TransportStateChange
        If PIDebuglevel > DebugLevel.dlEvents Then
            Log("TransportChangeCallback for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ElseIf PIDebuglevel > DebugLevel.dlErrorsOnly Then
            'Log("TransportChangeCallback for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
            Log("TransportChangeCallback for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName, LogType.LOG_TYPE_INFO)
        End If
        If StateVarName <> "LastChange" Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TransportStateChange callback for device = " & MyUPnPDeviceName & " received an unknown Variable = " & StateVarName & " and data = " & Value.ToString, LogType.LOG_TYPE_WARNING)
            Exit Sub
        End If
        MyTrackInfoHasChanged = False
        MyTransportStateHasChanged = False
        MyAlbumArtURIHasChanged = False
        MyNextTrackInfoHasChanged = False
        MyNextAlbumArtURIHasChanged = False
        MyTittleWasPresent = False
        MyDurationInfoHasChanged = False
        ProcessAVXML(Value, False, False, (PIDebuglevel > DebugLevel.dlErrorsOnly Or PIDebuglevel > DebugLevel.dlEvents))
        'If (Not MyPollForTransportChangeFlag) And (Not MyDurationInfoHasChanged) Then GetPositionDurationInfoOnly(0) 'SetPlayerPosition(0) ' do this because streaming does not provide info and HST doesn't update screens but this is a problem for LG-SP520
        If MyTrackInfoHasChanged Then MyPlayerWentThroughTrackChange = True
        If NextAvTransportIsAvailable And UseNextAvTransport Then CheckNextURIisPresent()
        If PIDebuglevel > DebugLevel.dlEvents Then Log("TransportChangeCallback for device = " & MyUPnPDeviceName & ". VarName = " & StateVarName & " = DONE!!", LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub TransportDied() Handles myAVTransportCallback.ControlDied 'TransportDied
        Log("TransportDied for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_WARNING)
        Try
            'Log("UPnP connection to device " & MyUPnPDeviceName & " was lost in TransportDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("Error in TransportDied for device = " & MyUPnPDeviceName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub CheckNextURIisPresent()
        If MyTrackInfoHasChanged And MyPlayerWentThroughPlayState Then
            MyCurrentQueueIndex = MyNextQueueIndex
            MyCurrentQueueObjectID = MyNextQueueObjectID
            MyCurrentQueueServerUDN = MyNextQueueServerUDN
            NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckNextURIisPresent for UPnPDevice = " & MyUPnPDeviceName & " has set GetNext flag because the trackinfo changed", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If MyTransportStateHasChanged Or (MyCurrentPlayerState <> player_state_values.Playing And MyCurrentPlayerState <> player_state_values.Forwarding And MyCurrentPlayerState <> player_state_values.Rewinding) Or Not MyTittleWasPresent Then
            ' do nothing
            Exit Sub
        End If
        ' so we had no track change, no transport state change and the player is playing, this could be either ff  or rev or next track which is same as current
        AVTGetMediaInfo(0, PIDebuglevel > DebugLevel.dlEvents)
        If MyNextURI = "" And MyNextURIMetaData = "" Then
            MyCurrentQueueIndex = MyNextQueueIndex
            MyCurrentQueueObjectID = MyNextQueueObjectID
            MyCurrentQueueServerUDN = MyNextQueueServerUDN
            NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNextURIisPresent for UPnPDevice = " & MyUPnPDeviceName & " has set GetNext flag because the NextURI was empty", LogType.LOG_TYPE_INFO)
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNextURIisPresent called for UPnPDevice = " & MyUPnPDeviceName & " with NextURI = " & MyNextURI & " with NextURIMetaData = " & MyNextURIMetaData, LogType.LOG_TYPE_INFO)
    End Sub


    Private Sub ProcessAVXML(inMetaData As String, TrackData As Boolean, NextData As Boolean, PrintDebug As Boolean)
        If PIDebuglevel > DebugLevel.dlEvents Then
            Log("ProcessAVXML called for UPnPDevice = " & MyUPnPDeviceName & " with TrackData = " & TrackData.ToString & " and NextData = " & NextData.ToString & " and XML = " & inMetaData.ToString, LogType.LOG_TYPE_INFO)
        End If
        If ProcessAVXMLReEntryFlag Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("REENTRY in ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " !!!!", LogType.LOG_TYPE_WARNING)
            Exit Sub
        End If
        ProcessAVXMLReEntryFlag = True
        If Trim(inMetaData.ToString) = "" Then
            ProcessAVXMLReEntryFlag = False
            Exit Sub
        End If
        If inMetaData.ToString.ToUpper = "NOT_IMPLEMENTED" Then
            ProcessAVXMLReEntryFlag = False
            Exit Sub
        End If

        Dim xmlData As XmlDocument = New XmlDocument
        xmlData.XmlResolver = Nothing

        Try
            inMetaData = CheckNameSpaceXMLData(inMetaData)
            inMetaData = CheckForXMLIssues(inMetaData)
            xmlData.LoadXml(inMetaData.ToString)
            'log( "XML=" & Value.ToString) ' used for testing
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & inMetaData.ToString, LogType.LOG_TYPE_ERROR)
            ProcessAVXMLReEntryFlag = False
            Exit Sub
        End Try

        MyCurrentTrackMetaData = ""
        MyNextTrackMetaData = ""

        Try
            If xmlData.HasChildNodes And (Not TrackData) Then
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO) ' this starts with <Event>
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node name: " & outerNode.Name & " and ID = " & outerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO) ' this will be InstanceID
                    'Check if this matches with our selected item
                    If outerNode.HasChildNodes Then
                        For Each InnerNode As XmlNode In outerNode.ChildNodes
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Name: " & InnerNode.Name, LogType.LOG_TYPE_INFO)
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Value: " & InnerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO) ' Here are the Values
                            ProcessAVTransportXML(InnerNode.Name, InnerNode.Attributes("val").Value, InnerNode.OuterXml, PrintDebug)
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & "  processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try


        Dim Metadata As String = ""

        If TrackData Then
            Metadata = inMetaData
        Else
            Metadata = MyCurrentTrackMetaData
        End If

        Dim TreatNextData As Boolean = NextData

        ' author and desc Artist



        For LoopCounter = 1 To 2
            AlbumArtURIWasFound = False
            AlbumArtSizeFound = ""
            AlbumArtURIFound = ""
            NewActor = ""
            NewDesc = ""
            NewArtist = ""
            NewFullArtist = ""
            NewAuthor = ""

            If Metadata <> "" And Metadata.ToString.ToUpper <> "NOT_IMPLEMENTED" Then
                Try
                    Metadata = CheckNameSpaceXMLData(Metadata)
                    Metadata = CheckForXMLIssues(Metadata)
                    xmlData.LoadXml(Metadata)
                    'Get a list of all the child elements
                    If xmlData.HasChildNodes Then
                        Dim Nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & Nodelist.Count, LogType.LOG_TYPE_INFO) ' this starts with <Event>
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                        For Each outerNode As XmlNode In Nodelist
                            Try
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node name: " & outerNode.Name, LogType.LOG_TYPE_INFO) ' & " and ID = " & outerNode.Attributes("val").Value) ' this will be InstanceID
                                If UCase(outerNode.Name) = "ITEM" Then ' need to go deeper
                                    Try
                                        If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node has Children = " & outerNode.HasChildNodes.ToString, LogType.LOG_TYPE_INFO)
                                        If outerNode.HasChildNodes Then
                                            For Each InnerNode As XmlNode In outerNode.ChildNodes
                                                Try
                                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Name: " & InnerNode.Name, LogType.LOG_TYPE_INFO)
                                                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Value: " & InnerNode.InnerText, LogType.LOG_TYPE_INFO)
                                                    ProcessAVMetaData(InnerNode.Name, InnerNode.InnerText, InnerNode.Attributes, TreatNextData, InnerNode.OuterXml, PrintDebug)
                                                Catch ex As Exception
                                                    Log("Error in ProcessAVXML 1 for device - " & MyUPnPDeviceName & " loading XML MetaData. XML = " & Metadata.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                    Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & Nodelist.Count, LogType.LOG_TYPE_INFO) ' 
                                                    Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                                                    Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node name: " & outerNode.Name & " and ID = " & outerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO)
                                                    Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Name: " & InnerNode.Name, LogType.LOG_TYPE_INFO)
                                                    Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & "------> Inner node Value: " & InnerNode.InnerText, LogType.LOG_TYPE_INFO)
                                                End Try
                                            Next
                                        End If
                                    Catch ex As Exception
                                        Log("Error in ProcessAVXML 2 for device - " & MyUPnPDeviceName & " loading XML MetaData. XML = " & Metadata.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                        Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & Nodelist.Count, LogType.LOG_TYPE_INFO) ' 
                                        Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                                        Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node name: " & outerNode.Name & " and ID = " & outerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO)
                                    End Try
                                Else
                                    ProcessAVMetaData(outerNode.Name, outerNode.InnerText, outerNode.Attributes, TreatNextData, outerNode.OuterXml, PrintDebug)
                                End If

                            Catch ex As Exception
                                Log("Error in ProcessAVXML 3 for device - " & MyUPnPDeviceName & " loading XML MetaData. XML = " & Metadata.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Nbr of items in XML Data = " & Nodelist.Count, LogType.LOG_TYPE_INFO) ' 
                                Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                                Log("-------->ProcessAVXML for device - " & MyUPnPDeviceName & " Outer node name: " & outerNode.Name & " and ID = " & outerNode.Attributes("val").Value, LogType.LOG_TYPE_INFO)
                            End Try
                        Next
                    End If
                    If AlbumArtURIWasFound Then
                        If TreatNextData = False Then
                            If MyCurrentArtworkURL <> AlbumArtURIFound Then
                                MyAlbumArtURIHasChanged = True
                                ArtworkURL = GetAlbumArtPath(AlbumArtURIFound, False)
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Album Art URI) = " & MyCurrentArtworkURL, LogType.LOG_TYPE_INFO)
                        Else
                            If MyNextArtworkURL <> AlbumArtURIFound Then
                                MyNextAlbumArtURIHasChanged = True
                                NextArtworkURL = GetAlbumArtPath(AlbumArtURIFound, True)
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Next Album Art URI) = " & MyNextArtworkURL, LogType.LOG_TYPE_INFO)
                        End If
                    End If
                    If NewActor <> "" Then
                        If TreatNextData = False Then

                        Else

                        End If
                    End If
                    If NewAuthor <> "" Then
                        If TreatNextData = False Then
                            MyCurrentAuthor = NewAuthor
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Author) = " & MyCurrentAuthor, LogType.LOG_TYPE_INFO)
                        Else
                            MyNextAuthor = NewAuthor
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Next Author) = " & MyNextAuthor, LogType.LOG_TYPE_INFO)
                        End If
                    End If
                    If NewFullArtist <> "" Then
                        If TreatNextData = False Then
                            If MyCurrentArtist <> NewFullArtist Then
                                MyTrackInfoHasChanged = True
                                If HSRefArtist <> -1 Then
                                    If NewArtist <> "" Then
                                        hs.SetDeviceString(HSRefArtist, NewArtist, True)
                                    Else
                                        hs.SetDeviceString(HSRefArtist, NewFullArtist, True)
                                    End If
                                End If
                                MyCurrentArtist = NewFullArtist
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Artist) = " & MyCurrentArtist, LogType.LOG_TYPE_INFO)
                        Else
                            If MyNextArtist <> NewFullArtist Then
                                MyNextTrackInfoHasChanged = True
                                If HSRefNextArtist <> -1 Then
                                    If NewArtist <> "" Then
                                        hs.SetDeviceString(HSRefNextArtist, NewArtist, True)
                                    Else
                                        hs.SetDeviceString(HSRefNextArtist, NewFullArtist, True)
                                    End If
                                End If
                                MyNextArtist = NewFullArtist
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received (Next Artist) = " & MyNextArtist, LogType.LOG_TYPE_INFO)
                        End If
                    End If
                    If NewDesc <> "" Then
                        If TreatNextData = False Then
                            CurrentTrackDescription = NewDesc
                        Else
                            NextTrackDescription = NewDesc
                        End If
                    End If
                Catch ex As Exception
                    If Trim(Metadata) <> "" Then
                        Log("Error in ProcessAVXML 4 for device - " & MyUPnPDeviceName & " loading Meta XML MetaData. XML = " & Metadata.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("-------->ProcessAVXML 4 for device - " & MyUPnPDeviceName & " loading root XML MetaData. XML = " & inMetaData.ToString, LogType.LOG_TYPE_INFO)
                        ProcessAVXMLReEntryFlag = False
                        Exit Sub
                    End If
                End Try
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning in ProcessAVXML for device - " & MyUPnPDeviceName & ". TrackNode is empty ", LogType.LOG_TYPE_WARNING)
            End If
            If MyNextTrackMetaData <> "" And MyNextTrackMetaData.ToString.ToUpper <> "NOT_IMPLEMENTED" And Not TrackData Then
                Metadata = MyNextTrackMetaData
                TreatNextData = True
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML for UPnPDevice = " & MyUPnPDeviceName & " received next Track Data and is processing it", LogType.LOG_TYPE_INFO)
            Else
                Exit For
            End If
        Next
        xmlData = Nothing
        If MyCurrentPlayerState = player_state_values.Transitioning Then
            ProcessAVXMLReEntryFlag = False
            Exit Sub ' we're done
        End If
        UpdateTransportState()
        If PIDebuglevel > DebugLevel.dlEvents Then Log("ProcessAVXML called for UPnPDevice = " & MyUPnPDeviceName & " is Done ", LogType.LOG_TYPE_INFO)
        ProcessAVXMLReEntryFlag = False
    End Sub

    Private Sub UpdateTransportState()
        Try
            TransportInfo = "<table><tr><td><img src=" & MyCurrentArtworkURL
            If ArtworkVSize <> 0 Then
                TransportInfo = TransportInfo & " height=""" & ArtworkVSize.ToString & """ "
                If ArtworkHSize <> 0 Then
                    TransportInfo = TransportInfo & " width=""" & ArtworkHSize.ToString & """ "
                End If
            End If
            TransportInfo = TransportInfo & "></td><td><p>" & MyCurrentPlayerState.ToString & "</p><p>" & MyCurrentArtist & "</p><p>" & MyCurrentAlbum & "</p><p>" & MyCurrentTrack & "</p></td></tr></table>"
            Try
                If MyTransportStateHasChanged Then
                    'If HSRefPlayer <> -1 Then hs.SetDeviceValueByRef(HSRefPlayState, MyCurrentPlayerState, True)
                    PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, MyCurrentPlayerState)
                End If
                If MyTrackInfoHasChanged Then
                    PlayChangeNotifyCallback(player_status_change.SongChanged, MyCurrentPlayerState)
                End If
                If MyTransportStateHasChanged Or MyTrackInfoHasChanged Then
                    If HSRefPlayer <> -1 Then hs.SetDeviceString(HSRefPlayer, TransportInfo, True)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("HS updated in UpdateTransportState. HSRef = " & HSRefPlayer & " and MyTransportStateHasChanged = " & MyTransportStateHasChanged.ToString & ", MyTrackInfoHasChanged = " & MyTrackInfoHasChanged.ToString & ". Info = " & TransportInfo, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
                Log("Error in UpdateTransportState updating HS with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("ERROR in UpdateTransportState 6 for UPnPDevice = " & MyUPnPDeviceName & " with error=" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function CheckNameSpaceXMLData(InData As String) As String
        If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:dlna="urn:schemas-dlna-org:metadata-1-0/" xmlns:pv="http://www.pv.com/pvns/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/"
        If (InData.IndexOf("xmlns:dlna") = -1) And (InData.IndexOf("dlna:") <> -1) Then
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                Dim FirstBlank As Integer = InData.IndexOf(" ")
                builder.Insert(FirstBlank, " xmlns:dlna=""urn:schemas-dlna-org:metadata-1-0/"" ")
                InData = builder.ToString
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " added dlna tag the MetaData", LogType.LOG_TYPE_WARNING)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_INFO)
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckNameSpaceXMLData in for device - " & MyUPnPDeviceName & " looking for dlna tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If (InData.IndexOf("xmlns:dc") = -1) And (InData.IndexOf("dc:") <> -1) Then
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                Dim FirstBlank As Integer = InData.IndexOf(" ")
                builder.Insert(FirstBlank, " xmlns:dc=""http://purl.org/dc/elements/1.1/"" ")
                InData = builder.ToString
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " added dc tag the MetaData", LogType.LOG_TYPE_WARNING)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_INFO)
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckNameSpaceXMLData in for device - " & MyUPnPDeviceName & " looking for dc tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If (InData.IndexOf("xmlns:upnp") = -1) And (InData.IndexOf("upnp:") <> -1) Then
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                Dim FirstBlank As Integer = InData.IndexOf(" ")
                builder.Insert(FirstBlank, " xmlns:upnp=""urn:schemas-upnp-org:metadata-1-0/upnp/"" ")
                InData = builder.ToString
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " added unpn tag the MetaData", LogType.LOG_TYPE_WARNING)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_INFO)
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckNameSpaceXMLData in for device - " & MyUPnPDeviceName & " looking for upnp tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If (InData.IndexOf("xmlns:pv") = -1) And (InData.IndexOf("pv:") <> -1) Then
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                Dim FirstBlank As Integer = InData.IndexOf(" ")
                builder.Insert(FirstBlank, " xmlns:pv=""http://www.pv.com/pvns/"" ")
                InData = builder.ToString
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " added pv tag the MetaData", LogType.LOG_TYPE_WARNING)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_INFO)
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckNameSpaceXMLData in for device - " & MyUPnPDeviceName & " looking for pv tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If (InData.IndexOf("xmlns:av") = -1) And (InData.IndexOf("av:") <> -1) Then
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                Dim FirstBlank As Integer = InData.IndexOf(" ")
                builder.Insert(FirstBlank, " xmlns:av=""urn:schemas-sony-com:av"" ")
                InData = builder.ToString
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " added av tag the MetaData", LogType.LOG_TYPE_WARNING)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckNameSpaceXMLData for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_INFO)
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckNameSpaceXMLData in for device - " & MyUPnPDeviceName & " looking for av tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        CheckNameSpaceXMLData = InData
    End Function

    Private Function CheckForXMLIssues(InData As String) As String
        'If piDebuglevel > DebugLevel.dlEvents Then log( "CheckForXMLIssues called for device - " & MyUPnPDeviceName)
        CheckForXMLIssues = InData
        Try
            Dim xmlData As XmlDocument = New XmlDocument
            xmlData.LoadXml(InData)
            ' we survived, just return
            xmlData = Nothing
            Exit Function
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning CheckForXMLIssues in for device - " & MyUPnPDeviceName & " and Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning CheckForXMLIssues in for device - " & MyUPnPDeviceName & " with XML = " & InData, LogType.LOG_TYPE_WARNING)
        End Try
        If InData.IndexOf("""xmlns") <> -1 Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Warning CheckForXMLIssues in for device - " & MyUPnPDeviceName & ", issue found with missing whitespace for XMLNS", LogType.LOG_TYPE_WARNING)
            Try
                Dim builder As New System.Text.StringBuilder(InData.ToString)
                builder.Replace("""xmlns", """ xmlns")
                InData = builder.ToString
                builder = Nothing
            Catch ex As Exception
                Log("Error CheckForXMLIssues in for device - " & MyUPnPDeviceName & " looking for ""xmlns and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        ' This part from WD TV gives us heartburn. It doesn't like the first < character as an invalid character attribute
        ' <item id="BROWSE_TYPE_KEYaudio*~>_<~*BROWSE_SORT_KEY8*~>_<~*BROWSE_GROUP_KEY0*~>_<~*BROWSE_FILTER_KEY*~>_<~*13 - Dance Sister Dance (Baila Mi Hermana).mp3" 
        ' parentID="BROWSE_TYPE_KEYaudio*~>_<~*BROWSE_SORT_KEY8*~>_<~*BROWSE_GROUP_KEY0*~>_<~*BROWSE_FILTER_KEYBROWSE_FILTER_TERM_STARTfilepath:<_EQUALS_>:/tmp/media/usb/Local/WDTVLiveHub/Santana/The Best Of SantanaBROWSE_FILTER_TERM_ENDBROWSE_FILTER_TERM_STARTname:<_EQUALS_>:13 - Dance Sister Dance (Baila Mi Hermana).mp3BROWSE_FILTER_TERM_END*~>_<~*" restricted="0">
        ' *~>_<~*
        ' :<_EQUALS_>:
        Try
            Dim UpdateHappened As Boolean = False
            Dim SafetyRetry As Integer = 0
            If InData.IndexOf(vbCr) <> -1 Then
                Log("CheckForXMLIssues in for device - " & MyUPnPDeviceName & " found vbCR at position = " & InData.IndexOf(vbCr).ToString & " for total data Length = " & InData.Length.ToString, LogType.LOG_TYPE_WARNING)
            End If
            If InData.IndexOf(vbLf) <> -1 Then
                Log("CheckForXMLIssues in for device - " & MyUPnPDeviceName & " found vbLf at position = " & InData.IndexOf(vbLf).ToString & " for total data Length = " & InData.Length.ToString, LogType.LOG_TYPE_WARNING)
            End If
            While InData.IndexOf("*~>_<~*") <> -1
                ' this is found in the <item tag, the ID attribute or parentid attribute
                ' I'm going to try to delete these attributes 
                Try
                    Dim builder As New System.Text.StringBuilder(InData.ToString)
                    builder.Remove(InData.IndexOf("*~>_<~*"), 7)
                    UpdateHappened = True
                    InData = builder.ToString
                    builder = Nothing
                Catch ex As Exception
                    Log("Error CheckForXMLIssues in for device - " & MyUPnPDeviceName & " looking for *~>_<~* tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                SafetyRetry += 1
                If SafetyRetry > 100 Then Exit While
            End While
            SafetyRetry = 0
            While InData.IndexOf(":<_EQUALS_>:") <> -1
                Try
                    Dim builder As New System.Text.StringBuilder(InData.ToString)
                    builder.Remove(InData.IndexOf(":<_EQUALS_>:"), 12)
                    InData = builder.ToString
                    UpdateHappened = True
                    builder = Nothing
                Catch ex As Exception
                    Log("Error CheckForXMLIssues in for device - " & MyUPnPDeviceName & " looking for :<_EQUALS_>: tag and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                SafetyRetry += 1
                If SafetyRetry > 100 Then Exit While
            End While
            If UpdateHappened Then
                Dim ItemTagStart As Integer = -1
                ItemTagStart = InData.IndexOf("<item ")
                If ItemTagStart <> -1 Then
                    Dim attributeStart As Integer = -1
                    attributeStart = InData.IndexOf("id=""", ItemTagStart)
                    If attributeStart <> -1 Then
                        Dim attributeEnd As Integer = -1
                        attributeEnd = InData.IndexOf("""", attributeStart + 4)
                        If (attributeEnd <> -1) And ((attributeEnd - attributeStart - 4) > 0) Then
                            Dim builder As New System.Text.StringBuilder(InData.ToString)
                            builder.Remove(attributeStart + 4, attributeEnd - attributeStart - 4)
                            InData = builder.ToString
                            builder = Nothing
                        End If
                    End If
                    attributeStart = -1
                    attributeStart = InData.IndexOf("parentid=""", ItemTagStart)
                    If attributeStart <> -1 Then
                        Dim attributeEnd As Integer = -1
                        attributeEnd = InData.IndexOf("""", attributeStart + 10)
                        If (attributeEnd <> -1) And ((attributeEnd - attributeStart - 10) > 0) Then
                            Dim builder As New System.Text.StringBuilder(InData.ToString)
                            builder.Remove(attributeStart + 10, attributeEnd - attributeStart - 10)
                            InData = builder.ToString
                            builder = Nothing
                        End If
                    End If
                    attributeStart = -1
                    attributeStart = InData.IndexOf("parentID=""", ItemTagStart)
                    If attributeStart <> -1 Then
                        Dim attributeEnd As Integer = -1
                        attributeEnd = InData.IndexOf("""", attributeStart + 10)
                        If (attributeEnd <> -1) And ((attributeEnd - attributeStart - 10) > 0) Then
                            Dim builder As New System.Text.StringBuilder(InData.ToString)
                            builder.Remove(attributeStart + 10, attributeEnd - attributeStart - 10)
                            InData = builder.ToString
                            builder = Nothing
                        End If
                    End If
                    'Change all & to &amp; First find the proper ones = &amp;, &quot;, &lt;, &gt;
                    Dim strbuilder As New System.Text.StringBuilder(InData.ToString)
                    strbuilder.Replace("&amp;", "#$#1#$#")
                    strbuilder.Replace("&quot;", "#$#2#$#")
                    strbuilder.Replace("&lt;", "#$#3#$#")
                    strbuilder.Replace("&gt;", "#$#4#$#")

                    strbuilder.Replace("&", "&amp;")

                    strbuilder.Replace("#$#1#$#", "&amp;")
                    strbuilder.Replace("#$#2#$#", "&quot;")
                    strbuilder.Replace("#$#3#$#", "&lt;")
                    strbuilder.Replace("#$#4#$#", "&gt;")

                    InData = strbuilder.ToString
                    strbuilder = Nothing
                End If
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in CheckForXMLIssues for device - " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckForXMLIssues for device - " & MyUPnPDeviceName & " changed the inMetaData to = " & InData.ToString, LogType.LOG_TYPE_WARNING)
        CheckForXMLIssues = InData
    End Function

    Private Sub ProcessAVTransportXML(VariableName As String, VariableValue As String, VariableXML As String, PrintDebug As Boolean)
        ' Onkyo
        ' <Event xmlns="urn:schemas-upnp-org:metadata-1-0/AVT/">
        '   <InstanceID val="0">
        '       <TransportState val="NO_MEDIA_PRESENT"/>
        '       <TransportStatus val="ERROR_OCCURRED"/>
        '       <PlaybackStorageMedium val="NONE"/>
        '       <CurrentPlayMode val="NORMAL"/>
        '       <TransportPlaySpeed val="1"/>
        '       <NumberOfTracks val="0"/>
        '       <CurrentTrack val="0"/>
        '       <CurrentTrackDuration val="00:00:00"/>
        '       <CurrentMediaDuration val="00:00:00"/>
        '       <CurrentTrackMetaData val=""/>
        '       <CurrentTrackURI val=""/>
        '       <AVTransportURI val=""/>
        '       <AVTransportURIMetaData val=""/>
        '       <NextAVTransportURI val=""/>
        '       <NextAVTransportURIMetaData val=""/>
        '       <CurrentTransportActions val=""/>
        '       <RecordStorageMedium val="NOT_IMPLEMENTED"/>
        '       <PossiblePlaybackStorageMedia val="NONE,NETWORK"/>
        '       <PossibleRecordStorageMedia val="NOT_IMPLEMENTED"/>
        '       <RecordMediumWriteStatus val="NOT_IMPLEMENTED"/>
        '       <CurrentRecordQualityMode val="NOT_IMPLEMENTED"/>
        '       <PossibleRecordQualityModes val="NOT_IMPLEMENTED"/>
        '   </InstanceID>
        '</Event>
        '
        ' Samsung
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/AVT/">
        '   <InstanceID val="0">
        '       <TransportState val="NO_MEDIA_PRESENT"/>
        '       <TransportStatus val="OK"/>
        '       <TransportPlaySpeed val="1"/>
        '       <NumberOfTracks val="0"/>
        '       <CurrentMediaDuration val="00:00:00"/>
        '       <AVTransportURI val=""/>
        '       <AVTransportURIMetaData val=""/>
        '       <PlaybackStorageMedium val="NETWORK,NONE"/>
        '       <CurrentTrack val="0"/>
        '       <CurrentTrackDuration val="00:00:00"/>
        '       <CurrentTrackMetaData val=""/>
        '       <CurrentTrackURI val=""/>
        '       <CurrentTransportActions val=""/>
        '       <NextAVTransportURI val="NOT_IMPLEMENTED"/>
        '       <NextAVTransportURIMetaData val="NOT_IMPLEMENTED"/>
        '       <RecordStorageMedium val="NOT_IMPLEMENTED"/>
        '       <RecordMediumWriteStatus val="NOT_IMPLEMENTED"/>
        '       <PossiblePlaybackStorageMedia val="NETWORK,NONE"/>
        '       <PossibleRecordStorageMedia val="NOT_IMPLEMENTED"/>
        '       <PossibleRecordQualityModes val="NOT_IMPLEMENTED"/>
        '       <CurrentPlayMode val="NORMAL"/>
        '       <CurrentRecordQualityMode val="NOT_IMPLEMENTED"/>
        '   </InstanceID>
        '</Event>
        Try
            Select Case VariableName
                Case "TransportState"
                    CurrentPlayerState = ConvertTransportStateToPlayerState(VariableValue.ToString)
                    'If MyTransportStateHasChanged Then MyCurrentTransportStatus = "" ' this to get rid of a lingering ERROR_OCCURRED state
                    MyCurrentTransportState = VariableValue.ToString
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (TransportState) = " & MyCurrentTransportState, LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "---->TransportStateChange for UPnPDevice = " & MyUPnPDeviceName & " has MyCurrentPlayerState = " & MyCurrentPlayerState.ToString & " and MyCurrentTransportState = " & MyCurrentTransportState.ToString)
                Case "TransportStatus"
                    'If VariableValue.ToString.ToUpper = "ERROR_OCCURRED" And MyCurrentTransportStatus.ToUpper <> "ERROR_OCCURRED" Then MyPlayerWentThroughPlayState = True
                    If VariableValue.ToString.ToUpper = "ERROR_OCCURRED" And Not MyTransportStateHasChanged Then MyPlayerWentThroughPlayState = True
                    MyCurrentTransportStatus = VariableValue.ToString
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (TransportStatus) = " & MyCurrentTransportStatus, LogType.LOG_TYPE_INFO)
                Case "TransportErrorDescription"
                    MyCurrentTransportErrorDescription = VariableValue.ToString
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (TransportErrorDescription) = " & MyCurrentTransportErrorDescription, LogType.LOG_TYPE_INFO)
                Case "TransportErrorURI"
                    MyCurrentTransportErrorURI = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (TransportErrorURI) = " & MyCurrentTransportErrorURI, LogType.LOG_TYPE_INFO)
                Case "TransportPlaySpeed"
                    If VariableValue.ToString.ToUpper = "NOT_IMPLEMENTED" Then Exit Select
                    If MyCurrentTransportPlaySpeed <> VariableValue Then
                        If VariableValue = "" Then
                            CurrentPlayerSpeed = 0
                        Else
                            CurrentPlayerSpeed = Val(VariableValue)
                        End If

                        If MyCurrentTransportPlaySpeed < 0 Then
                            CurrentPlayerState = player_state_values.Rewinding
                            MyTransportStateHasChanged = True
                        ElseIf MyCurrentTransportPlaySpeed > 1 Then
                            CurrentPlayerState = player_state_values.Forwarding
                            MyTransportStateHasChanged = True
                        Else
                            If MyCurrentPlayerState = player_state_values.Forwarding Or MyCurrentPlayerState = player_state_values.Rewinding Then
                                CurrentPlayerState = player_state_values.Playing
                                MyTransportStateHasChanged = True
                            End If
                        End If
                        ' go set a flag to GetPositionInfo because so fast forwarding is happening ..... or just stopped
                        'MyRefreshAVGetPositionInfo = True
                    End If
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (TransportPlaySpeed) = " & MyCurrentTransportPlaySpeed, LogType.LOG_TYPE_INFO)
                Case "NumberOfTracks"
                    MyCurrentNumberOfTracks = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (NumberOfTracks) = " & MyCurrentNumberOfTracks, LogType.LOG_TYPE_INFO)
                Case "CurrentMediaDuration"
                    MyCurrentMediaDuration = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentMediaDuration) = " & MyCurrentMediaDuration, LogType.LOG_TYPE_INFO)
                Case "AVTransportURI"
                    MyCurrentURI = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (AVTransportURI) = " & MyCurrentURI, LogType.LOG_TYPE_INFO)
                Case "AVTransportURIMetaData"
                    MyCurrentURIMetaData = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (AVTransportURIMetaData) = " & MyCurrentURIMetaData, LogType.LOG_TYPE_INFO)
                Case "PlaybackStorageMedium"
                    MyCurrentPlayMedium = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (PlaybackStorageMedium) = " & MyCurrentPlayMedium, LogType.LOG_TYPE_INFO)
                Case "CurrentTrack"
                    MyCurrentTrackNumber = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrack) = " & MyCurrentTrack, LogType.LOG_TYPE_INFO)
                Case "CurrentTrackDuration"
                    MyDurationInfoHasChanged = True
                    MyCurrentTrackDuration = VariableValue
                    If MyCurrentTrackDuration <> "NOT_IMPLEMENTED" Then SetTrackLength(GetSeconds(MyCurrentTrackDuration)) ' NOT_IMPLEMENTED
                    If CurrentTrackDuration = "" Or MyTrackLength = 0 Then SetPlayerPosition(0)
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrackDuration) = " & MyCurrentTrackDuration, LogType.LOG_TYPE_INFO)
                Case "CurrentTrackMetaData"
                    MyCurrentTrackMetaData = VariableValue
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrackMetaData) = " & MyCurrentTrackMetaData)
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrackMetaData) = " & MyCurrentTrackMetaData, LogType.LOG_TYPE_INFO)
                Case "CurrentTrackURI"
                    MyCurrentTrackURI = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrackURI) = " & MyCurrentTrackURI, LogType.LOG_TYPE_INFO)
                Case "CurrentTransportActions"
                    MyCurrentTransportActions = VariableValue
                    'If  piDebuglevel > DebugLevel.dlEvents  Then log( "ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTransportActions) = " & MyCurrentTransportActions)
                Case "NextAVTransportURI"
                    MyNextURI = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (NextAVTransportURI) = " & MyNextURI, LogType.LOG_TYPE_INFO)
                Case "NextAVTransportURIMetaData"
                    MyNextURIMetaData = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (NextAVTransportURIMetaData) = " & MyNextURIMetaData, LogType.LOG_TYPE_INFO)
                Case "RecordStorageMedium"
                    MyCurrentRecMedia = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (RecordStorageMedium) = " & MyCurrentRecMedia, LogType.LOG_TYPE_INFO)
                Case "RecordMediumWriteStatus"
                    MyCurrentWriteStatus = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (RecordMediumWriteStatus) = " & MyCurrentWriteStatus, LogType.LOG_TYPE_INFO)
                Case "PossiblePlaybackStorageMedia"
                    MyCurrentPossiblePlaybackStorageMedia = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (PossiblePlaybackStorageMedia) = " & MyCurrentPossiblePlaybackStorageMedia, LogType.LOG_TYPE_INFO)
                Case "PossibleRecordStorageMedia"
                    MyCurrentPossibleRecordStorageMedia = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (PossibleRecordStorageMedia) = " & MyCurrentPossibleRecordStorageMedia, LogType.LOG_TYPE_INFO)
                Case "CurrentPlayMode"
                    MyCurrentPlayMode = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentPlayMode) = " & MyCurrentPlayMode, LogType.LOG_TYPE_INFO)
                    'SetShuffleState(MyCurrentPlayMode) ' removing this because it hoses up my repeat behavior
                Case "CurrentRecordQualityMode"
                    MyCurrentRecQualityMode = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentRecordQualityMode) = " & MyCurrentRecQualityMode, LogType.LOG_TYPE_INFO)
                Case "PossibleRecordQualityModes"
                    MyPossibleRecordQualityModes = VariableValue
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (PossibleRecordQualityModes) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "CurrentCrossfadeMode"             ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentCrossfadeMode) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "CurrentSection"                   ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentSection) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:NextTrackURI"                   ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:NextTrackURI) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:NextTrackMetaData"              ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:NextTrackMetaData) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyNextTrackMetaData = VariableValue
                Case "r:EnqueuedTransportURI"           ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:EnqueuedTransportURI) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:EnqueuedTransportURIMetaData"   ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:EnqueuedTransportURIMetaData) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:SleepTimerGeneration"           ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:SleepTimerGeneration) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:AlarmRunning"                   ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:AlarmRunning) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:SnoozeRunning"                  ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:SnoozeRunning) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:RestartPending"                 ' Sonos
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (r:RestartPending) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "RelativeTimePosition"
                    MyDurationInfoHasChanged = True
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (RelativeTimePosition) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    If VariableValue.ToString.ToUpper <> "NOT_IMPLEMENTED" Then SetPlayerPosition(GetSeconds(VariableValue)) ' NOT_IMPLEMENTED
                Case "AbsoluteTimePosition"
                    MyDurationInfoHasChanged = True
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (AbsoluteTimePosition) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "RelativeCounterPosition"
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (RelativeCounterPosition) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "AbsoluteCounterPosition"
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (AbsoluteCounterPosition) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "CurrentTrackEmbeddedMetaData" ' from serviio
                    If PrintDebug Then Log("ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received (CurrentTrackEmbeddedMetaData) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case Else
                    If PrintDebug Then Log("Warning in ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " received untreated (" & VariableName & ") = " & VariableValue, LogType.LOG_TYPE_INFO)
            End Select

        Catch ex As Exception
            Log("Error in ProcessAVTransportXML for UPnPDevice = " & MyUPnPDeviceName & " with VariableName = " & VariableName & " and Value = " & VariableValue & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ProcessAVMetaData(VariableName As String, VariableValue As String, Attributes As System.Xml.XmlAttributeCollection, NextMetaData As Boolean, VariableXML As String, PrintDebug As Boolean)
        ' Buffy 
        '<item id="1$12$1013212268$2758058815" refID="1$268435466$1744833150" parentID="1$12$1013212268" restricted="1">
        '   <dc:title>Rubber Bullets</dc:title>
        '   <dc:creator>Unknown</dc:creator>
        '   <upnp:artist>10cc</upnp:artist>
        '   <upnp:author role="Composer">Graham Gouldman/Kevin Godley/Lol Creme</upnp:author>
        '   <upnp:album>Hits [Emi]</upnp:album>
        '   <upnp:genre>Rock</upnp:genre>
        '   <dc:date>2001-01-01</dc:date>
        '   <upnp:originalTrackNumber>3</upnp:originalTrackNumber>
        '   <upnp:albumArtURI>http://192.168.1.103:9050/cgi-bin/O1$12$1013212268$2758058815/W160/H160/S1/L1/Xjpeg-jpeg.desc.jpg</upnp:albumArtURI>
        '   <res size="5464748"  duration="0:05:18.000"  protocolInfo="http-get:*:audio/mpeg:DLNA.ORG_PN=MP3;DLNA.ORG_OP=11;DLNA.ORG_FLAGS=01700000000000000000000000000000" >http://192.168.1.103:9050/disk/music/DLNA-PNMP3-OP11-FLAGS01700000/O1$12$1013212268$2758058815.mp3</res>
        '   <upnp:class>object.item.audioItem.musicTrack</upnp:class>

        Try
            Select Case VariableName
                Case "dc:title"
                    If NextMetaData = False Then
                        MyTittleWasPresent = True
                        If MyCurrentTrack <> VariableValue Then
                            MyTrackInfoHasChanged = True
                            Track = VariableValue
                        End If
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Title) = " & MyCurrentTrack, LogType.LOG_TYPE_INFO)
                    Else
                        If MyNextTrack <> VariableValue Then
                            MyNextTrackInfoHasChanged = True
                            NextTrack = VariableValue
                        End If
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Title) = " & MyNextTrack, LogType.LOG_TYPE_INFO)
                    End If

                Case "dc:creator"
                    If NextMetaData = False Then
                        MyCurrentCreator = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Creator) = " & MyCurrentCreator, LogType.LOG_TYPE_INFO)
                    Else
                        MyNextCreator = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Creator) = " & MyNextCreator, LogType.LOG_TYPE_INFO)
                    End If

                Case "upnp:artist"
                    If Attributes IsNot Nothing Then
                        For Each attre As XmlAttribute In Attributes
                            If attre.Name.ToUpper = "ROLE" Then
                                If NewFullArtist <> "" Then NewFullArtist &= " " & vbCrLf
                                NewFullArtist &= attre.Value.ToString & " = " & VariableValue
                                If attre.Value.ToUpper = "PERFORMER" Then
                                    NewArtist = VariableValue
                                End If
                            End If
                        Next
                    Else
                        If NewFullArtist <> "" Then NewFullArtist &= " " & vbCrLf
                        NewFullArtist &= VariableValue
                        NewArtist = VariableValue
                    End If

                Case "upnp:author"
                    If Attributes IsNot Nothing Then
                        For Each attre As XmlAttribute In Attributes
                            If NewAuthor <> "" Then NewAuthor &= " " & vbCrLf
                            NewAuthor &= attre.Value.ToString & " = " & VariableValue
                        Next
                    Else
                        If NewAuthor <> "" Then NewAuthor &= " " & vbCrLf
                        NewAuthor &= VariableValue
                    End If

                Case "upnp:album"
                    If NextMetaData = False Then
                        If MyCurrentAlbum <> VariableValue Then
                            MyTrackInfoHasChanged = True
                            Album = VariableValue
                        End If
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Album) = " & MyCurrentAlbum, LogType.LOG_TYPE_INFO)
                    Else
                        If MyNextAlbum <> VariableValue Then
                            MyNextTrackInfoHasChanged = True
                            NextAlbum = VariableValue
                        End If
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Album) = " & MyNextAlbum, LogType.LOG_TYPE_INFO)
                    End If

                Case "upnp:genre"
                    If NextMetaData = False Then
                        Genre = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Genre) = " & MyCurrentGenre, LogType.LOG_TYPE_INFO)
                    Else
                        MyNextGenre = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Genre) = " & MyNextGenre, LogType.LOG_TYPE_INFO)
                    End If

                Case "dc:date"
                    If NextMetaData = False Then
                        MyCurrentTrackDate = VariableValue
                        ExtractDateInformation(MyCurrentTrackDate)
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (TrackDate) = " & MyCurrentTrackDate, LogType.LOG_TYPE_INFO)
                    Else
                        MyNextTrackDate = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next TrackDate) = " & MyNextTrackDate, LogType.LOG_TYPE_INFO)
                    End If

                Case "upnp:originalTrackNumber"
                    If NextMetaData = False Then
                        MyCurrentOriginalTrackNumber = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Original TrackNumber) = " & MyCurrentOriginalTrackNumber, LogType.LOG_TYPE_INFO)
                    Else
                        MyNextOriginalTrackNumber = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Original TrackNumber) = " & MyNextOriginalTrackNumber, LogType.LOG_TYPE_INFO)
                    End If

                Case "upnp:albumArtURI"
                    Dim attre As XmlAttribute ' we have attributes like dlna:profileID=JPEG_SM or JPEG_TN , not sure how to pick the largest one
                    For Each attre In Attributes
                        ' ohoh we may have multiple URIs , I would like to pick out the largest picture (TN, SM, MED, LRG)
                        If attre.Name.ToUpper = "DLNA:PROFILEID" Then
                            If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Album Art URI) = " & VariableValue & " with attribute = " & attre.Name & " and Value = " & attre.Value, LogType.LOG_TYPE_INFO)
                            Select Case attre.Value.ToUpper
                                Case "JPEG_TN"
                                    If AlbumArtSizeFound <> "" Then
                                        ' we already have another size which by definition must be larger than tiny
                                        Exit For
                                    End If
                                    AlbumArtURIWasFound = True
                                    AlbumArtSizeFound = "JPEG_TN"
                                    AlbumArtURIFound = VariableValue
                                Case "JPEG_SM"
                                    If AlbumArtSizeFound = "JPEG_MED" Or AlbumArtSizeFound = "JPEG_LRG" Then
                                        ' we already have another size which is larger
                                        Exit For
                                    End If
                                    AlbumArtURIWasFound = True
                                    AlbumArtSizeFound = "JPEG_SM"
                                    AlbumArtURIFound = VariableValue
                                Case "JPEG_MED"
                                    If AlbumArtSizeFound = "JPEG_LRG" Then
                                        ' we already have another size which is larger
                                        Exit For
                                    End If
                                    AlbumArtURIWasFound = True
                                    AlbumArtSizeFound = "JPEG_MED"
                                    AlbumArtURIFound = VariableValue
                                Case "JPEG_LRG"
                                    ' this is largest
                                    AlbumArtURIWasFound = True
                                    AlbumArtSizeFound = "JPEG_LRG"
                                    AlbumArtURIFound = VariableValue
                                Case Else
                            End Select
                            Exit For
                        End If
                    Next
                    If Not AlbumArtURIWasFound Then
                        If NextMetaData = False Then
                            If MyCurrentArtworkURL <> VariableValue Then
                                MyAlbumArtURIHasChanged = True
                                ArtworkURL = GetAlbumArtPath(VariableValue, False)
                            End If
                            If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Album Art URI) = " & MyCurrentArtworkURL, LogType.LOG_TYPE_INFO)
                        Else
                            If MyNextArtworkURL <> VariableValue Then
                                MyNextAlbumArtURIHasChanged = True
                                NextArtworkURL = GetAlbumArtPath(VariableValue, True)
                            End If
                            If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Album Art URI) = " & MyNextArtworkURL, LogType.LOG_TYPE_INFO)
                        End If
                    End If

                Case "res"
                    Dim xmlData As XmlDocument = New XmlDocument
                    Dim Duration As String = ""
                    Dim Size As String = ""
                    Dim SampleFrequency As String = ""
                    Dim bitsPerSample As String = ""
                    Dim nrAudioChannels As String = ""
                    Dim bitrate As String = ""
                    Try
                        xmlData.LoadXml(VariableXML)
                    Catch ex As Exception
                        Log("Error in ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " analyzing res XML with error = " & ex.Message & " and XML = " & VariableXML, LogType.LOG_TYPE_INFO)
                        xmlData = Nothing
                    End Try
                    If Not xmlData Is Nothing Then
                        Try
                            Duration = xmlData.GetElementsByTagName("res").Item(0).Attributes("duration").Value
                            If MyTrackLength = 0 And Duration <> "" And NextMetaData = False Then
                                MyCurrentTrackDuration = Duration
                                'SetTrackLength(GetSeconds(Duration))
                            ElseIf Duration <> "" Then
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Duration) = " & Duration & " and CurrentrackLength = " & MyTrackLength.ToString)
                            End If
                        Catch ex As Exception
                        End Try
                        Try
                            Size = xmlData.GetElementsByTagName("res").Item(0).Attributes("size").Value
                        Catch ex As Exception
                        End Try
                        Try
                            MyCurrentResolution = xmlData.GetElementsByTagName("res").Item(0).Attributes("resolution").Value
                        Catch ex As Exception
                            MyCurrentResolution = ""
                        End Try
                        Try
                            SampleFrequency = xmlData.GetElementsByTagName("res").Item(0).Attributes("sampleFrequency").Value
                        Catch ex As Exception
                        End Try
                        Try
                            bitsPerSample = xmlData.GetElementsByTagName("res").Item(0).Attributes("bitsPerSample").Value
                        Catch ex As Exception
                        End Try
                        Try
                            nrAudioChannels = xmlData.GetElementsByTagName("res").Item(0).Attributes("nrAudioChannels").Value
                        Catch ex As Exception
                        End Try
                        Try
                            bitrate = xmlData.GetElementsByTagName("res").Item(0).Attributes("bitrate").Value
                        Catch ex As Exception
                        End Try
                    End If
                    Dim TempInfo As String = ""
                    If NextMetaData = False Then
                        TempInfo = VariableValue & " Duration: " & Duration & " size: " & Size & " Resolution: " & MyCurrentResolution & " BitsPerSample: " & bitsPerSample
                        MyCurrentURI = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Res) = " & TempInfo, LogType.LOG_TYPE_INFO)
                    Else
                        TempInfo = VariableValue & " Duration: " & Duration & " size: " & Size & " Resolution: " & MyCurrentResolution & " BitsPerSample: " & bitsPerSample
                        MyNextURI = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Res) = " & TempInfo, LogType.LOG_TYPE_INFO)
                    End If
                    xmlData = Nothing
                Case "upnp:class"
                    If NextMetaData = False Then
                        MyCurrentClass = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Class) = " & MyCurrentClass, LogType.LOG_TYPE_INFO)
                    Else
                        MyNextClass = VariableValue
                        If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (Next Class) = " & MyNextClass, LogType.LOG_TYPE_INFO)
                    End If
                Case "r:albumArtist"                    ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:albumArtist) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:NextTrackURI"                   ' Sonos
                    MyNextURI = VariableValue
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:NextTrackURI) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:NextTrackMetaData"              ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:NextTrackMetaData) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:EnqueuedTransportURI"           ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:EnqueuedTransportURI) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:EnqueuedTransportURIMetaData"   ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:EnqueuedTransportURIMetaData) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:SleepTimerGeneration"           ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:SleepTimerGeneration) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:AlarmRunning"                   ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:AlarmRunning) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:SnoozeRunning"                  ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:SnoozeRunning) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:RestartPending"                 ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:RestartPending) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:radioShowMd"                 ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:radioShowMd) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "r:streamContent"                 ' Sonos
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (r:streamContent) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:album_art" ' serviio
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:album_art) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    ' realy not sure what this is
                Case "upnp:icon" ' serviio
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:icon) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "dc:description" ' serviio & Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (dc:description) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    If Attributes IsNot Nothing Then
                        For Each attre As XmlAttribute In Attributes
                            If attre.Name.ToUpper = "ID" Then
                                If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                                NewDesc &= attre.Value.ToString & " = " & VariableValue
                            End If
                        Next
                    Else
                        If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                        NewDesc &= VariableValue
                    End If
                Case "dc:rights" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (dc:rights) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:rating" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:rating) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyCurrentTrackRating = VariableValue
                Case "upnp:actor" ' Windows Media Player
                    If NewActor <> "" Then NewActor &= " " & vbCrLf
                    NewActor &= VariableValue
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:actor) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:userAnnotation" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:userAnnotation) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "desc" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (desc) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    If Attributes IsNot Nothing Then
                        For Each attre As XmlAttribute In Attributes
                            If attre.Name.ToUpper = "ID" Then
                                If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                                NewDesc &= attre.Value.ToString & " = " & VariableValue
                            End If
                        Next
                    Else
                        If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                        NewDesc &= VariableValue
                    End If
                Case "upnp:toc" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:toc) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:scheduledStartTime" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:scheduledStartTime) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "dc:publisher" ' Windows Media Player
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (dc:publisher) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:originalDiscNumber" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:originalDiscNumber) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:originalDiscCount" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:originalDiscCount) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:albumArtist" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:albumArtist) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:extension" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:extension) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:lastPlayedTime" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:lastPlayedTime) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyCurrentTrackLastPlayedDate = VariableValue
                Case "pv:playcount" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:playcount) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyCurrentTrackPlayedCount = Val(VariableValue)
                Case "pv:modificationTime" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:modificationTime) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:bookmark" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:bookmark) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:genre_crosslink" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:genre_crosslink) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:artist_crosslink" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:artist_crosslink) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:album_crosslink" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:album_crosslink) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:addedTime" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:addedTime) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:lastUpdated" ' Sony 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:lastUpdated) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:orientation" ' BubbleUPnP 
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:orientation) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:lastPlaybackTime" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:lastPlaybackTime) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyCurrentTrackLastPlayedDate = VariableValue
                Case "upnp:lastPlaybackPosition" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:lastPlaybackPosition) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    'MyCurrentTrackLastPlayedDate = VariableValue
                Case "upnp:playbackCount" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:playbackCount) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    MyCurrentTrackPlayedCount = Val(VariableValue)
                Case "upnp:director" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:director) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:longDescription" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:longDescription) = " & VariableValue, LogType.LOG_TYPE_INFO)
                    If Attributes IsNot Nothing Then
                        For Each attre As XmlAttribute In Attributes
                            If attre.Name.ToUpper = "ID" Then
                                If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                                NewDesc &= attre.Value.ToString & " = " & VariableValue
                            End If
                        Next
                    Else
                        If NewDesc <> "" Then NewDesc &= " " & vbCrLf
                        NewDesc &= VariableValue
                    End If
                Case "upnp:programTitle" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:programTitle) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "upnp:episodeNumber" ' XBMC
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (upnp:episodeNumber) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "pv:rating" ' JRiver
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (pv:rating) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case "av:liveType" ' Sony AV Receiver
                    If PrintDebug Then Log("ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received (av:liveType) = " & VariableValue, LogType.LOG_TYPE_INFO)
                Case Else
                    If PrintDebug Then Log("Warning in ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " received untreated (" & VariableName & ") = " & VariableValue, LogType.LOG_TYPE_WARNING)
            End Select

        Catch ex As Exception
            Log("Error in ProcessAVMetaData for UPnPDevice = " & MyUPnPDeviceName & " with VariableName = " & VariableName & " and Value = " & VariableValue & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function ConvertTransportStateToPlayerState(TransportState As String) As player_state_values
        '<allowedValue>STOPPED</allowedValue>
        '<allowedValue>PAUSED_PLAYBACK</allowedValue>
        '<allowedValue>PAUSED_RECORDING</allowedValue>
        '<allowedValue>PLAYING</allowedValue>
        '<allowedValue>RECORDING</allowedValue>
        '<allowedValue>TRANSITIONING</allowedValue>
        '<allowedValue>NO_MEDIA_PRESENT</allowedValue>
        ConvertTransportStateToPlayerState = player_state_values.Stopped
        If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
        Select Case UCase(TransportState)
            Case "STOPPED"
                If UCase(MyCurrentTransportState) <> UCase(TransportState) Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
                    MyTransportStateHasChanged = True
                End If
                ConvertTransportStateToPlayerState = player_state_values.Stopped
            Case "PAUSE_PLAYBACK", "PAUSED_RECORDING", "PAUSED_PLAYBACK"
                If UCase(MyCurrentTransportState) <> UCase(TransportState) Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
                    MyTransportStateHasChanged = True
                End If
                ConvertTransportStateToPlayerState = player_state_values.Paused
            Case "PLAYING", "RECORDING"
                MyPlayerWentThroughPlayState = True
                MyTimeIveBeenWaitingForPlayState = 0
                If UCase(MyCurrentTransportState) <> UCase(TransportState) Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
                    MyTransportStateHasChanged = True
                End If
                If strApplySeek <> "" Then
                    AVTSeek(HSPI.AVT_SeekMode.REL_TIME, strApplySeek)
                    strApplySeek = ""
                End If

                ConvertTransportStateToPlayerState = player_state_values.Playing
            Case "TRANSITIONING"
                If UCase(MyCurrentTransportState) <> UCase(TransportState) Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
                    MyTransportStateHasChanged = True
                End If
                ConvertTransportStateToPlayerState = player_state_values.Transitioning
            Case "NO_MEDIA_PRESENT"
                If UCase(MyCurrentTransportState) <> UCase(TransportState) Then
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState called with CurrentState = " & MyCurrentTransportState.ToString & " and transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
                    MyTransportStateHasChanged = True
                End If
                ConvertTransportStateToPlayerState = player_state_values.Stopped
            Case Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ConvertTransportStateToPlayerState found unkown transportstatechange = " & TransportState.ToString, LogType.LOG_TYPE_INFO)
        End Select
    End Function

    Public Sub SetPlayState(PlayState As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetPlayState called for UPnPDevice = " & MyUPnPDeviceName & " and state = " & PlayState & " while PlayFromQueue = " & MyPlayFromQueue.ToString & " and Current PlayerState = " & MyCurrentPlayerState.ToString & " and QueueState = " & MyQueuePlayState.ToString, LogType.LOG_TYPE_INFO)
        Select Case PlayState
            Case "Play"
                If MyPlayFromQueue Then
                    If MyQueuePlayState = player_state_values.Paused Or MyCurrentPlayerState = player_state_values.Forwarding Or MyCurrentPlayerState = player_state_values.Rewinding Then
                        CurrentPlayerState = player_state_values.Playing
                        MyQueuePlayState = player_state_values.Playing
                        MyTransportStateHasChanged = True
                        MyTrackInfoHasChanged = False
                        UpdateTransportState()
                    ElseIf MyCurrentPlayerState = player_state_values.Playing Then
                        CurrentPlayerState = player_state_values.Paused
                        MyQueuePlayState = player_state_values.Paused
                        MyTransportStateHasChanged = True
                        MyTrackInfoHasChanged = False
                        UpdateTransportState()
                        If MyUPnPDeviceServiceType = "HST" Then Exit Sub
                        AVTPause()
                        Exit Sub
                    Else
                        'MyCurrentPlayerState = player_state_values.Playing
                        MyQueuePlayState = player_state_values.Playing
                        MyTransportStateHasChanged = True
                        MyTrackInfoHasChanged = False
                        PlayFromQueue(1)
                        UpdateTransportState()
                        Exit Sub
                    End If
                End If
                AVTPlay()
            Case "Stop"
                MyQueuePlayState = player_state_values.Stopped
                MyCurrentQueueIndex = 1
                MyNbrOfQueueItemsPlayed = 0
                If MyPlayFromQueue Then
                    CurrentPlayerState = player_state_values.Stopped
                    MyTransportStateHasChanged = True
                    MyTrackInfoHasChanged = False
                    UpdateTransportState()
                    If MyUPnPDeviceServiceType = "HST" Then Exit Sub
                End If
                AVTStop()
            Case "Pause"
                If MyPlayFromQueue Then
                    If MyCurrentPlayerState = player_state_values.Paused Then
                        CurrentPlayerState = player_state_values.Playing
                        MyQueuePlayState = player_state_values.Playing
                        MyTransportStateHasChanged = True
                        MyTrackInfoHasChanged = False
                        UpdateTransportState()
                        If MyUPnPDeviceServiceType = "HST" Then Exit Sub
                        AVTPlay()
                        Exit Sub
                    End If
                    If MyQueuePlayState <> player_state_values.Stopped Then
                        CurrentPlayerState = player_state_values.Paused
                        MyQueuePlayState = player_state_values.Paused
                    End If
                    MyTransportStateHasChanged = True
                    MyTrackInfoHasChanged = False
                    UpdateTransportState()
                    If MyUPnPDeviceServiceType = "HST" Then Exit Sub
                End If
                AVTPause()
            Case "Next"
                If MyPlayFromQueue Then
                    NextQueue()
                Else
                    AVTNext()
                End If
            Case "Prev"
                If MyPlayFromQueue Then
                    PreviousQueue()
                Else
                    AVTPrevious()
                End If
        End Select
    End Sub

    Public Function GetTransportState() As String
        GetTransportState = ""
        If DeviceStatus = "Offline" Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(2)
            InArg(0) = 0
            AVTransport.InvokeAction("GetTransportInfo", InArg, OutArg)
            MyCurrentTransportState = OutArg(0)
            If OutArg(0) = "PLAYING" Then
                GetTransportState = "Play"
                CurrentPlayerState = player_state_values.Playing
            ElseIf OutArg(0) = "STOPPED" Then
                GetTransportState = "Stop"
                CurrentPlayerState = player_state_values.Stopped
            ElseIf OutArg(0) = "PAUSED_PLAYBACK" Then
                GetTransportState = "Pause"
                CurrentPlayerState = player_state_values.Paused
            Else
                GetTransportState = "Unknown"
                CurrentPlayerState = player_state_values.Stopped
            End If
        Catch ex As Exception
            Log("ERROR in GetTransportState for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub UpdatePositionInfo()
        If MyCurrentPlayerState = player_state_values.Playing Or MyCurrentPlayerState = player_state_values.Forwarding Or MyCurrentPlayerState = player_state_values.Rewinding Then
            ' OK I'm going to fake it here a bit
            MyCurrentPlayerPosition = MyCurrentPlayerPosition + MyCurrentTransportPlaySpeed
            If CurrentContentClass = UPnPClassType.ctMusic Or CurrentContentClass = UPnPClassType.ctVideo Then
                If MyCurrentPlayerPosition > (MyTrackLength + 5) Then ' add a 5 seconds buffer to it
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdatePositionInfo for UPnPDevice = " & MyUPnPDeviceName & " has exceeded TrackDuraction. CurrentPosition = " & MyCurrentPlayerPosition.ToString & " and Track Duration = " & MyTrackLength.ToString, LogType.LOG_TYPE_WARNING)
                    GetPositionDurationInfoOnly(0)
                    If MyTrackInfoHasChanged Then MyPlayerWentThroughTrackChange = True
                End If
            End If
        ElseIf MyCurrentPlayerState = player_state_values.Stopped Then
            'MyCurrentPlayerPosition = 0
        End If
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdatePositionInfo called for device = " & MyUPnPDeviceName & " and value is now = " & MyCurrentPlayerPosition.ToString, LogType.LOG_TYPE_INFO)
        If HSRefTrackPos <> -1 Then
            Select Case MyHSTrackPositionFormat
                Case HSSTrackPositionSettings.TPSSeconds
                    hs.SetDeviceString(HSRefTrackPos, MyCurrentPlayerPosition.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, MyCurrentPlayerPosition, True)
                Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                    hs.SetDeviceString(HSRefTrackPos, ConvertSecondsToTimeFormat(MyCurrentPlayerPosition), True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, MyCurrentPlayerPosition, True)
                Case HSSTrackPositionSettings.TPSPercentage
                    Dim TrackPos As Integer = 0
                    If MyTrackLength <> 0 Then
                        TrackPos = MyCurrentPlayerPosition / MyTrackLength * 100
                    End If
                    hs.SetDeviceString(HSRefTrackPos, TrackPos.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, TrackPos, True)
            End Select
        End If
    End Sub


    Public Property PlayerPosition As Integer 'As Integer 'Implements MediaCommon.MusicAPI.PlayerPosition
        Get ' The position of the player in the current track expressed as seconds.
            ' the HST plugin calls this a lot, I think it uses it to figure out when tracks have changed
            'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then log( "Get PlayerPosition called for device - " & MyUPnPDeviceName & " and position = " & MyCurrentPlayerPosition.ToString)
            PlayerPosition = MyCurrentPlayerPosition ' this is automatically updated by the own timer
        End Get
        Set(ByVal value As Integer) ' Sets the position of the player in the current track - parameter is expressed as seconds.         
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Set PlayerPosition called for UPnPDevice - " & MyUPnPDeviceName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Dim Time As String
            Time = ConvertSecondsToTimeFormat(value)
            AVTSeek(HSPI.AVT_SeekMode.REL_TIME, Time)
        End Set
    End Property

    Public Sub SetPlayerPosition(ByVal Position As Integer)
        ' I call this in response to updates received from the controllers. I cannot call "PlayerPosition because then I create a loop, instruction SONOS to go 
        ' to where it just reported it was
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetPlayerPosition called for device - " & MyUPnPDeviceName & " with Position = " & Position.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SetPlayerPosition called for device - " & MyUPnPDeviceName & " with Position = " & Position.ToString, LogType.LOG_TYPE_INFO)
        Try
            If MyCurrentPlayerPosition <> Position And HSRefTrackPos <> -1 Then
                Select Case MyHSTrackPositionFormat
                    Case HSSTrackPositionSettings.TPSSeconds
                        hs.SetDeviceString(HSRefTrackPos, Position.ToString, True)
                        hs.SetDeviceValueByRef(HSRefTrackPos, CType(Position, Double), True)
                    Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                        hs.SetDeviceString(HSRefTrackPos, ConvertSecondsToTimeFormat(Position), True)
                        hs.SetDeviceValueByRef(HSRefTrackPos, CType(Position, Double), True)
                    Case HSSTrackPositionSettings.TPSPercentage
                        Dim TrackPos As Integer = 0
                        If MyTrackLength <> 0 Then
                            TrackPos = Int(MyCurrentPlayerPosition / MyTrackLength) * 100
                        End If
                        hs.SetDeviceString(HSRefTrackPos, TrackPos.ToString, True)
                        hs.SetDeviceValueByRef(HSRefTrackPos, CType(TrackPos, Double), True)
                End Select
            End If
        Catch ex As Exception
            Log("Error in SetPlayerPosition called for device - " & MyUPnPDeviceName & " with Position = " & Position.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyCurrentPlayerPosition = Position
    End Sub

    Public Sub DoFastForward()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoFastForward called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If MyCurrentPlayerState = player_state_values.Stopped Then Exit Sub
        If MyPossibleFFSpeeds Is Nothing Then Exit Sub
        For index = 0 To UBound(MyPossibleFFSpeeds)
            If MyPossibleFFSpeeds(index) > MyCurrentTransportPlaySpeed Then
                CurrentPlayerState = player_state_values.Forwarding
                AVTPlay(MyPossibleFFSpeeds(index))
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoFastForward called for UPnPDevice = " & MyUPnPDeviceName & " and speed set to = " & MyPossibleFFSpeeds(index).ToString, LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
        Next
        ' no faster speed was found go back to 1
        'AVTPlay(1)
        'AVTPlay(MyPossibleFFSpeeds(UBound(MyPossibleFFSpeeds)))
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoFastForward called for UPnPDevice = " & MyUPnPDeviceName & " and speed set to = 1", LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub DoRewind()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRewind called for UPnPDevice = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If MyCurrentPlayerState = player_state_values.Stopped Then Exit Sub
        If MyPossibleREWSpeeds Is Nothing Then Exit Sub
        For index = 0 To UBound(MyPossibleREWSpeeds)
            If MyPossibleREWSpeeds(index) < MyCurrentTransportPlaySpeed Then
                CurrentPlayerState = player_state_values.Rewinding
                AVTPlay(MyPossibleREWSpeeds(index))
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRewind called for UPnPDevice = " & MyUPnPDeviceName & " and speed set to = " & MyPossibleREWSpeeds(index).ToString, LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
        Next
        ' no slower speed was found go back to 1
        AVTPlay(1)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRewind called for UPnPDevice = " & MyUPnPDeviceName & " and speed set to = 1", LogType.LOG_TYPE_INFO)
    End Sub



    Public Sub PlayDBItem(ObjectID As String, ServerDeviceName As String, CurrentItem As Boolean, Optional StartPlayFlag As Boolean = True)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem called for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & " and Server = " & ServerDeviceName & " and CurrentItem = " & CurrentItem.ToString & " and StartPlayFlag = " & StartPlayFlag.ToString, LogType.LOG_TYPE_INFO)
        Dim ServerApi As HSPI
        ServerApi = MyReferenceToMyController.GetAPIByUDN(ServerDeviceName)
        If ServerApi Is Nothing Then
            Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & ". Unable to get ServerApi for devicename = " & ServerDeviceName, LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem for device - " & MyUPnPDeviceName & " retrieved DeviceAPI for device = " & ServerApi.DeviceName, LogType.LOG_TYPE_INFO)
        Dim OutArg(3)
        OutArg = ServerApi.Browse(ObjectID, "BrowseMetadata", "*", 0, 1, "")
        If OutArg Is Nothing Then
            Log("Error in PlayDBItem for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse from server device = " & ServerDeviceName & " returned NOK", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        Dim BrowseResult As String = OutArg(0)
        Dim BrowseNumberReturned As Integer = OutArg(1)
        If BrowseNumberReturned < 1 Then
            Log("Error in PlayDBItem for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned = " & BrowseNumberReturned.ToString & " records", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        BrowseResult = Trim(BrowseResult)
        If BrowseResult = "" Then
            Log("Error in PlayDBItem for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned empty XML", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If MyReplaceLoopbackIPAddress And PlugInIPAddress <> "" Then
            Dim builder As New System.Text.StringBuilder(BrowseResult.ToString)
            builder.Replace(LoopBackIPv4Address, PlugInIPAddress)
            BrowseResult = builder.ToString
            If PIDebuglevel > DebugLevel.dlEvents Then Log("PlayDBItem changed the meta data for Loopback IP to = " & BrowseResult.ToString, LogType.LOG_TYPE_INFO)
            builder = Nothing
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then PrintOutXML(BrowseResult)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then PrintOutXML(BrowseResult)
        Dim ItemMetadata As String = BrowseResult
        Dim xmlData As XmlDocument = New XmlDocument

        OutArg = Nothing
        Try
            xmlData.LoadXml(BrowseResult.ToString)
        Catch ex As Exception
            Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & BrowseResult.ToString, LogType.LOG_TYPE_INFO)
            Exit Sub
        End Try

        Dim ItemClassInfo As String = ""
        Dim ItemClassInfos As String() = Nothing
        Try
            ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
            ItemClassInfos = Split(ItemClassInfo, ".")
            If UCase(ItemClassInfos(1)) <> "ITEM" Then
                ' this should not be
                Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & ". The Object is a container, it is = " & ItemClassInfo.ToString, LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for upnp:class. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_INFO)
        End Try

        Dim ItemURL As String = ""

        ItemURL = FindRightSizePictureURL(BrowseResult.ToString, MyPictureSize)
        '    <res colorDepth="24" protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=00;DLNA.ORG_CI=0;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="2048x1536" size="1324894">http://192.168.1.123:8895/resource/122/MEDIA_ITEM/JPEG_LRG*0</res>
        'If ItemURL = "" Then
        'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & ". URL is empty. XML = " & ItemMetadata.ToString)
        'Exit Sub
        'End If
        If Not MyUPnPDeviceServiceType = "HST" Then
            Dim AnythingHasChanged As Boolean = False
            Dim WeKeptatLeastOne As Boolean = False
            Try
                Dim LoopCount As Integer = 0
                Do
                    Dim ProtocolInfo As String = xmlData.GetElementsByTagName("res").Item(LoopCount).Attributes("protocolInfo").Value
                    If Not CheckForProtocol(False, ProtocolInfo) Then
                        ' the renderer cannot handle this protocol
                        Dim ResNode As XmlNode
                        ResNode = xmlData.GetElementsByTagName("res").Item(LoopCount)
                        If Not ResNode Is Nothing Then ResNode.ParentNode.RemoveChild(ResNode)
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("PlayDBItem for device = " & MyUPnPDeviceName & " removed RES = " & ProtocolInfo, LogType.LOG_TYPE_INFO)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "PlayDBItem for device = " & MyUPnPDeviceName & " removed RES = " & ProtocolInfo)
                        AnythingHasChanged = True
                    Else
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("PlayDBItem for device = " & MyUPnPDeviceName & " kept RES = " & ProtocolInfo, LogType.LOG_TYPE_INFO)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "PlayDBItem for device = " & MyUPnPDeviceName & " kept RES = " & ProtocolInfo)
                        WeKeptatLeastOne = True
                        LoopCount = LoopCount + 1
                    End If

                    If LoopCount > 100 Then Exit Do ' should never be. When no more res items found it will cause exception
                Loop
            Catch ex As Exception
            End Try
            If AnythingHasChanged And WeKeptatLeastOne Then
                ItemMetadata = xmlData.InnerXml
                If PIDebuglevel > DebugLevel.dlEvents Then Log("PlayDBItem for device = " & MyUPnPDeviceName & " changed Metadata. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_INFO)
            End If

        End If

        ' pick up the album art in case the player is not capable of returning it
        If CurrentItem Then
            If MyUPnPDeviceServiceType = "HST" Then
                ArtworkURL = GetAlbumArtPath(ItemURL, False)
            Else
                Try
                    Dim AlbumArtURL As String = ""
                    AlbumArtURL = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                    If AlbumArtURL <> "" Then
                        If MyCurrentArtworkURL <> AlbumArtURL Then
                            MyAlbumArtURIHasChanged = True
                            ArtworkURL = GetAlbumArtPath(AlbumArtURL, False)
                        End If
                    End If
                Catch ex As Exception
                    ' no big deal, there is no album art
                    MyAlbumArtURIHasChanged = True
                    If UCase(ItemClassInfos(2)) <> "IMAGEITEM" Then
                        ArtworkURL = GetAlbumArtPath("", False)
                    Else
                        ArtworkURL = GetAlbumArtPath(ItemURL, False)
                    End If
                End Try
            End If
        End If

        ' OK we should have a URL and Metadata
        Dim ItemName As String = ""
        Try
            ItemName = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem for device = " & MyUPnPDeviceName & " found itemName = " & ItemName, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
        End Try

        Dim ItemAlbum As String = ""
        Try
            ItemAlbum = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem for device = " & MyUPnPDeviceName & " found ItemAlbum = " & ItemAlbum, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
        End Try

        ItemMetadata = CheckNameSpaceXMLData(ItemMetadata)
        ItemMetadata = CheckForXMLIssues(ItemMetadata)

        If MyUPnPDeviceServiceType = "HST" Then
            ' we're going to have to copy the picture here but for time being let's only set playstate
            ' Use ItemURL to copy the picture
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem for HST device = " & MyUPnPDeviceName & " found Res URL = " & ItemURL, LogType.LOG_TYPE_INFO)
            Try
                If UCase(ItemClassInfos(2)) <> "IMAGEITEM" Then
                    ' this should not be
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in PlayDBItem for device = " & MyUPnPDeviceName & ". The items is not a picture, it is = " & ItemClassInfo.ToString, LogType.LOG_TYPE_INFO)
                    Exit Sub
                End If
            Catch ex As Exception
            End Try
            ' must be object.item.imageItem.photo to allow it to work
            ' I could use (protocolInfo) = http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_TN;DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=00D00000000000000000000000000000
            ' or (resolution) = 640x480
            ' or (colorDepth) = 24
            '  <item id="I_Y^YEAR_2005^I_122" parentID="I_Y^YEAR_2005" restricted="1">
            '    <dc:title>DSC02182</dc:title>
            '    <upnp:class>object.item.imageItem.photo</upnp:class>
            '    <upnp:albumArtURI dlna:profileID="JPEG_TN">http://192.168.1.123:8895/resource/122/COVER_IMAGE</upnp:albumArtURI>
            '    <upnp:icon>http://192.168.1.123:8895/resource/122/COVER_IMAGE</upnp:icon>
            '    <upnp:album>Belgium summer 2005</upnp:album>
            '    <dc:date>2005-07-12</dc:date>
            '    <res colorDepth="24" protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=00;DLNA.ORG_CI=0;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="2048x1536" size="1324894">http://192.168.1.123:8895/resource/122/MEDIA_ITEM/JPEG_LRG*0</res>
            '    <res protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_TN;DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="160x120">http://192.168.1.123:8895/resource/122/COVER_IMAGE</res>
            '    <res colorDepth="24" protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_SM;DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="640x480">http://192.168.1.123:8895/resource/122/MEDIA_ITEM/JPEG_SM*0</res>
            '    <res colorDepth="24" protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_MED;DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="1024x768">http://192.168.1.123:8895/resource/122/MEDIA_ITEM/JPEG_MED*0</res>
            '  </item>

            'MyCurrentArtworkURL = GetAlbumArtPath(ItemURL, False)
            CurrentPlayerState = player_state_values.Playing
            SetPlayerPosition(0)
            Track = ItemName
            Album = ItemAlbum
            Artist = "Pictures"
            MyTransportStateHasChanged = True
            MyTrackInfoHasChanged = True
            UpdateTransportState()
            Exit Sub
        End If

        ' Non HST devices
        ' If not Pictures then we still need to pick up the res URL else it is already stored in ItemURL

        Try
            If UCase(ItemClassInfos(2)) <> "IMAGEITEM" And ItemMetadata <> "" Then
                ItemURL = ""
                xmlData.LoadXml(ItemMetadata)
                ' we pick up the first RES, this may have to change in the future or make it selectable like the picture size
                ItemURL = xmlData.GetElementsByTagName("res").Item(0).InnerText
            End If
        Catch ex As Exception
            Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & "  retrieving RES from XML with Error = " & ex.Message & ". XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_INFO)
        End Try


        If CurrentItem Then
            'If MyPollForTransportChangeFlag Then
            'AVTGetTransportInfo()
            'If MyCurrentPlayerState <> ConvertTransportStateToPlayerState(MyCurrentTransportState) Then
            'MyCurrentPlayerState = ConvertTransportStateToPlayerState(MyCurrentTransportState)
            'AVTGetPositionInfo()
            'End If
            'End If
            If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" And MyIssueStopInBetween Then
                AVTStop(0)
            End If
            Try
                If AVTSetAVTransportURI(ItemURL, ItemMetadata).ToUpper <> "OK" Then
                    MyPlayerWentThroughPlayState = True
                    MyCurrentQueueObjectID = ObjectID
                    MyCurrentQueueServerUDN = ServerDeviceName
                    Exit Sub
                End If
                wait(1)
                MyCurrentQueueObjectID = ObjectID
                MyCurrentQueueServerUDN = ServerDeviceName
            Catch ex As Exception
                Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & " setting AVTransport with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            Try
                AVTPlay()
                If Not StartPlayFlag Then AVTPause()
            Catch ex As Exception
                Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & " setting AVPlay with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        Else
            Try
                AVTSetNextAVTransportURI(ItemURL, ItemMetadata)
                MyNextQueueObjectID = ObjectID
                MyNextQueueServerUDN = ServerDeviceName
            Catch ex As Exception
                Log("Error in PlayDBItem for device = " & MyUPnPDeviceName & " setting AVNextTransport with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        End If
    End Sub

    Public Sub GetNextPicture(ObjectID As String, ServerDeviceName As String)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextPicture called for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & " and Server = " & ServerDeviceName, LogType.LOG_TYPE_INFO)
        Dim ServerApi As HSPI
        ServerApi = MyReferenceToMyController.GetAPIByUDN(ServerDeviceName)
        If ServerApi Is Nothing Then
            Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & ". Unable to get ServerApi for devicename = " & ServerDeviceName, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextPicture for device - " & MyUPnPDeviceName & " retrieved DeviceAPI for device = " & ServerApi.DeviceName, LogType.LOG_TYPE_INFO)
        Dim OutArg(3)
        OutArg = ServerApi.Browse(ObjectID, "BrowseMetadata", "*", 0, 1, "")
        If OutArg Is Nothing Then
            Log("Error in GetNextPicture for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse from server device = " & ServerDeviceName & " returned NOK", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Dim BrowseResult As String = OutArg(0)
        Dim BrowseNumberReturned As Integer = OutArg(1)
        If BrowseNumberReturned < 1 Then
            Log("Error in GetNextPicture for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned = " & BrowseNumberReturned.ToString & " records", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        BrowseResult = Trim(BrowseResult)
        If BrowseResult = "" Then
            Log("Error in GetNextPicture for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned empty XML", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If MyReplaceLoopbackIPAddress And PlugInIPAddress <> "" Then
            Dim builder As New System.Text.StringBuilder(BrowseResult.ToString)
            builder.Replace(LoopBackIPv4Address, PlugInIPAddress)
            BrowseResult = builder.ToString
            If PIDebuglevel > DebugLevel.dlEvents Then Log("GetNextPicture changed the meta data for Loopback IP to = " & BrowseResult.ToString, LogType.LOG_TYPE_INFO)
            builder = Nothing
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then PrintOutXML(BrowseResult)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then PrintOutXML(BrowseResult)
        Dim ItemMetadata As String = BrowseResult
        Dim xmlData As XmlDocument = New XmlDocument

        OutArg = Nothing
        Try
            xmlData.LoadXml(BrowseResult.ToString)
        Catch ex As Exception
            Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & BrowseResult.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Dim ItemURL As String = ""
        ItemURL = FindRightSizePictureURL(BrowseResult.ToString, MyPictureSize)

        Dim ItemClassInfo As String = ""
        Dim ItemClassInfos As String() = Nothing
        Try
            ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
            ItemClassInfos = Split(ItemClassInfo, ".")
            If UCase(ItemClassInfos(1)) <> "ITEM" Then
                ' this should not be
                Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & ". The Object is a container, it is = " & ItemClassInfo.ToString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & " while searching XML for upnp:class. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_ERROR)
        End Try

        '    <res colorDepth="24" protocolInfo="http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=00;DLNA.ORG_CI=0;DLNA.ORG_FLAGS=00D00000000000000000000000000000" resolution="2048x1536" size="1324894">http://192.168.1.123:8895/resource/122/MEDIA_ITEM/JPEG_LRG*0</res>
        If ItemURL = "" Then
            Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & ". URL is empty. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        'ProtocolInfo = xmlData.GetElementsByTagName("res").Item(ResCount).Attributes("protocolInfo").Value
        'If ProtocolInfo <> "" Then
        'Dim ProtocolInfos As String()
        'ProtocolInfos = Split(ProtocolInfo, ";") ' in the format http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=00;DLNA.ORG_CI=0;DLNA.ORG_FLAGS=00D00000000000000000000000000000
        'CheckForProtocol(False, ProtocolInfos(0))
        'End If
        ' pick up the album art in case the player is not capable of returning it
        If MyUPnPDeviceServiceType = "HST" Then
            MyNextArtworkURL = GetAlbumArtPath(ItemURL, True)
        Else
            Try
                Dim AlbumArtURL As String = ""
                AlbumArtURL = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                If AlbumArtURL <> "" Then
                    If MyNextArtworkURL <> AlbumArtURL Then
                        MyNextAlbumArtURIHasChanged = True
                        MyNextArtworkURL = GetAlbumArtPath(AlbumArtURL, True)
                    End If
                End If
            Catch ex As Exception
                ' no big deal, there is no album art
                MyNextAlbumArtURIHasChanged = True
                If UCase(ItemClassInfos(2)) <> "IMAGEITEM" Then
                    MyNextArtworkURL = GetAlbumArtPath("", True)
                Else
                    MyNextArtworkURL = GetAlbumArtPath(ItemURL, True)
                End If
            End Try
        End If

        ' OK we should have a URL and Metadata
        Dim ItemName As String = ""
        Try
            ItemName = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextPicture for HST device = " & MyUPnPDeviceName & " found itemName = " & ItemName, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetNextPicture for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_ERROR)
        End Try
        Dim ItemAlbum As String = ""
        Try
            ItemAlbum = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayDBItem for HST device = " & MyUPnPDeviceName & " found ItemAlbum = " & ItemAlbum, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
        End Try
        NextTrack = ItemName
        NextAlbum = ItemAlbum
        NextArtist = "Pictures"
        PlayChangeNotifyCallback(player_status_change.NextSong, MyCurrentPlayerState)

    End Sub



    Public Function GetCurrentPlaylistTracks() As System.Array
        ' this is something undocumented but used by the ActionUI
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetCurrentPlaylistTracks called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        GetCurrentPlaylistTracks = GetQueue()
    End Function

    Public Sub ClearCurrentPlayList()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ClearCurrentPlayList called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        ClearQueue()
        SaveCurrentPlaylistTracks("")
    End Sub

    Public Sub AddTrackToCurrentPlaylist(ByVal AdditionalInfo As String, ByVal ServerUDN As String, Optional QueueAction As QueueActions = QueueActions.qaPlayLast)
        ' addionalInfo is ObjectName ":;:-:"  ObjectID  ":;:-:" ItemClass (CONTAINER or ITEM)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddTrackToCurrentPlaylist called for device - " & MyUPnPDeviceName & " with AdditionalInfo = " & AdditionalInfo & " and Server = " & ServerUDN, LogType.LOG_TYPE_INFO)
        Try
            Dim Infos As String() = Split(AdditionalInfo, ":;:-:")
            Dim QElement As myQueueElement = New myQueueElement
            QElement.Title = Infos(0)
            QElement.ObjectID = Infos(1)
            QElement.ServerUDN = ServerUDN
            Dim Index As Integer = 1
            Dim ObjectPath As String = ""
            ObjectPath = GetStringIniFile("DevicePage", MyUDN & "_" & "level" & Index.ToString, "")
            For Index = 2 To 10
                Dim ObjectTitle As String = GetStringIniFile("DevicePage", MyUDN & "_" & "level" & Index.ToString, "")
                If ObjectTitle = "" Then
                    ObjectPath = ObjectPath & ";:@;" & Infos(0)
                    Exit For
                End If
                ObjectPath = ObjectPath & ";:@;" & GetStringIniFile("DevicePage", MyUDN & "_" & "level" & Index.ToString, "")
            Next
            QElement.ObjectPath = ObjectPath
            If UBound(Infos) > 1 Then
                QElement.UPnPClass = ProcessClassInfo(Infos(2))
                'QElement.IconURL = Infos(4)
                'QElement.ArtistName = Infos(5)
                'QElement.AlbumName = Infos(6)
                'QElement.Genre = Infos(7)
                AddElementToQueue(QElement, False, QueueAction)
            Else
                AddElementToQueue(QElement, True, QueueAction)
            End If
        Catch ex As Exception
            Log("Error in AddTrackToCurrentPlaylist for device = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub AddObjectToCurrentPlaylist(ByVal ObjectID As String, ByVal ServerUDN As String, NavigationString As String, QueueAction As QueueActions)
        '
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist called for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & " and Server = " & ServerUDN & " and NavigationString = " & NavigationString, LogType.LOG_TYPE_INFO)
        Dim ServerApi As HSPI
        ServerApi = MyReferenceToMyController.GetAPIByUDN(ServerUDN)
        If ServerApi Is Nothing Then
            Log("Error in AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & ". Unable to get ServerApi for devicename = " & ServerUDN, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " retrieved DeviceAPI for device = " & ServerApi.DeviceName, LogType.LOG_TYPE_INFO)
        Dim OutArg(3) As Object
        OutArg = ServerApi.Browse(ObjectID, "BrowseMetadata", "*", 0, 1, "")
        If OutArg Is Nothing Then
            Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse from server device = " & ServerUDN & " returned NOK", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Dim BrowseResult As String = OutArg(0)
        Dim BrowseNumberReturned As Integer = OutArg(1)
        If BrowseNumberReturned < 1 Then
            Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned = " & BrowseNumberReturned.ToString & " records", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        BrowseResult = Trim(BrowseResult)
        If BrowseResult = "" Then
            Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse returned empty XML", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If MyReplaceLoopbackIPAddress And PlugInIPAddress <> "" Then
            Dim builder As New System.Text.StringBuilder(BrowseResult.ToString)
            builder.Replace(LoopBackIPv4Address, PlugInIPAddress)
            BrowseResult = builder.ToString
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AddObjectToCurrentPlaylist changed the meta data for Loopback IP to = " & BrowseResult.ToString, LogType.LOG_TYPE_INFO)
            builder = Nothing
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then PrintOutXML(BrowseResult)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then PrintOutXML(BrowseResult)
        Dim ItemMetadata As String = BrowseResult
        Dim xmlData As XmlDocument = New XmlDocument

        OutArg = Nothing
        Try
            xmlData.LoadXml(BrowseResult.ToString)
        Catch ex As Exception
            Log("Error in AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & BrowseResult.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Dim NbrOfObjects As Integer = 1
        Dim ItemClassInfo As String = ""
        Dim ItemClassInfos As String() = Nothing
        Try
            ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
            ItemClassInfos = Split(ItemClassInfo, ".")
            If UCase(ItemClassInfos(1)) <> "ITEM" Then
                OutArg = ServerApi.Browse(ObjectID, "BrowseDirectChildren", "*", 0, 0, "")
                If OutArg Is Nothing Then
                    Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse container from server device = " & ServerUDN & " returned NOK", LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End If
                BrowseResult = OutArg(0)
                BrowseNumberReturned = OutArg(1)
                If BrowseNumberReturned < 1 Then
                    Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse container returned = " & BrowseNumberReturned.ToString & " records", LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End If
                NbrOfObjects = BrowseNumberReturned
                BrowseResult = Trim(BrowseResult)
                If BrowseResult = "" Then
                    Log("Error in AddObjectToCurrentPlaylist for device - " & MyUPnPDeviceName & " with ObjectID = " & ObjectID & ". Browse container returned empty XML", LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End If
                If MyReplaceLoopbackIPAddress And PlugInIPAddress <> "" Then
                    Dim builder As New System.Text.StringBuilder(BrowseResult.ToString)
                    builder.Replace(LoopBackIPv4Address, PlugInIPAddress)
                    BrowseResult = builder.ToString
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("AddObjectToCurrentPlaylist changed the meta data for Loopback IP to = " & BrowseResult.ToString, LogType.LOG_TYPE_INFO)
                    builder = Nothing
                End If
                OutArg = Nothing
                Try
                    xmlData.LoadXml(BrowseResult.ToString)
                Catch ex As Exception
                    Log("Error in AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & BrowseResult.ToString, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
            End If
        Catch ex As Exception
            Log("Error in AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " while searching XML for upnp:class. XML = " & ItemMetadata.ToString, LogType.LOG_TYPE_ERROR)
        End Try

        Dim ObjectNavigationParts As String()
        Dim ObjectPath As String = ""

        Try
            If NavigationString <> "" Then
                ObjectNavigationParts = Split(NavigationString, ";--::")
                If ObjectNavigationParts.Count <> 0 Then
                    For Each objectnavigationPart In ObjectNavigationParts
                        Dim ObjectParts As String() = Nothing
                        ObjectParts = Split(objectnavigationPart, ";::--")
                        If ObjectParts IsNot Nothing Then
                            If ObjectParts.Count > 0 Then
                                If ObjectPath <> "" Then ObjectPath = ObjectPath & ";:@;"
                                ObjectPath = ObjectPath & ObjectParts(0)
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Log("Error in AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " while building the ObjectPath with NavigationString = " & NavigationString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Dim LoopIndex As Integer = 0
        Do While LoopIndex < NbrOfObjects
            Try
                ObjectID = xmlData.GetElementsByTagName("item").Item(LoopIndex).Attributes("id").Value 'GetElementsByTagName("dc:title").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found ObjectID = " & ObjectID, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try
            Dim ItemName As String = ""
            Try
                ItemName = xmlData.GetElementsByTagName("dc:title").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found itemName = " & ItemName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try

            Dim ItemAlbum As String = ""
            Try
                ItemAlbum = xmlData.GetElementsByTagName("upnp:album").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found ItemAlbum = " & ItemAlbum, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try

            Dim ItemArtist As String = ""
            Try
                ItemArtist = xmlData.GetElementsByTagName("upnp:artist").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found ItemArtist = " & ItemArtist, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try

            Dim ItemGenre As String = ""
            Try
                ItemGenre = xmlData.GetElementsByTagName("upnp:genre").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found ItemGenre = " & ItemGenre, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try

            Dim ItemArtURI As String = ""
            Try
                ItemArtURI = xmlData.GetElementsByTagName("upnp:upnp:albumArtURI").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found ItemArtURI = " & ItemArtURI, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                'log( "Error in PlayDBItem for device = " & MyUPnPDeviceName & " while searching XML for dc:title. XML = " & ItemMetadata.ToString)
            End Try
            Try
                ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(LoopIndex).InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddObjectToCurrentPlaylist for device = " & MyUPnPDeviceName & " found class = " & ItemClassInfo, LogType.LOG_TYPE_INFO)
            Catch ex As Exception

            End Try
            Dim QElement As myQueueElement = New myQueueElement

            Try
                QElement.Title = ItemName
                QElement.ObjectID = ObjectID
                QElement.ServerUDN = ServerUDN
                QElement.ObjectPath = ObjectPath
                QElement.UPnPClass = ProcessClassInfo(ItemClassInfo)
                QElement.IconURL = ItemArtURI
                QElement.ArtistName = ItemArtist
                QElement.AlbumName = ItemAlbum
                QElement.Genre = ItemGenre
                AddElementToQueue(QElement, False, QueueAction)
            Catch ex As Exception
                Log("Error in AddTrackToCurrentPlaylist for device = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                QElement = Nothing
            End Try
            LoopIndex = LoopIndex + 1
        Loop

    End Sub


    Public Sub SaveCurrentPlaylistTracks(Optional inPlayListName As String = "")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveCurrentPlaylistTracks called with PlayListName = " & inPlayListName, LogType.LOG_TYPE_INFO)
        'If MyQueueLinkedList Is Nothing Then Exit Sub ' Nothing to save
        Dim PlaylistName As String
        Dim PlayListId As String = ""
        If inPlayListName <> "" Then
            PlaylistName = CurrentAppPath & gPlaylistPath & inPlayListName & ".xml"
            PlayListId = inPlayListName
        Else
            PlaylistName = CurrentAppPath & gPlaylistPath & MyUDN & ".xml"
            PlayListId = MyUDN
        End If
        Try
            If System.IO.File.Exists(PlaylistName) = True Then System.IO.File.Delete(PlaylistName)
        Catch ex As Exception
            Log("Error in SaveCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " deleting the Playlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Dim objWriter As New System.IO.StreamWriter(PlaylistName)
        Try
            Dim QEntry As String = "<PlayList><Version>1.0</Version><PlayListName>" & PlayListId & "</PlayListName>"
            For Each QElement As myQueueElement In MyQueueLinkedList
                QEntry = QEntry & "<Item><Title>" & EncodeURI(QElement.Title) & "</Title>"
                QEntry = QEntry & "<ServerUDN>" & QElement.ServerUDN & "</ServerUDN>"
                QEntry = QEntry & "<UPnPClass>" & QElement.UPnPClass & "</UPnPClass>"
                QEntry = QEntry & "<IconURL>" & EncodeURI(QElement.IconURL) & "</IconURL>"
                QEntry = QEntry & "<ObjectID>" & EncodeURI(QElement.ObjectID) & "</ObjectID>"
                If QElement.ObjectPath <> "" Then
                    QEntry = QEntry & "<Path>"
                    Dim PathElements As String() = Split(QElement.ObjectPath, ";:@;")
                    For Each PathElement In PathElements
                        QEntry = QEntry & "<PathItem>" & EncodeURI(PathElement) & "</PathItem>"
                    Next
                    QEntry = QEntry & "</Path>"
                End If
                QEntry = QEntry & "</Item>"
            Next
            QEntry = QEntry & "</PlayList>"
            objWriter.WriteLine(QEntry)
        Catch ex As Exception
            Log("Error in SaveCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            objWriter.Close()
        Catch ex As Exception
        End Try
        objWriter.Dispose()
    End Sub

    Public Sub LoadCurrentPlaylistTracks(Optional inPlayListName As String = "")
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadCurrentPlaylistTracks called with PlayListName = " & inPlayListName, LogType.LOG_TYPE_INFO)
        If Not MyQueueLinkedList Is Nothing Then
            ClearQueue()
        End If
        Dim PlaylistName As String
        If inPlayListName <> "" Then
            PlaylistName = CurrentAppPath & gPlaylistPath & inPlayListName & ".xml"
        Else
            PlaylistName = CurrentAppPath & gPlaylistPath & MyUDN & ".xml"
        End If
        Try
            If Not System.IO.File.Exists(PlaylistName) = True Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadCurrentPlaylistTracks didn't find a Playlist for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " deleting the Playlist with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim xmlDoc As New XmlDocument
        Try
            xmlDoc.Load(PlaylistName)
        Catch ex As Exception
            Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " loading the XML for Playlist = " & PlaylistName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim PlayListID As String = ""
        Dim Version As String = ""
        Try
            PlayListID = xmlDoc.GetElementsByTagName("PlayListName").Item(0).InnerText
        Catch ex As Exception
            Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " retrieving the PlayListName for Playlist = " & PlaylistName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Version = xmlDoc.GetElementsByTagName("Version").Item(0).InnerText
        Catch ex As Exception
            Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " retrieving the Version ID for Playlist = " & PlaylistName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Dim ItemCount As Integer = 0
        Do
            Dim ObjectXMLString As String = ""
            Try
                ObjectXMLString = xmlDoc.GetElementsByTagName("Item").Item(ItemCount).OuterXml
            Catch ex As Exception
                ObjectXMLString = ""
                Exit Do
            End Try
            ObjectXMLString = Trim(ObjectXMLString)
            If ObjectXMLString = "" Then Exit Do
            ItemCount = ItemCount + 1
            Dim ObjectXML As New XmlDocument
            Try
                ObjectXML.LoadXml(ObjectXMLString)
            Catch ex As Exception
                Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " loading the item XML for Playlist = " & PlaylistName & " and XML = " & ObjectXMLString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                xmlDoc = Nothing
                Exit Do
            End Try
            Dim QElement As myQueueElement = New myQueueElement
            Try
                QElement.Title = DecodeURI(ObjectXML.GetElementsByTagName("Title").Item(0).InnerText)
            Catch ex As Exception
            End Try
            Try
                QElement.ServerUDN = ObjectXML.GetElementsByTagName("ServerUDN").Item(0).InnerText
            Catch ex As Exception
            End Try
            Try
                QElement.ObjectID = DecodeURI(ObjectXML.GetElementsByTagName("ObjectID").Item(0).InnerText)
            Catch ex As Exception
            End Try
            Try
                QElement.UPnPClass = CType(Val(ObjectXML.GetElementsByTagName("UPnPClass").Item(0).InnerText), UPnPClassType)
            Catch ex As Exception
            End Try
            Try
                QElement.IconURL = DecodeURI(ObjectXML.GetElementsByTagName("IconURL").Item(0).InnerText)
            Catch ex As Exception
            End Try
            Dim PathXMLString As String = ""
            Try
                PathXMLString = ObjectXML.GetElementsByTagName("Path").Item(0).OuterXml
            Catch ex As Exception
                PathXMLString = ""
            End Try
            Dim ObjectPath As String = ""
            If PathXMLString <> "" Then
                Dim PathXML As New XmlDocument
                Try
                    PathXML.LoadXml(PathXMLString)
                Catch ex As Exception
                    Log("Error in LoadCurrentPlaylistTracks for device = " & MyUPnPDeviceName & " Loading the Path XML for Playlist = " & PlaylistName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Dim PathIndex As Integer = 0
                Do
                    Dim TempPath As String = ""
                    Try
                        TempPath = DecodeURI(PathXML.GetElementsByTagName("PathItem").Item(PathIndex).InnerText)
                        PathIndex = PathIndex + 1
                    Catch ex As Exception
                        Exit Do
                    End Try
                    If TempPath = "" Then Exit Do
                    If ObjectPath <> "" Then ObjectPath = ObjectPath & ";:@;"
                    ObjectPath = ObjectPath & TempPath
                Loop
            End If
            QElement.ObjectPath = ObjectPath
            AddElementToQueue(QElement, False, QueueActions.qaPlayLast)
        Loop
        xmlDoc = Nothing
        If inPlayListName <> "" Then SaveCurrentPlaylistTracks("")
    End Sub


    Private Function GetQueue(Optional QueueName As String = "") As System.Array
        GetQueue = Nothing
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetQueue called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If MyQueueLinkedList Is Nothing Then Exit Function
        Dim SList() As String = {""}
        Dim KeyIndex As Integer = 0
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                Dim QEntry As String = ""
                QEntry = QElement.Title & ":;:-:" & QElement.ObjectID & ":;:-:" & QElement.ServerUDN & ":;:-:" & QElement.UPnPClass & ":;:-:" & QElement.IconURL & ":;:-:" & QElement.ArtistName & ":;:-:" & QElement.AlbumName & ":;:-:" & QElement.Genre
                ReDim Preserve SList(KeyIndex)
                SList(KeyIndex) = QEntry
                KeyIndex = KeyIndex + 1
            Next
        Catch ex As Exception
            Log("Error in GetQueue for device = " & MyUPnPDeviceName & " for queue = " & QueueName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If KeyIndex > 0 Then
            GetQueue = SList
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetQueue called for device = " & MyUPnPDeviceName & " and returned " & KeyIndex.ToString & " entries.", LogType.LOG_TYPE_INFO)
        End If
        SList = Nothing
    End Function

    Private Sub ClearQueue()
        If MyQueueLinkedList Is Nothing Then Exit Sub
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ClearQueue called for device = " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            MyQueueLinkedList.Clear()
            MyQueueHasChanged = True
            If MyUPnPDeviceServiceType <> "HST" Then MyPlayFromQueue = False
        Catch ex As Exception
            Log("Error in ClearQueue for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub AddElementToQueue(QueueElement As myQueueElement, LookUpClass As Boolean, QueueAction As QueueActions)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "AddElementToQueue called for device - " & MyUPnPDeviceName & " with LookUpClass = " & LookUpClass.ToString)
        Dim ItemMetaData As String = ""
        If LookUpClass Then
            Try
                If QueueElement.ServerUDN <> "" Then
                    Dim serverAPI As HSPI = MyReferenceToMyController.GetAPIByUDN(QueueElement.ServerUDN)
                    If Not serverAPI Is Nothing Then
                        Dim OutArg(3)
                        OutArg = serverAPI.Browse(QueueElement.ObjectID, "BrowseMetadata", "*", 0, 0, "")
                        ItemMetaData = OutArg(0)
                        If ItemMetaData <> "" Then
                            QueueElement.MetaData = ItemMetaData
                            Dim xmldata As XmlDocument = New XmlDocument
                            xmldata.LoadXml(ItemMetaData)
                            Dim MyItemClass As String = xmldata.GetElementsByTagName("upnp:class").Item(0).InnerText
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddElementToQueue found for device - " & MyUPnPDeviceName & " class = " & MyItemClass, LogType.LOG_TYPE_INFO)
                            QueueElement.UPnPClass = ProcessClassInfo(MyItemClass)
                            Dim MyIconURL As String = ""
                            If QueueElement.UPnPClass = UPnPClassType.ctPictures Then
                                Try
                                    MyIconURL = xmldata.GetElementsByTagName("res").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                            Else
                                Try
                                    MyIconURL = xmldata.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                                Try
                                    QueueElement.AlbumName = xmldata.GetElementsByTagName("upnp:album").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                                Try
                                    QueueElement.ArtistName = xmldata.GetElementsByTagName("upnp:artist").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                                Try
                                    QueueElement.Genre = xmldata.GetElementsByTagName("upnp:genre").Item(0).InnerText
                                Catch ex As Exception
                                End Try
                            End If
                            If MyIconURL <> "" Then
                                Dim IConImage As Bitmap
                                Try
                                    IConImage = New Bitmap(MyIconURL)
                                Catch ex As Exception
                                    'just use the no art image
                                    IConImage = New Bitmap(CurrentAppPath & "/html" & NoArtPath)
                                    Log("Error in AddElementToQueue for device - " & MyUPnPDeviceName & " getting imageart with error = " & ex.Message & " Path = " & MyIconURL.ToString, LogType.LOG_TYPE_ERROR)
                                End Try
                                'Dim IConImage As Image
                                'IConImage. = 150
                                Try
                                    Dim propItem As System.Drawing.Imaging.PropertyItem
                                    For Each propItem In IConImage.PropertyItems
                                        Log("Picture info in AddElementToQueue for device - " & MyUPnPDeviceName & " Id = " & propItem.Id.ToString, LogType.LOG_TYPE_INFO)
                                        Log("Picture info in AddElementToQueue for device - " & MyUPnPDeviceName & " Type = " & propItem.Type.ToString, LogType.LOG_TYPE_INFO)
                                        Log("Picture info in AddElementToQueue for device - " & MyUPnPDeviceName & " Value = " & propItem.Value.ToString, LogType.LOG_TYPE_INFO)
                                        Log("Picture info in AddElementToQueue for device - " & MyUPnPDeviceName & " Length = " & propItem.Len.ToString, LogType.LOG_TYPE_INFO)
                                    Next
                                Catch ex As Exception
                                    Log("Error in AddElementToQueue for device - " & MyUPnPDeviceName & " getting imageinfo with error = " & ex.Message & " Path = " & MyIconURL.ToString, LogType.LOG_TYPE_ERROR)
                                End Try

                                IConImage.Dispose()


                                'GetPicture(MyIconURL)
                            End If
                            QueueElement.IconURL = MyIconURL
                        End If
                    End If
                End If
            Catch ex As Exception
                Log("Error in AddElementToQueue for device = " & MyUPnPDeviceName & " retrieving class info from UDN = " & QueueElement.ServerUDN.ToString & " and ObjectID = " & QueueElement.ObjectID.ToString & " with error = " & ex.Message & " and MetaData = " & ItemMetaData.ToString, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        Try
            MyQueueLinkedList.AddLast(QueueElement)
            MyQueueHasChanged = True
            If PIDebuglevel > DebugLevel.dlEvents Then Log("AddElementToQueue for device = " & MyUPnPDeviceName & " added to queue = " & QueueElement.Title.ToString, LogType.LOG_TYPE_INFO)
            MyPlayFromQueue = True
        Catch ex As Exception
            Log("Error in AddElementToQueue for device = " & MyUPnPDeviceName & " adding to queue = " & QueueElement.Title.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DeleteFromQueue(ObjectId As String)
        If MyQueueLinkedList Is Nothing Then Exit Sub
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                If QElement.ObjectID = ObjectId Then
                    MyQueueLinkedList.Remove(QElement)
                    MyQueueHasChanged = True
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteFromQueue for device = " & MyUPnPDeviceName & " removed from queue = " & QElement.Title.ToString, LogType.LOG_TYPE_INFO)
                    If MyQueueLinkedList.Count = 0 And MyUPnPDeviceServiceType <> "HST" Then MyPlayFromQueue = False
                    QElement = Nothing
                End If
                Exit Sub
            Next
        Catch ex As Exception
            Log("Error in DeleteFromQueue for device = " & MyUPnPDeviceName & " removing from queue ObjectId  = " & ObjectId & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function GetElementFromQueueByObjectId(ObjectID As String) As myQueueElement
        GetElementFromQueueByObjectId = Nothing
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetElementFromQueueByObjectId called for device = " & MyUPnPDeviceName & " with ObjectId = " & ObjectID.ToString)
        If MyQueueLinkedList Is Nothing Then Exit Function
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                If QElement.ObjectID = ObjectID Then
                    GetElementFromQueueByObjectId = QElement
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetElementFromQueueByObjectId for device = " & MyUPnPDeviceName & " found ObjectId = " & ObjectID & " and  Title = " & QElement.Title.ToString)
                    Exit Function
                End If
            Next
        Catch ex As Exception
            Log("Error in GetElementFromQueueByObjectId for device = " & MyUPnPDeviceName & " going through queue looking for ObjectId  = " & ObjectID & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then
            Log("Error in GetElementFromQueueByObjectId for device = " & MyUPnPDeviceName & " did not find ObjectId = " & ObjectID, LogType.LOG_TYPE_ERROR)
            For Each QElement As myQueueElement In MyQueueLinkedList
                Log("----------> ObjectId = " & QElement.ObjectID & " and  Title = " & QElement.Title.ToString, LogType.LOG_TYPE_ERROR)
            Next
        End If
    End Function

    Private Function GetElementFromQueueByIndex(Index As Integer) As myQueueElement
        GetElementFromQueueByIndex = Nothing
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetElementFromQueueByIndex called for device = " & MyUPnPDeviceName & " with index = " & Index.ToString)
        If MyQueueLinkedList Is Nothing Then Exit Function
        If Index = 0 Then Exit Function
        If Index > MyQueueLinkedList.Count Then
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "Error in GetElementFromQueueByIndex for device = " & MyUPnPDeviceName & " index = " & Index.ToString & " is larger then QueuelistCount = " & MyQueueLinkedList.Count.ToString)
            Exit Function
        End If
        Dim QueueIndex As Integer = 0
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                QueueIndex = QueueIndex + 1
                If QueueIndex = Index Then
                    GetElementFromQueueByIndex = QElement
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetElementFromQueueByIndex for device = " & MyUPnPDeviceName & " found ObjectId = " & QElement.ObjectID & " and  Title = " & QElement.Title.ToString)
                    Exit Function
                End If
            Next
        Catch ex As Exception
            Log("Error in GetElementFromQueueByIndex for device = " & MyUPnPDeviceName & " removing from queue index  = " & Index.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetElementFromQueueByIndex for device = " & MyUPnPDeviceName & " did not find index = " & Index.ToString, LogType.LOG_TYPE_ERROR)
            For Each QElement As myQueueElement In MyQueueLinkedList
                Log("----------> ObjectId = " & QElement.ObjectID & " and  Title = " & QElement.Title.ToString, LogType.LOG_TYPE_ERROR)
            Next
        End If
    End Function

    Private Function GetQueueIndex(ObjectID As String) As Integer
        GetQueueIndex = 0
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetQueueIndex called for device = " & MyUPnPDeviceName & " with ObjectId = " & ObjectID.ToString)
        If MyQueueLinkedList Is Nothing Then Exit Function
        Dim QIndex As Integer = 0
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                QIndex = QIndex + 1
                If QElement.ObjectID = ObjectID Then
                    GetQueueIndex = QIndex
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "GetQueueIndex for device = " & MyUPnPDeviceName & " found ObjectId = " & ObjectID & " at index = " & QIndex.ToString)
                    Exit Function
                End If
            Next
        Catch ex As Exception
            Log("Error in GetQueueIndex for device = " & MyUPnPDeviceName & " going through queue looking for ObjectId  = " & ObjectID & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then
            Log("Error in GetQueueIndex for device = " & MyUPnPDeviceName & " did not find ObjectId = " & ObjectID, LogType.LOG_TYPE_ERROR)
            For Each QElement As myQueueElement In MyQueueLinkedList
                Log("----------> ObjectId = " & QElement.ObjectID & " and  Title = " & QElement.Title.ToString, LogType.LOG_TYPE_ERROR)
            Next
        End If
    End Function

    Public Sub PlayFromQueue(QueueIndex As Integer, Optional StartPlayFlag As Boolean = True, Optional IgnoreShuffleState As Boolean = False)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayFromQueue called for device - " & MyUPnPDeviceName & " and QueueIndex = " & QueueIndex.ToString & " and StartPlayFlag = " & StartPlayFlag.ToString, LogType.LOG_TYPE_INFO)
        If MyQueueLinkedList.Count = 0 Then Exit Sub ' nothing in the queue
        If QueueIndex > MyQueueLinkedList.Count Then QueueIndex = MyQueueLinkedList.Count
        If QueueIndex = 0 Then Exit Sub
        If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
            AVTStop(0)
        End If
        If MyQueueShuffleState And Not IgnoreShuffleState Then
            MyCurrentQueueIndex = GenerateANewIndex(MyCurrentQueueIndex)
        Else
            MyCurrentQueueIndex = QueueIndex
        End If
        MyNbrOfQueueItemsPlayed = 0
        'MyQueuePlayState = player_state_values.Playing
        Dim QElement As myQueueElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
        If QElement Is Nothing Then
            MyQueueReEntry = False
            Exit Sub
        End If
        MyPlayerWentThroughPlayState = False
        PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True, StartPlayFlag)
        MyQueuePlayState = player_state_values.Playing
        MyTimeIveBeenWaitingForPlayState = 0
        If (NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST" Then
            NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
        End If
    End Sub

    Public Sub PlayFromQueue(Title As String, Optional StartPlayFlag As Boolean = True)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayFromQueue called for device - " & MyUPnPDeviceName & " and Title = " & Title.ToString & " and StartPlayFlag = " & StartPlayFlag.ToString, LogType.LOG_TYPE_INFO)
        If MyQueueLinkedList Is Nothing Then Exit Sub
        If MyQueueLinkedList.Count = 0 Then Exit Sub ' nothing in the queue
        Dim QueueIndex As Integer = 1
        Try
            For Each QElement As myQueueElement In MyQueueLinkedList
                If QElement.Title = Title Then
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "PlayFromQueue for device = " & MyUPnPDeviceName & " found ObjectId = " & ObjectID & " and  Title = " & QElement.Title.ToString)
                    If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
                        AVTStop(0)
                    End If
                    MyCurrentQueueIndex = QueueIndex
                    MyNbrOfQueueItemsPlayed = 0
                    'MyQueuePlayState = player_state_values.Playing
                    MyPlayerWentThroughPlayState = False
                    PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True, StartPlayFlag)
                    MyQueuePlayState = player_state_values.Playing
                    MyTimeIveBeenWaitingForPlayState = 0
                    If (NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST" Then
                        NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    End If
                    Exit Sub
                End If
                QueueIndex = QueueIndex + 1
            Next
        Catch ex As Exception
            Log("Error in PlayFromQueue for device = " & MyUPnPDeviceName & " going through queue looking for Title  = " & Title & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayFromQueue for device = " & MyUPnPDeviceName & " going through queue looking for Title  = " & Title & " but found nothing", LogType.LOG_TYPE_ERROR)

    End Sub


    Public Sub SaveQueue()
        If MyQueuePlayState <> player_state_values.Playing Then Exit Sub
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveQueue called for device = " & MyUPnPDeviceName & " and CurrentQueueIndex = " & MyCurrentQueueIndex, LogType.LOG_TYPE_INFO)
        Try
            AVTGetPositionInfo(0)
            MySavedTrackPosition = MyCurrentRelTime
        Catch ex As Exception
            MySavedTrackPosition = ""
            Log("Error in SaveQueue for Device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If MyQueueLinkedList.Count = 0 Then
            ' Queue was most likely cleared
            Exit Sub
        End If
        MySavedQueueIndex = MyCurrentQueueIndex
        AVTStop(0)
        MyQueuePlayState = player_state_values.Paused
        wait(0.5)
    End Sub

    Public Sub RestoreQueue()
        If MyQueuePlayState <> player_state_values.Paused Then Exit Sub
        MyCurrentQueueIndex = MySavedQueueIndex
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreQueue called for device = " & MyUPnPDeviceName & " and CurrentQueueIndex = " & MyCurrentQueueIndex, LogType.LOG_TYPE_INFO)
        Dim QElement As myQueueElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
        If QElement Is Nothing Then
            Exit Sub
        End If
        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreQueue for device = " & MyUPnPDeviceName & " found class = " & QElement.UPnPClass.ToString, LogType.LOG_TYPE_INFO)
            Select Case QElement.UPnPClass
                Case UPnPClassType.ctMusic, UPnPClassType.ctVideo
                    MyPlayerWentThroughPlayState = True
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreQueue for device = " & MyUPnPDeviceName & " has shuffle state = " & MyQueueShuffleState.ToString & " and RepeatState = " & MyQueueRepeatState.ToString, LogType.LOG_TYPE_INFO)
                    PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
                    NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    If MySavedTrackPosition <> "" Then
                        Dim MaxWait As Integer = 0
                        While MaxWait < 10 ' wait max 5 sec
                            If NoQueuePlayerState = player_state_values.Playing Then
                                AVTSeek(AVT_SeekMode.REL_TIME, MySavedTrackPosition)
                                MaxWait = 11
                            Else
                                wait(0.5)
                                MaxWait = MaxWait + 1
                            End If
                        End While
                    End If
                Case UPnPClassType.ctPictures
                    PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
                    NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    MyQueueDelay = 0
            End Select
        Catch ex As Exception
            Log("Error in RestoreQueue for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        'MyQueuePlayState = player_state_values.Playing
    End Sub

    Private Sub CheckQueue()
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckQueue called for device = " & MyUPnPDeviceName & " and MyQueuePlayState = " & MyQueuePlayState.ToString & " and RentryFlag = " & MyQueueReEntry.ToString & " and DeviceOnLine = " & DeviceOnLine, LogType.LOG_TYPE_INFO, LogColorNavy)
        If (MyQueuePlayState <> player_state_values.Playing And MyQueuePlayState <> player_state_values.Forwarding And MyQueuePlayState <> player_state_values.Rewinding And MyQueuePlayState <> player_state_values.Transitioning) Or Not DeviceOnLine Then Exit Sub
        If MyQueueReEntry Then Exit Sub
        MyQueueReEntry = True

        Try
            'If MyPollForTransportChangeFlag Then
            'AVTGetTransportInfo()
            'If MyCurrentPlayerState <> ConvertTransportStateToPlayerState(MyCurrentTransportState) Then
            'MyCurrentPlayerState = ConvertTransportStateToPlayerState(MyCurrentTransportState)
            'AVTGetPositionInfo()
            'End If
            'End If
        Catch ex As Exception
            Log("Error in CheckQueue for Device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If MyQueueLinkedList.Count = 0 Then
            ' Queue was most likely cleared
            ResetQueueParameters()
            If MyUPnPDeviceServiceType <> "HST" Then MyPlayFromQueue = False
            MyQueueReEntry = False
            Exit Sub
        End If

        If NextElementPrefetchState = AVT_PrefetchState.PfSGetNext And ((NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST") Then
            ' go fetch the next element
            If GetNextElementInQueue(False) Then
                ' this would mean we're done
                NextElementPrefetchState = AVT_PrefetchState.PfSEnd
            Else
                ' OK there are more MediaObjects
                Dim NextQElement As myQueueElement = GetElementFromQueueByIndex(MyNextQueueIndex)
                If NextQElement Is Nothing Then
                    NextElementPrefetchState = AVT_PrefetchState.PfSEnd
                Else
                    If NextAvTransportIsAvailable And UseNextAvTransport Then
                        PlayDBItem(NextQElement.ObjectID, NextQElement.ServerUDN, False)
                        NextElementPrefetchState = AVT_PrefetchState.PfSLoaded
                    Else
                        Select Case NextQElement.UPnPClass
                            Case UPnPClassType.ctPictures
                                If MyUPnPDeviceServiceType = "HST" Then
                                    GetNextPicture(NextQElement.ObjectID, NextQElement.ServerUDN)
                                    NextElementPrefetchState = AVT_PrefetchState.PfSLoaded
                                End If
                            Case Else
                                ' no need to do anything
                                NextElementPrefetchState = AVT_PrefetchState.PfSNotPicture
                        End Select
                    End If
                End If
            End If
        End If

        Dim QElement As myQueueElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
        If QElement Is Nothing Then
            MyQueueReEntry = False
            Exit Sub
        End If

        Try
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CheckQueue for device = " & MyUPnPDeviceName & " found class = " & QElement.UPnPClass.ToString & " and PlayerWentThroughPlayStateFlag = " & MyPlayerWentThroughPlayState.ToString & " and PlayerState = " & MyCurrentPlayerState.ToString & " and CurrentQueueIndex = " & MyCurrentQueueIndex)
            Select Case QElement.UPnPClass
                Case UPnPClassType.ctMusic, UPnPClassType.ctVideo
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckQueue for device = " & MyUPnPDeviceName & " found music/video. Current index = " & MyCurrentQueueIndex.ToString, LogType.LOG_TYPE_INFO)
                    If MyCurrentPlayerState = player_state_values.Stopped And MyPlayerWentThroughPlayState Then ' removed  Or MyCurrentPlayerState = player_state_values.Transitioning)
                        MyPlayerWentThroughPlayState = False
                        MyTimeIveBeenWaitingForPlayState = 0
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("CheckQueue for device = " & MyUPnPDeviceName & " has shuffle state = " & MyQueueShuffleState.ToString & " and RepeatState = " & MyQueueRepeatState.ToString, LogType.LOG_TYPE_INFO)
                        If GetNextElementInQueue(True) Then
                            MyQueueReEntry = False
                            Exit Sub
                        End If
                        QElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
                        PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
                        MyTimeIveBeenWaitingForPlayState = 0
                        NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    Else
                        ' let's build in some safety net incase it never transitions to playing
                        If MyCurrentPlayerState = player_state_values.Stopped Then
                            MyTimeIveBeenWaitingForPlayState = MyTimeIveBeenWaitingForPlayState + 1
                            If MyTimeIveBeenWaitingForPlayState > MaxWaitTimeToGoThroughPlayState Then
                                MyPlayerWentThroughPlayState = True ' this will force the next track to be selected
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in CheckQueue for device = " & MyUPnPDeviceName & " has not seen a start play or song change in = " & MaxWaitTimeToGoThroughPlayState.ToString & " seconds. Next track enforced.", LogType.LOG_TYPE_WARNING)
                                MyTimeIveBeenWaitingForPlayState = 0
                            End If
                        End If
                    End If
                Case UPnPClassType.ctPictures
                    ' we'll do our own playing
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "CheckQueue for device = " & MyUPnPDeviceName & " found picture. Current index = " & MyCurrentQueueIndex.ToString)
                    MyQueueDelay = MyQueueDelay + 1
                    If MyQueueDelay <= MyQueueDelayUserDefined Then
                        MyQueueReEntry = False
                        Exit Sub
                    End If
                    MyQueueDelay = 0
                    If (NextAvTransportIsAvailable And UseNextAvTransport) And Not NextElementPrefetchState = AVT_PrefetchState.PfSEnd Then
                        If UCase(AVTNext(0)) <> "OK" Then
                            PlayDBItem(MyNextQueueObjectID, MyNextQueueServerUDN, True)
                        Else
                            MyCurrentQueueIndex = MyNextQueueIndex
                            MyCurrentQueueObjectID = MyNextQueueObjectID
                            MyCurrentQueueServerUDN = MyNextQueueServerUDN
                        End If

                        NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    Else
                        If GetNextElementInQueue(True) Then
                            MyQueueReEntry = False
                            Exit Sub
                        End If
                        QElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
                        If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" And MyIssueStopInBetween Then
                            AVTStop(0)
                        End If
                        PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
                        NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
                    End If
                Case Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CheckQueue for device = " & MyUPnPDeviceName & " with unknown class = " & QElement.UPnPClass.ToString & " and ObjectID = " & QElement.ObjectID & " and Title = " & QElement.Title & " and UDN = " & QElement.ServerUDN, LogType.LOG_TYPE_ERROR)
                    MyQueueDelay = 0
                    If GetNextElementInQueue(True) Then
                        MyQueueReEntry = False
                        Exit Sub
                    End If
                    NextElementPrefetchState = AVT_PrefetchState.PfSNone
            End Select
        Catch ex As Exception
            Log("Error in CheckQueue for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyQueueReEntry = False
    End Sub

    Private Sub NextQueue()
        If MyQueuePlayState = player_state_values.Stopped Then Exit Sub
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "NextQueue called for device = " & MyUPnPDeviceName & " and QueueIsPlaying = " & MyQueueIsPlaying.ToString)
        Try
            If MyQueueLinkedList.Count = 0 Then
                ' Queue was most likely cleared
                If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
                    AVTStop(0)
                End If
                MyQueuePlayState = player_state_values.Stopped
                MyCurrentQueueIndex = 1
                MyNbrOfQueueItemsPlayed = 0
                If MyUPnPDeviceServiceType <> "HST" Then MyPlayFromQueue = False
                Exit Sub
            End If
            If MyQueueShuffleState Then
                ' do things random
                MyNbrOfQueueItemsPlayed = MyNbrOfQueueItemsPlayed + 1
                If Not MyQueueRepeatState Then
                    If MyNbrOfQueueItemsPlayed > MyQueueLinkedList.Count Then
                        MyQueuePlayState = player_state_values.Stopped
                        MyCurrentQueueIndex = 1
                        MyNbrOfQueueItemsPlayed = 0
                        Exit Sub
                    End If
                End If
                Dim PreviousIndex As Integer = MyCurrentQueueIndex
                MyCurrentQueueIndex = GenerateANewIndex(MyCurrentQueueIndex)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NextQueue for device = " & MyUPnPDeviceName & " generated random index = " & MyCurrentQueueIndex.ToString & " had PreviousIndex = " & PreviousIndex & " and NbrOfPlayed = " & MyNbrOfQueueItemsPlayed.ToString & " out of " & MyQueueLinkedList.Count.ToString, LogType.LOG_TYPE_INFO)
            Else
                MyCurrentQueueIndex = MyCurrentQueueIndex + 1
                If MyCurrentQueueIndex > MyQueueLinkedList.Count Then
                    If Not MyQueueRepeatState Then
                        MyQueuePlayState = player_state_values.Stopped
                        MyCurrentQueueIndex = 1
                        MyNbrOfQueueItemsPlayed = 0
                        Exit Sub
                    Else
                        MyCurrentQueueIndex = 1
                    End If
                End If
            End If
            If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
                AVTStop(0)
            End If
            Dim QElement As myQueueElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
            QElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
            PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
            If (NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST" Then
                NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            End If
        Catch ex As Exception
            Log("Error in NextQueue for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub PreviousQueue()
        If MyQueuePlayState = player_state_values.Stopped Then Exit Sub
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "PreviousQueue called for device = " & MyUPnPDeviceName & " and QueueIsPlaying = " & MyQueueIsPlaying.ToString)
        Try
            If MyQueueLinkedList.Count = 0 Then
                ' Queue was most likely cleared
                If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
                    AVTStop(0)
                End If
                MyQueuePlayState = player_state_values.Stopped
                MyCurrentQueueIndex = 1
                MyNbrOfQueueItemsPlayed = 0
                If MyQueueLinkedList.Count = 0 And MyUPnPDeviceServiceType <> "HST" Then MyPlayFromQueue = False
                Exit Sub
            End If
            If MyQueueShuffleState Then
                ' do things random
                'MyNbrOfQueueItemsPlayed = MyNbrOfQueueItemsPlayed + 1
                'If Not MyQueueRepeatState Then
                'If MyNbrOfQueueItemsPlayed > MyQueueLinkedList.Count Then
                'MyQueuePlayState = player_state_values.Stopped
                'MyCurrentQueueIndex = 1
                'MyNbrOfQueueItemsPlayed = 0
                'Exit Sub
                'End If
                'End If
                MyCurrentQueueIndex = MyPreviousQueueIndex
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("PreviousQueue for device = " & MyUPnPDeviceName & " generated random index = " & MyCurrentQueueIndex.ToString & " and NbrOfPlayed = " & MyNbrOfQueueItemsPlayed.ToString & " out of " & MyQueueLinkedList.Count.ToString, LogType.LOG_TYPE_INFO)
            Else
                MyCurrentQueueIndex = MyCurrentQueueIndex - 1
                If MyCurrentQueueIndex < 1 Then
                    If Not MyQueueRepeatState Then
                        MyQueuePlayState = player_state_values.Stopped
                        MyCurrentQueueIndex = 1
                        MyNbrOfQueueItemsPlayed = 0
                        Exit Sub
                    Else
                        MyCurrentQueueIndex = MyQueueLinkedList.Count
                    End If
                End If
            End If
            If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
                AVTStop(0)
            End If
            Dim QElement As myQueueElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
            QElement = GetElementFromQueueByIndex(MyCurrentQueueIndex)
            PlayDBItem(QElement.ObjectID, QElement.ServerUDN, True)
            If (NextAvTransportIsAvailable And UseNextAvTransport) Or MyUPnPDeviceServiceType = "HST" Then
                NextElementPrefetchState = AVT_PrefetchState.PfSGetNext
            End If
        Catch ex As Exception
            Log("Error in PreviousQueue for device = " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ResetQueueParameters()
        If MyCurrentPlayerState <> player_state_values.Stopped And MyUPnPDeviceServiceType <> "HST" Then
            AVTStop(0)
        End If
        MyQueuePlayState = player_state_values.Stopped
        MyCurrentQueueIndex = 1
        MyNbrOfQueueItemsPlayed = 0
        MyNextQueueIndex = 1
        MyNextQueueObjectID = ""
        MyNextQueueServerUDN = ""
        MyCurrentQueueObjectID = ""
        MyCurrentQueueServerUDN = ""
    End Sub

    Private Function GetNextElementInQueue(CurrentFlag As Boolean) As Boolean ' return true if end is reached
        GetNextElementInQueue = False
        If MyQueueShuffleState Then
            ' do things random
            MyNbrOfQueueItemsPlayed = MyNbrOfQueueItemsPlayed + 1
            If Not MyQueueRepeatState Then
                If MyNbrOfQueueItemsPlayed > MyQueueLinkedList.Count Then
                    If CurrentFlag Then ResetQueueParameters()
                    GetNextElementInQueue = True
                    Exit Function
                End If
            End If
            Dim PreviousIndex As Integer = 0
            If CurrentFlag Then
                PreviousIndex = MyCurrentQueueIndex
                MyCurrentQueueIndex = GenerateANewIndex(MyCurrentQueueIndex)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextElementInQueue for device = " & MyUPnPDeviceName & " had previously index = " & PreviousIndex & " generated current random index = " & MyCurrentQueueIndex.ToString & " and NbrOfPlayed = " & MyNbrOfQueueItemsPlayed.ToString & " out of " & MyQueueLinkedList.Count.ToString & " and CurrentFlag = " & CurrentFlag.ToString, LogType.LOG_TYPE_INFO)
            Else
                PreviousIndex = MyNextQueueIndex
                MyNextQueueIndex = GenerateANewIndex(MyNextQueueIndex)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextElementInQueue for device = " & MyUPnPDeviceName & " had previously Nextindex = " & PreviousIndex & " generated next random index = " & MyNextQueueIndex.ToString & " and NbrOfPlayed = " & MyNbrOfQueueItemsPlayed.ToString & " out of " & MyQueueLinkedList.Count.ToString & " and CurrentFlag = " & CurrentFlag.ToString, LogType.LOG_TYPE_INFO)
            End If
        Else
            If CurrentFlag Then
                MyCurrentQueueIndex = MyCurrentQueueIndex + 1
                If MyCurrentQueueIndex > MyQueueLinkedList.Count Then
                    If Not MyQueueRepeatState Then
                        ResetQueueParameters()
                        GetNextElementInQueue = True
                        Exit Function
                    Else
                        MyCurrentQueueIndex = 1
                    End If
                End If
            Else
                MyNextQueueIndex = MyCurrentQueueIndex + 1
                If MyNextQueueIndex > MyQueueLinkedList.Count Then
                    If Not MyQueueRepeatState Then
                        GetNextElementInQueue = True
                        Exit Function
                    Else
                        MyNextQueueIndex = 1
                    End If
                End If
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextElementInQueue for device = " & MyUPnPDeviceName & " generated next index = " & MyNextQueueIndex.ToString & " and CurrentFlag = " & CurrentFlag.ToString, LogType.LOG_TYPE_INFO)
        End If
    End Function

    Private Function GenerateANewIndex(PreviousIndex As Integer) As Integer
        GenerateANewIndex = 1
        If MyQueueLinkedList Is Nothing Then Exit Function
        MyPreviousQueueIndex = PreviousIndex
        Dim RetryIndex As Integer = 0
        Try
            Do While RetryIndex < 20
                RetryIndex += 1
                GenerateANewIndex = MyRandomNumberGenerator.Next(1, MyQueueLinkedList.Count + 1)
                If GenerateANewIndex <> PreviousIndex Then Exit Function
            Loop
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GenerateANewIndex for device = " & MyUPnPDeviceName & " with PreviousIndex = " & PreviousIndex.ToString & " and Queue Item Count = " & MyQueueLinkedList.Count.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function ProcessClassInfo(ClassString As String) As UPnPClassType
        ProcessClassInfo = UPnPClassType.ctUnknown
        If ClassString = "" Then Exit Function
        Dim ClassItems As String()
        ClassItems = ClassString.Split(".")
        If UBound(ClassItems) > 1 Then
            Select Case UCase(ClassItems(2))
                Case "IMAGEITEM"
                    ProcessClassInfo = UPnPClassType.ctPictures
                Case "AUDIOITEM"
                    ProcessClassInfo = UPnPClassType.ctMusic
                Case "VIDEOITEM"
                    ProcessClassInfo = UPnPClassType.ctVideo
                Case "PLAYLISTITEM"
                    ProcessClassInfo = UPnPClassType.ctMusic
                    'Case "TEXTITEM"
                    'Case "BOOKMARKITEM"
                    'Case "EPGITEM"
                Case Else
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning ProcessClassInfo called with unknown UPnPClass = " & ClassString, LogType.LOG_TYPE_WARNING)
            End Select
        End If
    End Function

    Private Sub ProcessSpeedSettings(SpeedInfo As String())
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessSpeedSettings called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        If SpeedInfo Is Nothing Then Exit Sub
        For index = 0 To UBound(SpeedInfo, 1)
            If SpeedInfo(index) <> "1" Then
                Try
                    Dim SpeedValue As Integer = Val(SpeedInfo(index))
                    If SpeedValue > 1 Then
                        SpeedIsConfigurable = True
                        If MyPossibleFFSpeeds Is Nothing Then
                            ReDim MyPossibleFFSpeeds(0)
                            MyPossibleFFSpeeds(0) = Val(SpeedInfo(index))
                        Else
                            ReDim Preserve MyPossibleFFSpeeds(UBound(MyPossibleFFSpeeds, 1) + 1)
                            MyPossibleFFSpeeds(UBound(MyPossibleFFSpeeds, 1)) = Val(SpeedInfo(index))
                        End If
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessSpeedSettings called for device - " & MyUPnPDeviceName & " added FF Speed = " & SpeedInfo(index), LogType.LOG_TYPE_INFO)
                    ElseIf SpeedValue < 0 Then
                        SpeedIsConfigurable = True
                        If MyPossibleREWSpeeds Is Nothing Then
                            ReDim MyPossibleREWSpeeds(0)
                            MyPossibleREWSpeeds(0) = Val(SpeedInfo(index))
                        Else
                            ReDim Preserve MyPossibleREWSpeeds(UBound(MyPossibleREWSpeeds, 1) + 1)
                            MyPossibleREWSpeeds(UBound(MyPossibleREWSpeeds, 1)) = Val(SpeedInfo(index))
                        End If
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessSpeedSettings called for device - " & MyUPnPDeviceName & " added REW Speed = " & SpeedInfo(index), LogType.LOG_TYPE_INFO)
                    End If
                Catch ex As Exception
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ProcessSpeedSettings for device - " & MyUPnPDeviceName & " converting Speed to an integer with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        Next

    End Sub

    Private Function FindRightSizePictureURL(inXML As String, inPictureSize As HS_GLOBAL_VARIABLES.PictureSize) As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL called for device - " & MyUPnPDeviceName & " and PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO)

        FindRightSizePictureURL = ""
        Dim xmlData As XmlDocument = New XmlDocument

        Try
            xmlData.LoadXml(inXML)
        Catch ex As Exception
            Log("Error in FindRightSizePictureURL for device = " & MyUPnPDeviceName & "  loading XML. Error = " & ex.Message & ". XML = " & inXML, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Dim ItemClassInfo As String = ""
        Dim ItemClassInfos As String() = Nothing

        Try
            ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
            ItemClassInfos = Split(ItemClassInfo, ".")
            If UCase(ItemClassInfos(1)) <> "ITEM" Then
                ' this should not be
                Log("Error in FindRightSizePictureURL for device = " & MyUPnPDeviceName & ". The Object is a container, it is = " & ItemClassInfo.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in FindRightSizePictureURL for device = " & MyUPnPDeviceName & " while searching XML for upnp:class. XML = " & inXML.ToString, LogType.LOG_TYPE_ERROR)
        End Try

        Dim ItemURL As String = ""
        Dim ProtocolInfo As String = ""
        Dim ResCount As Integer = 0
        Dim FirstTime As Boolean = True
        Dim ThisIsNotAnImage As Boolean = UCase(ItemClassInfos(2)) <> "IMAGEITEM"

        If ThisIsNotAnImage Then
            Try
                ItemURL = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                FindRightSizePictureURL = ItemURL
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL called for device - " & MyUPnPDeviceName & " found picture = " & FindRightSizePictureURL.ToString, LogType.LOG_TYPE_INFO)
                Exit Function
            Catch ex As Exception
            End Try
            Try
                ItemURL = xmlData.GetElementsByTagName("upnp:icon").Item(0).InnerText
                FindRightSizePictureURL = ItemURL
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL called for device - " & MyUPnPDeviceName & " found icon = " & FindRightSizePictureURL.ToString, LogType.LOG_TYPE_INFO)
                Exit Function
            Catch ex As Exception
            End Try
        End If

        If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psTiny Then
            ' we can also use the albumart if available
            Try
                FindRightSizePictureURL = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                If FindRightSizePictureURL <> "" Then
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL called for device - " & MyUPnPDeviceName & " found picture = " & FindRightSizePictureURL.ToString, LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
            Catch ex As Exception
            End Try
        End If

        Do
            Try
                Dim res As XmlNode
                'Dim PictureSize As Integer = 0
                'Dim PictureResolution As String = ""
                Dim PictureProtocolInfoString As String = ""
                Dim PictureURL As String
                res = xmlData.GetElementsByTagName("res").Item(ResCount)
                'Try
                'PictureSize = Val(res.Attributes("size").Value)
                'Catch ex As Exception
                'PictureSize = 0
                'End Try
                'Try
                'PictureResolution = res.Attributes("resolution").Value
                'Catch ex As Exception
                'PictureResolution = ""
                'End Try
                Try
                    PictureURL = res.InnerText
                    If FirstTime Then
                        FirstTime = False
                        ItemURL = PictureURL
                    End If
                Catch ex As Exception
                    PictureURL = ""
                End Try
                If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psDefault Then
                    FindRightSizePictureURL = PictureURL
                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and selected default PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO) ' & " and Resolution = " & PictureResolution & " and ProtocolInfo = " & PictureProtocolInfoString)
                    Exit Function
                End If
                Try
                    PictureProtocolInfoString = res.Attributes("protocolInfo").Value
                    If PictureProtocolInfoString <> "" Then
                        Dim PictureProtocolInfos As String()
                        PictureProtocolInfos = Split(PictureProtocolInfoString, ";") ' in the format http-get:*:image/jpeg:DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=00;DLNA.ORG_CI=0;DLNA.ORG_FLAGS=00D00000000000000000000000000000
                        PictureProtocolInfos = Split(PictureProtocolInfos(0), ":")
                        If UBound(PictureProtocolInfos) > 2 Then
                            If PictureProtocolInfos(3).Contains("_LRG") Then
                                If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psLarge Then
                                    FindRightSizePictureURL = PictureURL
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and selected PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO) ' & " and Resolution = " & PictureResolution & " and ProtocolInfo = " & PictureProtocolInfoString)
                                    Exit Function
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and PictureSize = LRG ", LogType.LOG_TYPE_INFO)
                                End If
                            ElseIf PictureProtocolInfos(3).Contains("_MED") Then
                                If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psMedium Then
                                    FindRightSizePictureURL = PictureURL
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and selected PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO) ' & " and Resolution = " & PictureResolution & " and ProtocolInfo = " & PictureProtocolInfoString)
                                    Exit Function
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and PictureSize = MED ", LogType.LOG_TYPE_INFO)
                                End If
                            ElseIf PictureProtocolInfos(3).Contains("_SM") Then
                                If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psSmall Then
                                    FindRightSizePictureURL = PictureURL
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and selected PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO) ' & " and Resolution = " & PictureResolution & " and ProtocolInfo = " & PictureProtocolInfoString)
                                    Exit Function
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and PictureSize = SM ", LogType.LOG_TYPE_INFO)
                                End If
                            ElseIf PictureProtocolInfos(3).Contains("_TN") Then
                                If inPictureSize = HS_GLOBAL_VARIABLES.PictureSize.psTiny Then
                                    FindRightSizePictureURL = PictureURL
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and selected PictureSize = " & inPictureSize.ToString, LogType.LOG_TYPE_INFO) ' & " and Resolution = " & PictureResolution & " and ProtocolInfo = " & PictureProtocolInfoString)
                                    Exit Function
                                Else
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindRightSizePictureURL found res info for device - " & MyUPnPDeviceName & " and PictureSize = TN ", LogType.LOG_TYPE_INFO)
                                End If
                            End If
                        End If
                    Else
                        Exit Do
                    End If
                Catch ex As Exception
                    PictureProtocolInfoString = ""
                End Try
            Catch ex As Exception
                Exit Do
            End Try
            ResCount = ResCount + 1
            If ResCount > 10 Then Exit Do
        Loop

        If Not ThisIsNotAnImage Then
            ' we did not find the right picture, but this is an image, take the first res
            Try
                Dim resNode As XmlNode = xmlData.GetElementsByTagName("res").Item(0)
                FindRightSizePictureURL = resNode.InnerText
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning FindRightSizePictureURL did not find res info for device - " & MyUPnPDeviceName & " and PictureSize = " & inPictureSize.ToString & " but used the first found", LogType.LOG_TYPE_WARNING)
                resNode = Nothing
            Catch ex As Exception
            End Try
        End If

    End Function

    Private Sub ExtractDateInformation(DateInfo As String)
        ' 2004-05-14T14:30:05+09:00 this is the format
        DateInfo = Trim(DateInfo)
        If DateInfo = "" Then Exit Sub
        Dim DateInfos As String()
        DateInfos = Split(DateInfo, "-")
        If UBound(DateInfos) < 1 Then Exit Sub ' nothing in it
        MyCurrentTrackYear = DateInfos(0)
        DateInfos = Nothing
    End Sub



    Public Function AVTPlay(Optional Speed As Integer = 1, Optional InstanceID As Integer = 0) As String
        AVTPlay = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTPlay called for device " & MyUPnPDeviceName & " and Speed = " & Speed.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Speed
            AVTransport.InvokeAction("Play", InArg, OutArg)
            AVTPlay = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTPlay for device = " & MyUPnPDeviceName & " and Speed = " & Speed.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTStop(Optional InstanceID As Integer = 0) As String
        AVTStop = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTStop called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As String
            Dim OutArg(0) As String
            InArg(0) = InstanceID
            AVTransport.InvokeAction("Stop", InArg, OutArg)
            AVTStop = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTStop for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTNext(Optional InstanceID As Integer = 0) As String
        AVTNext = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTNext called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("Next", InArg, OutArg)
            AVTNext = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTNext for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTPrevious(Optional InstanceID As Integer = 0) As String
        AVTPrevious = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTPrevious called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("Previous", InArg, OutArg)
            AVTPrevious = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTPrevious for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTSetPlayMode(Optional NewPlayMode As String = "NORMAL", Optional InstanceID As Integer = 0) As String
        AVTSetPlayMode = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTSetPlayMode called for device " & MyUPnPDeviceName & " and NewPlayMode = " & NewPlayMode.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = NewPlayMode
            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)
            AVTSetPlayMode = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTSetPlayMode for device = " & MyUPnPDeviceName & " and NewPlayMode = " & NewPlayMode.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTPause(Optional InstanceID As Integer = 0) As String
        AVTPause = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTPause called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("Pause", InArg, OutArg)
            AVTPause = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTPause for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTSeek(Unit As AVT_SeekMode, Target As String, Optional InstanceID As Integer = 0) As String
        ' Unit could be ABS_TIME, REL_TIME, ABS_COUNT, REL_COUNT, TRACK_NR
        AVTSeek = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTSeek called for device " & MyUPnPDeviceName & " and Unit = " & Unit.ToString & " and Target = " & Target.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = Unit.ToString
            InArg(2) = Trim(Target)
            AVTransport.InvokeAction("Seek", InArg, OutArg)
            AVTSeek = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTSeek for device = " & MyUPnPDeviceName & " and Unit = " & Unit.ToString & " and Target = " & Target.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTSetAVTransportURI(CurrentURI As String, CurrentURIMetaData As String, Optional InstanceID As Integer = 0) As String
        ' Onkyo example
        ' URI
        'http://192.168.1.103:9050/disk/music/DLNA-PNMP3-OP11-FLAGS01700000/O1$268435466$1744840323.mp3

        ' MetaData
        ' <DIDL-Lite xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:dlna="urn:schemas-dlna-org:metadata-1-0/"><item id="-1" parentID="-1" restricted="1"><dc:title>A Forest</dc:title><dc:creator>The Cure</dc:creator><dc:date>2000-01-01</dc:date><upnp:artist>The Cure</upnp:artist><upnp:album>Torhout Werchter 20</upnp:album><upnp:genre>Electronic</upnp:genre><dc:publisher>The Cure</dc:publisher><upnp:albumArtURI>http://192.168.1.103:9050/cgi-bin/O1$268435466$1744858504/W160/H160/S1/L1/Xjpeg-jpeg.desc.jpg</upnp:albumArtURI><upnp:originalTrackNumber>1</upnp:originalTrackNumber><res duration="0:05:12" protocolInfo="http-get:*:audio/mpeg:DLNA.ORG_PN=MP3;DLNA.ORG_OP=11;DLNA.ORG_FLAGS=01700000000000000000000000000000">http://192.168.1.103:9050/disk/music/DLNA-PNMP3-OP11-FLAGS01700000/O1$268435466$1744840323.mp3</res><upnp:class>object.item.audioItem.musicTrack</upnp:class></item></DIDL-Lite>
        AVTSetAVTransportURI = ""
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTSetAVTransportURI called for device " & MyUPnPDeviceName & " and CurrentURI = " & CurrentURI.ToString & " and CurrentURIMetaData = " & CurrentURIMetaData.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Function
        'If Not piDebuglevel > DebugLevel.dlEvents And piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTSetAVTransportURI called for device " & MyUPnPDeviceName & " and CurrentURI = " & CurrentURI.ToString & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2) As Object
            Dim OutArg(0) As Object
            InArg(0) = InstanceID
            InArg(1) = CurrentURI
            InArg(2) = CurrentURIMetaData
            AVTransport.InvokeAction("SetAVTransportURI", InArg, OutArg)
            AVTSetAVTransportURI = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTSetAVTransportURI for device = " & MyUPnPDeviceName & " and CurrentURI = " & CurrentURI.ToString & " and CurrentURIMetaData = " & CurrentURIMetaData.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTSetNextAVTransportURI(NextURI As String, NextURIMetaData As String, Optional InstanceID As Integer = 0) As String
        AVTSetNextAVTransportURI = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTSetNextAVTransportURI called for device " & MyUPnPDeviceName & " and NextURI = " & NextURI.ToString & " and NextURIMetaData = " & NextURIMetaData.ToString & " and InstanceID = " & InstanceID.ToString & " and ServiceIsAvailable = " & NextAvTransportIsAvailable.ToString, LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "AVTSetNextAVTransportURI called for device " & MyUPnPDeviceName & " and NextURI = " & NextURI.ToString & " and InstanceID = " & InstanceID.ToString & " and ServiceIsAvailable = " & NextAvTransportIsAvailable.ToString & " and NextIsAllowed = " & UseNextAvTransport.ToString)
        If Not NextAvTransportIsAvailable Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = InstanceID
            InArg(1) = NextURI
            InArg(2) = NextURIMetaData
            AVTransport.InvokeAction("SetNextAVTransportURI", InArg, OutArg)
            AVTSetNextAVTransportURI = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTSetNextAVTransportURI for device = " & MyUPnPDeviceName & " and NextURI = " & NextURI.ToString & " and NextURIMetaData = " & NextURIMetaData.ToString & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetMediaInfo(InstanceID As Integer, Optional PrintDebug As Boolean = True) As String
        AVTGetMediaInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If AVTransport Is Nothing Then Exit Function
        If PrintDebug Then Log("AVTGetMediaInfo called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As Object
            Dim OutArg(8) As Object
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetMediaInfo", InArg, OutArg)
            MyCurrentNrTracks = Val(OutArg(0))      ' UI4
            MyCurrentMediaDuration = OutArg(1)      ' String
            MyCurrentURI = OutArg(2)                ' String
            MyCurrentURIMetaData = OutArg(3)        ' String
            ProcessAVXML(MyCurrentURIMetaData, True, False, PIDebuglevel > DebugLevel.dlEvents)
            MyNextURI = OutArg(4)                   ' String
            MyNextURIMetaData = OutArg(5)           ' String
            ProcessAVXML(MyNextURIMetaData, True, True, PIDebuglevel > DebugLevel.dlEvents)
            MyCurrentPlayMedium = OutArg(6)         ' String
            MyCurrentRecordMedium = OutArg(7)       ' String
            MyCurrentWriteStatus = OutArg(8)        ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetMediaInfo : MyNrTracks           = " & MyCurrentNrTracks.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyMediaDuration      = " & MyCurrentMediaDuration.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyCurrentURI         = " & MyCurrentURI.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyCurrentURIMetaData = " & MyCurrentURIMetaData.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyNextURI            = " & MyNextURI.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyNextURIMetaData    = " & MyNextURIMetaData.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyPlayMedium         = " & MyCurrentPlayMedium.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyRecordMedium       = " & MyCurrentRecordMedium.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetMediaInfo : MyWriteStatus        = " & MyCurrentWriteStatus.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetMediaInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTGetMediaInfo for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetDeviceCapabilities(Optional InstanceID As Integer = 0) As String
        AVTGetDeviceCapabilities = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTGetDeviceCapabilities called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(2)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetDeviceCapabilities", InArg, OutArg)
            MyCurrentPlayMedia = OutArg(0)         ' String
            MyCurrentRecMedia = OutArg(1)          ' String
            MyCurrentRecQualityModes = OutArg(2)   ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetDeviceCapabilities : MyPlayMedia       =  " & MyCurrentPlayMedia.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetDeviceCapabilities : MyRecMedia        = " & MyCurrentRecMedia.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetDeviceCapabilities : MyRecQualityModes = " & MyCurrentRecQualityModes.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetDeviceCapabilities = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTGetDeviceCapabilities for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetTransportSettings(Optional InstanceID As Integer = 0) As String
        AVTGetTransportSettings = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTGetTransportSettings called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As Object
            Dim OutArg(1) As Object
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetTransportSettings", InArg, OutArg)
            MyCurrentPlayMode = OutArg(0)         ' String
            MyCurrentRecQualityMode = OutArg(1)   ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetTransportSettings : MyPlayMode       =  " & MyCurrentPlayMode.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetTransportSettings : MyRecQualityMode = " & MyCurrentRecQualityMode.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetTransportSettings = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTGetTransportSettings for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetTransportInfo(Optional InstanceID As Integer = 0) As String
        AVTGetTransportInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If AVTransport Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlEvents Then Log("AVTGetTransportInfo called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As Object
            Dim OutArg(2) As Object
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetTransportInfo", InArg, OutArg)
            CurrentPlayerState = ConvertTransportStateToPlayerState(OutArg(0).ToString)
            MyCurrentTransportState = OutArg(0)  ' String
            MyCurrentTransportStatus = OutArg(1) ' String
            If MyCurrentTransportPlaySpeed <> OutArg(2) Then
                MyCurrentTransportPlaySpeed = Val(OutArg(2))
                If MyCurrentTransportPlaySpeed < 0 Then
                    CurrentPlayerState = player_state_values.Rewinding
                    MyTransportStateHasChanged = True
                ElseIf MyCurrentTransportPlaySpeed > 1 Then
                    CurrentPlayerState = player_state_values.Forwarding
                    MyTransportStateHasChanged = True
                Else
                    If MyCurrentPlayerState = player_state_values.Forwarding Or MyCurrentPlayerState = player_state_values.Rewinding Then
                        CurrentPlayerState = player_state_values.Playing
                        MyTransportStateHasChanged = True
                    End If
                End If
                ' go set a flag to GetPositionInfo because so fast forwarding is happening ..... or just stopped
                'MyRefreshAVGetPositionInfo = True
            End If
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetTransportInfo : MyCurrentTransportState  =  " & MyCurrentTransportState.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetTransportInfo : MyCurrentTransportStatus = " & MyCurrentTransportStatus.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetTransportInfo : MyCurrentSpeed           = " & MyCurrentTransportPlaySpeed.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetTransportInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in AVTGetTransportInfo for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetPositionInfo(Optional InstanceID As Integer = 0) As String
        AVTGetPositionInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If AVTransport Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlEvents Then Log("AVTGetPositionInfo called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As Object
            Dim OutArg(7) As Object
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetPositionInfo", InArg, OutArg)
            MyCurrentTrackNumber = OutArg(0)         ' String
            If OutArg(1) <> "NOT_IMPLEMENTED" Then
                MyCurrentTrackDuration = OutArg(1) ' String
                SetTrackLength(GetSeconds(MyCurrentTrackDuration))
            End If
            MyCurrentTrackMetaData = OutArg(2) ' String
            ProcessAVXML(MyCurrentTrackMetaData, True, False, PIDebuglevel > DebugLevel.dlEvents)
            MyCurrentTrackURI = OutArg(3)      ' String
            If OutArg(4) <> "NOT_IMPLEMENTED" Then
                MyCurrentRelTime = OutArg(4)       ' String
                SetPlayerPosition(ConvertTimeFormatToSeconds(MyCurrentRelTime))
            End If
            MyCurrentAbsTime = OutArg(5)       ' String
            MyCurrentRelCount = Val(OutArg(6)) ' UI4
            MyCurrentAbsCount = Val(OutArg(7)) ' UI4
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetPositionInfo : MyTrack         =  " & MyCurrentTrack.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyTrackDuration = " & MyCurrentTrackDuration.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyTrackMetaData = " & MyCurrentTrackMetaData.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyTrackURI      = " & MyCurrentTrackURI.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyRelTime       = " & MyCurrentRelTime.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyAbsTime       = " & MyCurrentAbsTime.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyRelCount      = " & MyCurrentRelCount.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTGetPositionInfo : MyAbsCount      = " & MyCurrentAbsCount.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetPositionInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in AVTGetPositionInfo for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTGetCurrentTransportActions(Optional InstanceID As Integer = 0) As String
        AVTGetCurrentTransportActions = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTGetCurrentTransportActions called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("GetCurrentTransportActions", InArg, OutArg)
            MyActions = OutArg(0)   ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTGetCurrentTransportActions : MyActions =  " & MyActions.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTGetCurrentTransportActions = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AVTGetCurrentTransportActions for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AVTX_DLNA_GetBytePositionInfo(Optional InstanceID As Integer = 0) As String
        AVTX_DLNA_GetBytePositionInfo = ""
        If DeviceStatus = "Offline" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AVTX_DLNA_GetBytePositionInfo called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(2)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("X_DLNA_GetBytePositionInfo", InArg, OutArg)
            MyCurrentTrackSize = OutArg(0) ' String
            MyCurrentRelByte = OutArg(1)   ' String
            MyCurrentAbsByte = OutArg(2)   ' String
            If PIDebuglevel > DebugLevel.dlEvents Then
                Log("AVTX_DLNA_GetBytePositionInfo : MyTrackSize =  " & MyCurrentTrackSize.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTX_DLNA_GetBytePositionInfo : MyRelByte   =  " & MyCurrentRelByte.ToString, LogType.LOG_TYPE_INFO)
                Log("AVTX_DLNA_GetBytePositionInfo : MyAbsByte   =  " & MyCurrentAbsByte.ToString, LogType.LOG_TYPE_INFO)
            End If
            AVTX_DLNA_GetBytePositionInfo = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in AVTX_DLNA_GetBytePositionInfo for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetPositionDurationInfoOnly(Optional InstanceID As Integer = 0) As String
        GetPositionDurationInfoOnly = ""
        If DeviceStatus = "Offline" Then Exit Function
        If AVTransport Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetPositionDurationInfoOnly called for device " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0) As String
            Dim OutArg(7) As String
            InArg(0) = InstanceID.ToString
            AVTransport.InvokeAction("GetPositionInfo", InArg, OutArg)
            MyCurrentTrackNumber = OutArg(0)         ' String
            If OutArg(1) <> "NOT_IMPLEMENTED" Then
                MyCurrentTrackDuration = OutArg(1) ' String
                SetTrackLength(GetSeconds(MyCurrentTrackDuration))
            End If
            If OutArg(4) <> "NOT_IMPLEMENTED" Then
                MyCurrentRelTime = OutArg(4)       ' String
                SetPlayerPosition(ConvertTimeFormatToSeconds(MyCurrentRelTime))
            End If
            GetPositionDurationInfoOnly = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Error in GetPositionDurationInfoOnly for device = " & MyUPnPDeviceName & " and InstanceID = " & InstanceID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Enum AVT_SeekMode
        TRACK_NR = 0
        REL_TIME = 1
        SECTION = 3
        ABS_TIME = 4
        ABS_COUNT = 5
        REL_COUNT = 6
        X_DLNA_REL_BYTE = 7
    End Enum

    Public Enum AVT_TransportState
        STOPPED = 0
        PAUSED_PLAYBACK = 1
        PLAYING = 2
        TRANSITIONING = 3
        NO_MEDIA_PRESENT = 4
    End Enum

    Public Enum AVT_PrefetchState
        PfSNone = 0
        PfSGetNext = 1
        PfSLoaded = 2
        PfSEnd = 3
        PfSNotPicture = 4
    End Enum

End Class

Public Enum UPnPClassType
    ctUnknown = 0
    ctMusic = 1
    ctVideo = 2
    ctPictures = 3
End Enum

<Serializable()> _
Public Class myQueueElement
    Private MyTitle As String = ""
    Private MyObjectID As String = ""
    Private MyMetaData As String = ""
    Private MyServerUDN As String = ""
    Private MyUPNPClass As UPnPClassType
    Private MyIconURL As String = ""
    Private MyAlbumName As String = ""
    Private MyArtistName As String = ""
    Private MyGenre As String = ""
    Private MyObjectPath As String = ""
    Public Sub New()
        MyBase.New()
        MyTitle = ""
        MyObjectID = ""
        MyMetaData = ""
        MyServerUDN = ""
        MyUPNPClass = UPnPClassType.ctUnknown
        MyIconURL = ""
        MyAlbumName = ""
        MyArtistName = ""
        MyGenre = ""
        MyObjectPath = ""
    End Sub

    Property Title As String
        Get
            Title = MyTitle
        End Get
        Set(value As String)
            MyTitle = value
        End Set
    End Property
    Property ObjectID As String
        Get
            ObjectID = MyObjectID
        End Get
        Set(value As String)
            MyObjectID = value
        End Set
    End Property
    Property ServerUDN As String
        Get
            ServerUDN = MyServerUDN
        End Get
        Set(value As String)
            MyServerUDN = value
        End Set
    End Property
    Property UPnPClass As UPnPClassType
        Get
            UPnPClass = MyUPNPClass
        End Get
        Set(value As UPnPClassType)
            MyUPNPClass = value
        End Set
    End Property
    Property MetaData As String
        Get
            MetaData = MyMetaData
        End Get
        Set(value As String)
            MyMetaData = value
        End Set
    End Property

    Property IconURL As String
        Get
            IconURL = MyIconURL
        End Get
        Set(value As String)
            MyIconURL = value
        End Set
    End Property

    Property AlbumName As String
        Get
            AlbumName = MyAlbumName
        End Get
        Set(value As String)
            MyAlbumName = value
        End Set
    End Property

    Property ArtistName As String
        Get
            ArtistName = MyArtistName
        End Get
        Set(value As String)
            MyArtistName = value
        End Set
    End Property

    Property Genre As String
        Get
            Genre = MyGenre
        End Get
        Set(value As String)
            MyGenre = value
        End Set
    End Property

    Property ObjectPath As String
        Get
            ObjectPath = MyObjectPath
        End Get
        Set(value As String)
            MyObjectPath = value
        End Set
    End Property

End Class
