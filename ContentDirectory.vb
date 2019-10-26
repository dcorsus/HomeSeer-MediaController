Imports System.Xml
Imports ADODB
Imports System.Drawing

Partial Public Class HSPI

    Public WithEvents myContentDirectoryCallback As New myUPnPControlCallback
    Private ContentDirectory As MyUPnPService = Nothing
    Private MyContentDirectoryHSRef As Integer = -1
    Private MySystemUpdateID As String = ""
    Private MySearchCapabilities As String = ""
    Private MySortCapabilities As String = ""
    Private MySearchResult As String = ""
    Private MySearchNumberReturned As Integer = 0
    Private MySearchTotalMatches As Integer = 0
    Private MySearchUpdateID As Integer = 0
    Private MyX_browseResult As String = ""
    Private MyX_browseNumberReturned As Integer = 0
    Private MyX_browseTotalMatches As Integer = 0
    Private MyX_browseUpdateID As Integer = 0
    Private MYX_browseStartingIndex As Integer = 0
    Private MyBrowseResult As String = ""
    Private MyBrowseNumberReturned As Integer = 0
    Private MyBrowseTotalMatches As Integer = 0
    Private MyBrowseUpdateID As Integer = 0
    Private MusicDBIsBeingEstablished As Boolean
    Private MyMaxDepthCounter As Integer = 0
    Private MySearchMaxDepthCounter As Integer
    Private Const MAXLOOPDEPTH = 20
    Private MyReplaceLoopbackIPAddress As Boolean = True
    Private MyRecordUpdateCounter As Integer = 0
    Private SupportSearch As Boolean = False

    'Private MyConnection As ADODB.Connection
    'Private MyRecordset As ADODB.Recordset

    Private SkipTitleCheck As Boolean = True   ' dcoreventissue

    Private Sub ContentDirectoryStateChange(ByVal StateVarName As String, ByVal Value As Object) Handles myContentDirectoryCallback.ControlStateChange 'ContentDirectoryStateChange
        ' ContentDirectoryControlStateChange
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "ContentDirectoryStateChange for UPnPDevice = " & MyUPnPDeviceName & ": Var Name = " & StateVarName & " Value = " & Value.ToString)
        'log( "ContentDirectoryStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString)
        ProcessContentDirectory(StateVarName, Value)
    End Sub

    Private Sub ContentDirectoryDied() Handles myContentDirectoryCallback.ControlDied 'ContentDirectoryDied
        'Log( "Content Directory Callback Died. UPnPDevice - " & MyUPnPDeviceName,LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to UPnPDevice " & MyUPnPDeviceName & " was lost in ContentDirectory.", LogType.LOG_TYPE_WARNING)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR in ContentDirectoryDied for UPnPDevice - " & MyUPnPDeviceName & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub ProcessContentDirectory(VariableName As String, VariableValue As Object)
        ' Sonos
        ' Browseable = Boolean
        ' RecentlyPlayedUpdateID = String
        ' ShareListUpdateID = String
        ' SavedQueuesUpdateID = String
        ' UserRadioUpdateID = String
        ' ShareIndexLastError = String
        ' ShareIndexInProgress = String
        ' ShareListRefreshState = NOTRUN RUNNING DONE
        ' ContainerUpdateIDs = String
        ' SystemUpdateID = UI4

        ' Buffy
        ' TransferIDs
        ' ContainerUpdateIDs
        ' SystemUpdateID
        ' RemoteSharingEnabled
        ' ShareListRefreshState

        Select Case VariableName
            Case "Browseable"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (Browseable) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "RecentlyPlayedUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (RecentlyPlayedUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "ShareListUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (ShareListUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SavedQueuesUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (SavedQueuesUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "RadioLocationUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (RadioLocationUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "RadioFavoritesUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (RadioFavoritesUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "FavoritePresetsUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (FavoritePresetsUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "FavoritesUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (FavoritesUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "UserRadioUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (UserRadioUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "ShareIndexLastError"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (ShareIndexLastError) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "ShareIndexInProgress"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (ShareIndexInProgress) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "ContainerUpdateIDs"
                TreatContainerUpdateIDs(VariableValue.ToString)
            Case "TransferIDs"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (TransferIDs) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SystemUpdateID"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (SystemUpdateID) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                TreatSystemUpdateID(VariableValue.ToString)
            Case "RemoteSharingEnabled"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (RemoteSharingEnabled) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "ShareListRefreshState"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (ShareListRefreshState) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "X_RemoteSharingEnabled"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (X_RemoteSharingEnabled) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
            Case "SearchCapabilities"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (SearchCapabilities) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                MySearchCapabilities = VariableValue
            Case "SortCapabilities"
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log(MyUPnPDeviceName & " received (SortCapabilities) = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
                MySortCapabilities = VariableValue
            Case Else
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning " & MyUPnPDeviceName & " received Not Implemented (" & VariableName & ") = " & VariableValue.ToString, LogType.LOG_TYPE_INFO)
        End Select
    End Sub

    Private Sub TreatSystemUpdateID(SystemUpdateID As String)
        If MySystemUpdateID = "" Then MySystemUpdateID = GetStringIniFile(MyUDN, DeviceInfoIndex.diSystemUpdateID.ToString, "")
        If MySystemUpdateID <> SystemUpdateID Then
            WriteStringIniFile(MyUDN, DeviceInfoIndex.diSystemUpdateID.ToString, SystemUpdateID)
            MySystemUpdateID = SystemUpdateID
            'MyTimeoutActionArray(TOProcessBuildDB) = TOProcessBuildDBValue tobe fixed dcor
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatSystemUpdateID for device - " & MyUPnPDeviceName & " set Update flag", LogType.LOG_TYPE_INFO)
        End If

    End Sub

    Private Sub TreatContainerUpdateIDs(ContainerUpdateIDs As String)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( MyUPnPDeviceName & " received (ContainerUpdateIDs) = " & ContainerUpdateIDs.ToString)
        Try
            Dim ContainerObjects As String() = Split(ContainerUpdateIDs, ",")
            ' the format is ObjectId - comma - updatedID - comma

            Dim NbrofUpdates As Integer = UBound(ContainerObjects, 1)
            If NbrofUpdates > 0 Then NbrofUpdates = NbrofUpdates + 1
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatContainerUpdateIDs for device - " & MyUPnPDeviceName & " received " & NbrofUpdates.ToString & " ContainerUpdateIDs", LogType.LOG_TYPE_INFO)
            'For Index = 0 To (UBound(ContainerObjects, 1) - 1)
            '   Dim ObjectID As String = ContainerObjects(Index)
            '   Dim UpdateID As String = ContainerObjects(Index + 1)
            '   Index = Index + 2
            'Next
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatContainerUpdateIDs for device - " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Function SearchForSearchCapability(Field As String) As Boolean
        SearchForSearchCapability = False
        If MySearchCapabilities = "" Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchForSearchCapability called for Device - " & MyUPnPDeviceName & " with Name = " & Field.ToString, LogType.LOG_TYPE_INFO)
        Dim SearchCaps As String()
        SearchCaps = Split(MySearchCapabilities, ",")
        For Each SearchCap In SearchCaps
            If SearchCap.ToString.ToUpper = Field.ToUpper Then
                SearchForSearchCapability = True
                Exit Function
            End If
        Next
    End Function







    Public Function GetSystemUpdateID(Id As String) As String
        GetSystemUpdateID = ""
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSystemUpdateID called for device " & MyUPnPDeviceName & " and ID = " & Id.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = Id
            ContentDirectory.InvokeAction("GetSystemUpdateID", InArg, OutArg)
            MySystemUpdateID = OutArg(0)
            GetSystemUpdateID = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSystemUpdateID for device = " & MyUPnPDeviceName & " and ID = " & Id.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function Search(ContainerID As String, SearchCriteria As String, Filter As String, StartingIndex As Integer, RequestedCount As Integer, SortCriteria As String) As String
        ' use this to find all tracks 
        ' upnp:class derivedfrom "object.item.audioItem" and @refID exists false
        ' (upnp:class derivedfrom "object.item.audioItem") and (@refID exists false)  and dc:creator contains "ABC"
        ' (upnp:class derivedfrom "object.container") and (@refID exists false)  and dc:title = "genre"
        ' (upnp:class derivedfrom "object.container") and (@refID exists false)  and dc:title = "Artist"
        ' use   dc:creator
        '       dc:title
        '       upnp:album
        '       upnp:genre
        '       upnp:artist
        ' Use object.Container
        '       object.container.person.musicArtist
        ' 
        Search = ""
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Search called for device " & MyUPnPDeviceName & " with ContainerID = " & ContainerID.ToString & " and SearchCriteria = " & SearchCriteria.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(5)
            Dim OutArg(3)
            InArg(0) = ContainerID
            InArg(1) = SearchCriteria
            InArg(2) = Filter
            InArg(3) = StartingIndex
            InArg(4) = RequestedCount
            InArg(5) = SortCriteria
            ContentDirectory.InvokeAction("Search", InArg, OutArg)
            MySearchResult = OutArg(0)
            MySearchNumberReturned = OutArg(1)
            MySearchTotalMatches = OutArg(2)
            MySearchUpdateID = OutArg(3)
            Search = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Search for device = " & MyUPnPDeviceName & " and ContainerID = " & ContainerID.ToString & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetSearchCapabilities() As String
        GetSearchCapabilities = ""
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSearchCapabilities called for device " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            ContentDirectory.InvokeAction("GetSearchCapabilities", InArg, OutArg)
            GetSearchCapabilities = "OK"
            MySearchCapabilities = OutArg(0)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSearchCapabilities for device = " & MyUPnPDeviceName & " returned = " & MySearchCapabilities, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSearchCapabilities for device = " & MyUPnPDeviceName & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetSortCapabilities() As String
        GetSortCapabilities = ""
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSortCapabilities called for device " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            ContentDirectory.InvokeAction("GetSortCapabilities", InArg, OutArg)
            GetSortCapabilities = "OK"
            MySortCapabilities = OutArg(0)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSortCapabilities for device = " & MyUPnPDeviceName & " returned = " & MySortCapabilities, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSortCapabilities for device = " & MyUPnPDeviceName & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function Browse(ObjectID As String, BrowseFlag As String, Filter As String, StartingIndex As Integer, RequestedCount As Integer, SortCriteria As String) As Object
        Browse = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Browse called for device " & MyUPnPDeviceName & " with ObjectID=" & ObjectID & ", Browseflag=" & BrowseFlag & ", filter=" & Filter & ", StartingIndex=" & StartingIndex.ToString & ", requestedcount=" & RequestedCount.ToString & ", Sortcriteria=" & SortCriteria.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(5)
            Dim OutArg(3)
            InArg(0) = ObjectID
            InArg(1) = BrowseFlag
            InArg(2) = Filter
            InArg(3) = StartingIndex
            InArg(4) = RequestedCount
            InArg(5) = SortCriteria
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            MyBrowseResult = OutArg(0)
            MyBrowseNumberReturned = OutArg(1)
            MyBrowseTotalMatches = OutArg(2)
            MyBrowseUpdateID = OutArg(3)
            Browse = OutArg
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Browse for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function X_BrowseByLetter(ObjectID As String, BrowseFlag As String, Filter As String, StartingLetter As String, RequestedCount As Integer, SortCriteria As String) As String
        X_BrowseByLetter = ""
        If DeviceStatus = "Offline" Then Exit Function
        If ContentDirectory Is Nothing Then Exit Function
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("X_BrowseByLetter called for device " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(5)
            Dim OutArg(4)
            InArg(0) = ObjectID
            InArg(1) = BrowseFlag
            InArg(2) = Filter
            InArg(3) = StartingLetter
            InArg(4) = RequestedCount
            InArg(5) = SortCriteria
            ContentDirectory.InvokeAction("X_BrowseByLetter", InArg, OutArg)
            MyX_browseResult = OutArg(0)
            MyX_browseNumberReturned = OutArg(1)
            MyX_browseTotalMatches = OutArg(2)
            MyX_browseUpdateID = OutArg(3)
            MYX_browseStartingIndex = OutArg(4)
            X_BrowseByLetter = "OK"
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in X_BrowseByLetter for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function
    ' 

    ' Sonos
    ' GetSystemUpdateID
    ' GetAlbumArtistDisplayOption
    ' GetLastIndexChange
    ' FindPrefix
    ' GetAllPrefixLocations
    ' CreateObject
    ' UpdateObject
    ' DestroyObject
    ' RefreshShareList
    ' RefreshShareIndex
    ' RequestResort
    ' GetShareIndexInProgress
    ' GetBrowseable

    Public Sub SetBrowsable(Browse As Boolean)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetBrowsable called for device = " & MyUPnPDeviceName & " and Browse flag = " & Browse.ToString, LogType.LOG_TYPE_INFO)
        If ContentDirectory Is Nothing Then Exit Sub
        Dim InArg(0)
        Dim OutArg(0)
        InArg(0) = Browse              ' isBrowseable Boolean
        Try
            ContentDirectory.InvokeAction("SetBrowseable", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetBrowsable for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub InitNewRecordSet(ByRef RecordSet As Object)
        Try
            RecordSet("Id") = ""
            'RecordSet("ParentID") = ""
            'RecordSet("Level") = MyMaxDepthCounter
            'RecordSet("Restricted") = ""
            'RecordSet("Searchable") = ""
            RecordSet("Title") = ""
            RecordSet("Album") = ""
            RecordSet("Artist") = ""
            'RecordSet("Creator") = ""
            'RecordSet("Author") = ""
            RecordSet("Genre") = ""
            'RecordSet("Date") = ""
            'RecordSet("URI") = ""
            'RecordSet("AlbumArtURI") = ""
            'RecordSet("Icon") = ""
            'RecordSet("Res") = ""
            'RecordSet("ProtocolInfo") = ""
            RecordSet("Class") = ""
            RecordSet("TrackNo") = 0
            'RecordSet("Size") = ""
            'RecordSet("Duration") = ""
            'RecordSet("AlbumArtist") = ""
            'RecordSet("Desc") = ""
            'RecordSet("Resolution") = ""
        Catch ex As Exception
            Log("Error in InitNewRecordSet for device = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub PrintRecordSet(ByRef RecordSet As Object)
        Try
            Log("PrintRecordSet (Id)           = " & RecordSet("Id").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (ParentID)     = " & RecordSet("ParentID").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Level)        = " & RecordSet("Level").value, LogType.LOG_TYPE_INFO)
            'log( "PrintRecordSet (Restricted)   = " & RecordSet("Restricted").value, LogType.LOG_TYPE_INFO)
            'log( "PrintRecordSet (Searchable)   = " & RecordSet("Searchable").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Title)        = " & RecordSet("Title").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Album)        = " & RecordSet("Album").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Artist)       = " & RecordSet("Artist").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Creator)      = " & RecordSet("Creator").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Author)       = " & RecordSet("Author").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Genre)        = " & RecordSet("Genre").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Date)         = " & RecordSet("Date").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (URI)          = " & RecordSet("URI").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (AlbumArtURI)  = " & RecordSet("AlbumArtURI").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Icon)         = " & RecordSet("Icon").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Res)          = " & RecordSet("Res").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (ProtocolInfo) = " & RecordSet("ProtocolInfo").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Class)        = " & RecordSet("Class").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (TrackNo)      = " & RecordSet("TrackNo").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Size)         = " & RecordSet("Size").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (Duration)     = " & RecordSet("Duration").value, LogType.LOG_TYPE_INFO)
            Log("PrintRecordSet (AlbumArtist)  = " & RecordSet("AlbumArtist").value, LogType.LOG_TYPE_INFO)
            'log( "PrintRecordSet (Desc)         = " & RecordSet("Desc").value, LogType.LOG_TYPE_INFO)
            'log( "PrintRecordSet (Resolution)   = " & RecordSet("Resolution").value, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in PrintRecordSet for device = " & MyUPnPDeviceName & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Function GetContainerFromServer(ObjectID As String, MetaDataOnly As Boolean) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer called for device - " & MyUPnPDeviceName & " with ObjectID=" & ObjectID & " and MetaDataOnly = " & MetaDataOnly.ToString, LogType.LOG_TYPE_INFO)
        GetContainerFromServer = Nothing
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_INFO)
            Exit Function
        End If

        Dim SList As DBRecord() = Nothing
        Dim LoopIndex As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0
        'Dim ObjectFilter As String = "dc:title,upnp:album,upnp:artist,upnp:genre,upnp:albumArtURI,res"
        Dim ObjectFilter As String = "dc:title"

        Dim InArg(5) As String
        Dim OutArg(3) As String
        InArg(0) = ObjectID                 ' Object ID     String 
        If MetaDataOnly Then
            InArg(1) = "BrowseMetadata"         ' Browse Flag   String
        Else
            InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        End If
        InArg(2) = ObjectFilter             ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String


        StartIndex = 0
        Dim RecordIndex As Integer = 0

        Dim StartTime As DateTime = DateTime.Now
        Dim elapsed_time As TimeSpan

        Do

            InArg(3) = StartIndex               ' Index         UI4

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in GetContainerFromServer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)

            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer found " & TotalMatches.ToString & " and returned " & NumberReturned.ToString & " entries, at Index = " & StartIndex.ToString & " for device = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_INFO)

            If NumberReturned = 0 Then
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End If

            Dim xmlData As XmlDocument = New XmlDocument
            Try
                xmlData.LoadXml(OutArg(0).ToString)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("XML=" & OutArg(0).ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in GetContainerFromServer at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End Try
            Try
                ReDim Preserve SList(StartIndex + NumberReturned - 1)
            Catch ex As Exception
                Log("Error in GetContainerFromServer re-dimensioning the List Array for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End Try
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetContainerFromServer Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO) ' this starts with <Event>
                If PIDebuglevel > DebugLevel.dlEvents Then Log("GetContainerFromServer Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO) ' this starts with <Event> 
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                Dim WaitLoopCounter As Integer = 0
                For Each outerNode As XmlNode In nodelist
                    If UCase(outerNode.Name) = "CONTAINER" Then
                        Dim NewRecord = New DBRecord
                        Dim ObjectClass As String() = Nothing
                        Dim ClassInfo As String = ""
                        NewRecord.Id = ""
                        NewRecord.ParentID = ""
                        NewRecord.Title = ""
                        NewRecord.AlbumName = ""
                        NewRecord.ArtistName = ""
                        NewRecord.Genre = ""
                        NewRecord.IconURL = ""
                        Try
                            NewRecord.Id = outerNode.Attributes("id").Value
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ParentID = outerNode.Attributes("parentID").Value
                        Catch ex As Exception
                        End Try

                        Try
                            NewRecord.Title = outerNode.Item("dc:title").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            ClassInfo = outerNode.Item("upnp:class").InnerText
                            NewRecord.ClassType = ClassInfo
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ItemOrContainer = "CONTAINER"
                        Catch ex As Exception
                        End Try

                        'If SList Is Nothing Then
                        'ReDim SList(0)
                        'SList(0) = NewRecord
                        'Else
                        'ReDim Preserve SList(RecordIndex)
                        SList(RecordIndex) = NewRecord
                        'End If
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("Container Record #" & RecordIndex.ToString & " with value " & NewRecord.Title & " and ObjectID = " & NewRecord.Id, LogType.LOG_TYPE_INFO)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Container Record #" & RecordIndex.ToString & " with value " & NewRecord.Title & " and ObjectID = " & NewRecord.Id, LogType.LOG_TYPE_INFO) 
                        RecordIndex = RecordIndex + 1
                    ElseIf UCase(outerNode.Name) = "ITEM" Then
                        Dim NewRecord = New DBRecord
                        Dim ObjectClass As String() = Nothing
                        NewRecord.Id = ""
                        NewRecord.ParentID = ""
                        NewRecord.Title = ""
                        NewRecord.AlbumName = ""
                        NewRecord.ArtistName = ""
                        NewRecord.Genre = ""
                        NewRecord.IconURL = ""
                        Dim ClassInfo As String = ""
                        Try
                            NewRecord.Id = outerNode.Attributes("id").Value
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.Title = outerNode.Item("dc:title").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            ClassInfo = outerNode.Item("upnp:class").InnerText
                            NewRecord.ClassType = ClassInfo
                        Catch ex As Exception
                        End Try
                        If ProcessClassInfo(NewRecord.ClassType) = UPnPClassType.ctPictures Then
                            Try
                                'NewRecord.IconURL = outerNode.Item("res").InnerText
                            Catch ex As Exception
                            End Try
                        Else
                            Try
                                'NewRecord.IconURL = outerNode.Item("upnp:albumArtURI").InnerText
                            Catch ex As Exception
                            End Try
                        End If
                        If NewRecord.IconURL = "yyyy" Then ' I guess this still needs to be fixed dcor
                            Dim IConImage As Bitmap
                            Try
                                IConImage = New Bitmap(NewRecord.IconURL)
                            Catch ex As Exception
                                'just use the no art image
                                IConImage = New Bitmap(CurrentAppPath & "/html" & NoArtPath)
                                Log("Error in GetContainerFromServer for device - " & MyUPnPDeviceName & " getting imageart with error = " & ex.Message & " Path = " & MyIConURL.ToString, LogType.LOG_TYPE_ERROR)
                            End Try
                            'Dim IConImage As Image
                            'IConImage. = 150
                            Try
                                Dim propItem As System.Drawing.Imaging.PropertyItem
                                For Each propItem In IConImage.PropertyItems
                                    Log("Picture info in GetContainerFromServer for device - " & MyUPnPDeviceName & " Id = " & propItem.Id.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & MyUPnPDeviceName & " Type = " & propItem.Type.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & MyUPnPDeviceName & " Value = " & propItem.Value.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & MyUPnPDeviceName & " Length = " & propItem.Len.ToString, LogType.LOG_TYPE_INFO)
                                Next
                            Catch ex As Exception
                                Log("Error in GetContainerFromServer for device - " & MyUPnPDeviceName & " getting imageinfo with error = " & ex.Message & " Path = " & MyIConURL.ToString, LogType.LOG_TYPE_ERROR)
                            End Try

                            IConImage.Dispose()


                            'GetPicture(MyIconURL)
                        End If
                        Try
                            NewRecord.ItemOrContainer = "ITEM"
                        Catch ex As Exception
                        End Try
                        SList(RecordIndex) = NewRecord
                        'End If
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("Item Record #" & RecordIndex.ToString & " with value " & NewRecord.Title & " and ObjectID = " & NewRecord.Id, LogType.LOG_TYPE_INFO)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Item Record #" & RecordIndex.ToString & " with value " & NewRecord.Title & " and ObjectID = " & NewRecord.Id, LogType.LOG_TYPE_INFO) 
                        RecordIndex = RecordIndex + 1
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetContainerFromServer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " processing Childnodes found node = " & outerNode.Name.ToString, LogType.LOG_TYPE_ERROR)
                    End If
                    WaitLoopCounter += 1
                    If WaitLoopCounter >= 100 Then
                        WaitLoopCounter = 0
                        elapsed_time = DateTime.Now.Subtract(StartTime)
                        If elapsed_time.TotalSeconds > MaxWaitTimeRetrievingContainer Then
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " exceeded maximum wait time", LogType.LOG_TYPE_WARNING)
                            Exit Do
                        End If
                    End If
                Next
            Catch ex As Exception
                Log("Error in GetContainerFromServer at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " processing Childnodes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            'hs.WaitEvents()

            elapsed_time = DateTime.Now.Subtract(StartTime)
            If elapsed_time.TotalSeconds > MaxWaitTimeRetrievingContainer Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " exceeded maximum wait time", LogType.LOG_TYPE_WARNING)
                Exit Do
            End If

        Loop
        Try
            ReDim Preserve SList(RecordIndex - 1)
        Catch ex As Exception

        End Try
        GetContainerFromServer = SList
        SList = Nothing

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerFromServer for device - " & MyUPnPDeviceName & " returned " & RecordIndex.ToString & " Server records ", LogType.LOG_TYPE_INFO)

    End Function

    Public Function GetObjectFromServer(ObjectID As String, DCTitle As String) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetObjectFromServer called for device - " & MyUPnPDeviceName & " with ObjectID=" & ObjectID, LogType.LOG_TYPE_INFO)
        GetObjectFromServer = Nothing
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetObjectFromServer called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        Dim SList As DBRecord() = Nothing
        Dim LoopIndex As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0
        'Dim ObjectFilter As String = "dc:title,upnp:album,upnp:artist,upnp:genre,upnp:albumArtURI,res"
        Dim ObjectFilter As String = "dc:title"

        Dim InArg(5) As String
        Dim OutArg(3) As String
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseMetadata"         ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in GetObjectFromServer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            GetObjectFromServer = SList
            SList = Nothing
            Exit Function
        End Try

        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(OutArg(0).ToString)
        Catch ex As Exception
            Log("Error in GetObjectFromServer loading the MetaData for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  with Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
            GetObjectFromServer = SList
            SList = Nothing
            Exit Function
        End Try


        If DCTitle <> "" Then
            Try
                If Trim(UCase(DCTitle)) <> Trim(UCase(xmlData.GetElementsByTagName("dc:title").Item(0).InnerText)) Then
                    GetObjectFromServer = SList
                    SList = Nothing
                    Exit Function
                End If
            Catch ex As Exception
                GetObjectFromServer = SList
                SList = Nothing
                Exit Function
            End Try
        End If

        ' so either there is no Title specified or they are equal, which means, just return this object!
        Try
            If xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value = ObjectID Then
                ' this is a container
                GetObjectFromServer = GetContainerFromServer(ObjectID, False)
                Exit Function
            End If

        Catch ex As Exception

        End Try
        Try
            ReDim SList(0)
        Catch ex As Exception
            Log("Error in GetObjectFromServer re-dimensioning the List Array for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            GetObjectFromServer = SList
            SList = Nothing
            Exit Function
        End Try

        Try
            Dim NewRecord = New DBRecord
            Dim ObjectClass As String() = Nothing
            Dim ClassInfo As String = ""
            NewRecord.Id = ""
            NewRecord.ParentID = ""
            NewRecord.Title = ""
            NewRecord.AlbumName = ""
            NewRecord.ArtistName = ""
            NewRecord.Genre = ""
            NewRecord.IconURL = ""
            Try
                NewRecord.Id = ObjectID
            Catch ex As Exception
            End Try
            Try
                NewRecord.Title = DCTitle
            Catch ex As Exception
            End Try
            Try
                ClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
                NewRecord.ClassType = ClassInfo
            Catch ex As Exception
            End Try
            Try
                NewRecord.ItemOrContainer = "ITEM"
            Catch ex As Exception
            End Try
            SList(0) = NewRecord
        Catch ex As Exception
            Log("Error in GetObjectFromServer filling record for UPnPDevice = " & MyUPnPDeviceName & ", ObjectID = " & ObjectID.ToString & ", Title = " & DCTitle & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        GetObjectFromServer = SList
        SList = Nothing

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetObjectFromServer for device - " & MyUPnPDeviceName & " returned 1 Item", LogType.LOG_TYPE_INFO)

    End Function


    Public Function GetContainerAtLevel(PlayerUDN As String, StartingLevel As Integer) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerAtLevel called for device - " & MyUPnPDeviceName & " with UDN = " & PlayerUDN & " and StartingLevel = " & StartingLevel.ToString, LogType.LOG_TYPE_INFO)
        GetContainerAtLevel = Nothing
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerAtLevel called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If

        ' we'll use the values stored in the .ini file to pull up the right container

        Dim LevelIndex As Integer = 0
        Dim ObjectID As String = ""
        Dim ObjectTitle As String = ""
        Dim ObjectFilter As String = "dc:title"

        If StartingLevel <> 0 Then
            ObjectID = GetStringIniFile("DevicePage", PlayerUDN & "_" & "ObjectIDlevel" & StartingLevel.ToString, "")
            If ObjectID <> "" Then
                LevelIndex = StartingLevel ' start here! Else start at 0
            End If
        End If

        Do
            Dim StartIndex As Integer = 0
            Dim NumberReturned As Integer = 0
            Dim TotalMatches As Integer = 0

            If LevelIndex = 0 Then
                ObjectID = "0" ' everything starts here
            End If

            LevelIndex = LevelIndex + 1
            If LevelIndex > 10 Then
                Log("Error in GetContainerAtLevel at level = " & LevelIndex.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " max level exceeded", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If

            ObjectTitle = GetStringIniFile("DevicePage", PlayerUDN & "_" & "level" & LevelIndex.ToString, "")
            If ObjectTitle = "" Then
                ' we're at the end
                GetContainerAtLevel = GetContainerFromServer(ObjectID, False)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerAtLevel for device - " & MyUPnPDeviceName & " reached end at level = " & LevelIndex.ToString, LogType.LOG_TYPE_INFO)
                Exit Function
            End If

            Dim InArg(5) As String
            Dim OutArg(3) As String
            InArg(0) = ObjectID                 ' Object ID     String 
            InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
            InArg(2) = ObjectFilter             ' Filter        String
            InArg(3) = 0                        ' Index         UI4
            InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
            InArg(5) = ""                       ' Sort Criteria String

            StartIndex = 0

            Do
                InArg(3) = StartIndex               ' Index         UI4
                InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all

                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in GetContainerAtLevel for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                TotalMatches = OutArg(2)
                NumberReturned = OutArg(1)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerAtLevel for device = " & MyUPnPDeviceName & " retrieved ObjectID = " & ObjectID.ToString & " and got " & TotalMatches.ToString & " matches", LogType.LOG_TYPE_INFO)
                If NumberReturned = 0 Then
                    Log("ERROR in GetContainerAtLevel for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and zero returned items", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                If TotalMatches = 0 Then
                    Log("ERROR in GetContainerAtLevel for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and no matches for title = " & ObjectTitle, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If

                Dim xmlData As XmlDocument = New XmlDocument
                Try
                    xmlData.LoadXml(OutArg(0).ToString)
                    'log( "XML=" & Value.ToString) ' used for testing
                Catch ex As Exception
                    Log("Error in GetContainerAtLevel at level = " & LevelIndex.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try

                Dim ItemIndex As Integer = 0
                Dim Found As Boolean = False

                For ItemIndex = 1 To NumberReturned
                    Try
                        If xmlData.GetElementsByTagName("dc:title").Item(ItemIndex - 1).InnerText = ObjectTitle Then
                            ' OK this is mine
                            ObjectID = xmlData.GetElementsByTagName("container").Item(ItemIndex - 1).Attributes("id").Value
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetContainerAtLevel at level = " & LevelIndex.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " found title = " & ObjectTitle, LogType.LOG_TYPE_INFO)
                            Found = True
                            Exit Do
                        End If
                    Catch ex As Exception
                        Log("Error in GetContainerAtLevel at level = " & LevelIndex.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " tag title or id wasn't found with Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                Next

                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Log("Error in GetContainerAtLevel at level = " & LevelIndex.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " and Title = " & ObjectTitle & " wasn't found", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                wait(0.25)
            Loop
        Loop

    End Function

    Public Function NavigateIntoContainer(NavigationObject As String, NavigationString As String, ByRef inNavigationPart As String, ByRef inEndOfTree As Boolean) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer called for device - " & MyUPnPDeviceName & " and NavigationObject = " & NavigationObject & " and NavigationString = " & NavigationString & " and inEndofTree= " & inEndOfTree.ToString, LogType.LOG_TYPE_INFO)
        NavigateIntoContainer = Nothing
        inNavigationPart = ""
        inEndOfTree = False

        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If


        Dim ObjectNavigationParts As String() = Nothing
        Dim ObjectID As String = ""
        Dim IsRoot As Boolean = False

        If NavigationString = "" Then   ' this is the start, should I set myself to "root" ????
            IsRoot = True
            ObjectNavigationParts = {";::--0"}
        Else
            ObjectNavigationParts = Split(NavigationString, ";--::")
            If ObjectNavigationParts.Count = 0 Then
                IsRoot = True
                ObjectNavigationParts = {";::--0"}
            End If
        End If

        Dim ObjectTitle As String = ""

        Dim NavigationStringIndex As Integer = 0

        Try
            For Each ObjectNavigationPart As String In ObjectNavigationParts
                If ObjectNavigationPart IsNot Nothing Then
                    Dim ObjectParts As String() = Nothing
                    ObjectParts = Split(ObjectNavigationPart, ";::--")
                    If ObjectParts IsNot Nothing Then
                        ObjectTitle = ObjectParts(0)
                        If ObjectParts.Count >= 2 Then
                            ObjectID = ObjectParts(1)
                        Else
                            ObjectID = "0"
                        End If
                        Dim StartIndex As Integer = 0
                        Dim NumberReturned As Integer = 0
                        Dim TotalMatches As Integer = 0
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " is retrieving ObjectID = " & ObjectID & " and looking for ObjectTitle = " & ObjectTitle, LogType.LOG_TYPE_WARNING)
                        Dim InArg(5) As String
                        Dim OutArg(3) As String
                        InArg(0) = ObjectID                 ' Object ID     String 
                        InArg(1) = "BrowseMetadata"         ' Browse Flag   String
                        InArg(2) = "*"                      ' Filter        String
                        InArg(3) = 0                        ' Index         UI4
                        InArg(4) = 0                        ' Count         UI4  - 0 means all
                        InArg(5) = ""                       ' Sort Criteria String

                        Try
                            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                        Catch ex As Exception
                            Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End Try
                        TotalMatches = Val(OutArg(2))
                        NumberReturned = Val(OutArg(1))
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " retrieved for ObjectID = " & ObjectID & " TotalMatches = " & TotalMatches.ToString & " and NbrReturned = " & NumberReturned.ToString, LogType.LOG_TYPE_WARNING)
                        If NumberReturned = 0 Then
                            Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and zero returned items", LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End If
                        If TotalMatches = 0 Then
                            Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and no matches for NavigationString = " & NavigationString, LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End If
                        If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " retrieved MetaData for ObjectID = " & ObjectID & " XML = " & OutArg(0).ToString, LogType.LOG_TYPE_WARNING)

                        Dim xmlData As XmlDocument = New XmlDocument
                        Try
                            xmlData.LoadXml(OutArg(0).ToString)
                        Catch ex As Exception
                            Log("Error in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End Try

                        Dim TitleInXML As String = ""
                        Try
                            TitleInXML = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & ". No Title Tag found in MetaData. Not GOOD!! XML = " & OutArg(0).ToString.ToString, LogType.LOG_TYPE_WARNING)
                        End Try
                        Dim UPnPClass As String = ""
                        Try
                            UPnPClass = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
                        Catch ex As Exception
                            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & ". No Class Tag found in MetaData. Not GOOD!! XML = " & OutArg(0).ToString.ToString, LogType.LOG_TYPE_WARNING)
                        End Try
                        Try
                            If IsRoot Then ' added the NavigationObject = ""  here to fix issue with PLEX server having root with dc:title = empty
                                ' this must be the root, pick up the title and return it!
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " hit root and exiting function", LogType.LOG_TYPE_WARNING)
                                ObjectTitle = TitleInXML
                                NavigateIntoContainer = GetContainerFromServer("0", False)
                                inNavigationPart = ObjectTitle & ";::--" & ObjectID
                                Exit Function
                            ElseIf (TitleInXML = ObjectTitle) Or SkipTitleCheck Then ' dcoreventissue
                                ' OK this is mine
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " found ObjectTitle = " & ObjectTitle & " and UPnPClass = " & UPnPClass, LogType.LOG_TYPE_INFO)
                                Dim ObjectClass As String() = Split(UPnPClass, ".")
                                If Not ObjectClass Is Nothing Then
                                    If UBound(ObjectClass) > 0 Then
                                        If (ObjectClass(1).ToUpper = "ITEM") Or (ObjectClass(1).ToUpper = "ITEMOBJECT") Then ' the itemObject is in my opinion non-standard SONOS
                                            NavigateIntoContainer = GetContainerFromServer(ObjectID, True)
                                            inEndOfTree = True
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                Log("Error in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " Navigation didn't work for NavigationString = " & NavigationString.ToString, LogType.LOG_TYPE_ERROR)
                                Exit Function
                            End If
                        Catch ex As Exception
                            Log("Error in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and NavigationString = " & NavigationString.ToString & " tag title or id wasn't found with Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in NavigateIntoContainer for device - " & MyUPnPDeviceName & "  with NavigationString  = " & NavigationString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ' Ok the navigation panned out, Objectid should now hold the latest Object
        If NavigationObject = "" Then ' either start @ root, up-level or end-of-selection has the NavigationObject set to empty
            ' we're at the end
            NavigateIntoContainer = GetContainerFromServer(ObjectID, False)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " reached end successfully", LogType.LOG_TYPE_INFO)
            Exit Function
        End If

        Try

            Dim StartIndex As Integer = 0
            Dim NumberReturned As Integer = 0
            Dim TotalMatches As Integer = 0
            Dim ObjectFilter As String = "dc:title"

            Dim InArg(5) As String
            Dim OutArg(3) As String
            InArg(0) = ObjectID                 ' Object ID     String 
            InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
            InArg(2) = ObjectFilter             ' Filter        String
            InArg(3) = 0                        ' Index         UI4
            InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
            InArg(5) = ""                       ' Sort Criteria String

            StartIndex = 0

            Do
                InArg(3) = StartIndex               ' Index         UI4
                InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all

                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                TotalMatches = OutArg(2)
                NumberReturned = OutArg(1)
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer for device = " & MyUPnPDeviceName & " retrieved ObjectID = " & ObjectID.ToString & " and got " & TotalMatches.ToString & " matches" & " and NbrReturned = " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
                If NumberReturned = 0 Then
                    Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and zero returned items", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                If TotalMatches = 0 Then
                    Log("ERROR in NavigateIntoContainer for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and no matches for NavigationObject = " & NavigationObject, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " retrieved ChildData for ObjectID = " & ObjectID & " XML = " & OutArg(0).ToString, LogType.LOG_TYPE_WARNING)

                Dim xmlData As XmlDocument = New XmlDocument
                Try
                    xmlData.LoadXml(OutArg(0).ToString)
                Catch ex As Exception
                    Log("Error in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try

                Dim ItemIndex As Integer = 0
                Dim Found As Boolean = False

                For ItemIndex = 1 To NumberReturned
                    Try
                        If xmlData.GetElementsByTagName("dc:title").Item(ItemIndex - 1).InnerText = NavigationObject Then
                            ' OK this is what we are looking for
                            Dim UPnPClass As String = xmlData.GetElementsByTagName("upnp:class").Item(ItemIndex - 1).InnerText
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " for dc:title = " & NavigationObject & " found upnp:class = " & UPnPClass.ToString, LogType.LOG_TYPE_WARNING)
                            Dim ObjectClass As String() = Split(UPnPClass, ".")
                            If Not ObjectClass Is Nothing Then
                                If UBound(ObjectClass) > 0 Then
                                    inEndOfTree = (ObjectClass(1).ToUpper = "ITEM") Or (ObjectClass(1).ToUpper = "ITEMOBJECT")
                                Else
                                    inEndOfTree = False
                                End If
                            End If
                            ObjectID = ""
                            If inEndOfTree Then
                                Try
                                    ObjectID = xmlData.GetElementsByTagName("item").Item(ItemIndex - 1).Attributes("id").Value
                                Catch ex As Exception
                                End Try
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " for dc:title = " & NavigationObject & " found item ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_WARNING)
                            Else
                                Try
                                    ObjectID = xmlData.GetElementsByTagName("container").Item(ItemIndex - 1).Attributes("id").Value
                                Catch ex As Exception
                                End Try
                                If PIDebuglevel > DebugLevel.dlEvents Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " for dc:title = " & NavigationObject & " found container ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_WARNING)
                            End If
                            If ObjectID <> "" Then
                                inNavigationPart = NavigationObject & ";::--" & ObjectID
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " found NavigationObject = " & NavigationObject, LogType.LOG_TYPE_INFO)
                                Found = True
                                Exit Do
                            End If
                        End If
                    Catch ex As Exception
                        Log("Error in NavigateIntoContainer1 for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " tag title or id wasn't found with Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                Next

                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Log("Error in NavigateIntoContainer for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & " and NavigationObject = " & NavigationObject & " wasn't found", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                wait(0.25)
            Loop
        Catch ex As Exception
            Log("Error in NavigateIntoContainer for device - " & MyUPnPDeviceName & "  with NavigationString  = " & NavigationString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        If ObjectID <> "" Then
            ' we're at the end
            NavigateIntoContainer = GetContainerFromServer(ObjectID, inEndOfTree)
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainer for device - " & MyUPnPDeviceName & " reached end successfully", LogType.LOG_TYPE_INFO)
        End If

    End Function

    Public Function NavigateIntoContainerByTitle(NavigationString As String) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle called for device - " & MyUPnPDeviceName & " and NavigationString = " & NavigationString, LogType.LOG_TYPE_INFO)
        NavigateIntoContainerByTitle = Nothing

        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If


        Dim ObjectNavigationParts As String() = Nothing
        Dim ObjectID As String = "0"

        If NavigationString = "" Then
            ObjectNavigationParts = {";::--0"}
        Else
            ObjectNavigationParts = Split(NavigationString, ";--::")
            If ObjectNavigationParts Is Nothing Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle called for device - " & MyUPnPDeviceName & " but no ObjectNavigation parts in NavigationString = " & NavigationString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            If ObjectNavigationParts.Count = 0 Then
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle called for device - " & MyUPnPDeviceName & " but no ObjectNavigation parts in NavigationString = " & NavigationString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        End If

        Dim ObjectTitle As String = ""
        Dim IsContainer As Boolean = False
        Dim ClassType As String = ""

        Try
            For Each ObjectNavigationPart As String In ObjectNavigationParts
                If ObjectNavigationPart IsNot Nothing Then
                    Dim ObjectParts As String() = Nothing
                    ObjectParts = Split(ObjectNavigationPart, ";::--")
                    If ObjectParts IsNot Nothing Then
                        ObjectTitle = ObjectParts(0)
                        If UCase(ObjectTitle) = "ROOT" Then GoTo NextPart ' skip the first
                        Try

                            Dim StartIndex As Integer = 0
                            Dim NumberReturned As Integer = 0
                            Dim TotalMatches As Integer = 0
                            Dim ObjectFilter As String = "dc:title"

                            Dim InArg(5) As String
                            Dim OutArg(3) As String
                            InArg(0) = ObjectID                 ' Object ID     String 
                            InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
                            InArg(2) = ObjectFilter             ' Filter        String
                            InArg(3) = 0                        ' Index         UI4
                            InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
                            InArg(5) = ""                       ' Sort Criteria String

                            StartIndex = 0

                            Do
                                InArg(3) = StartIndex               ' Index         UI4
                                InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all

                                Try
                                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                                Catch ex As Exception
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ERROR in NavigateIntoContainerByTitle for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    Exit Function
                                End Try
                                TotalMatches = OutArg(2)
                                NumberReturned = OutArg(1)
                                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle for device = " & MyUPnPDeviceName & " retrieved ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " and got " & TotalMatches.ToString & " matches", LogType.LOG_TYPE_INFO)
                                If NumberReturned = 0 Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ERROR in NavigateIntoContainerByTitle for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " and zero returned items", LogType.LOG_TYPE_ERROR)
                                    Exit Function
                                End If
                                If TotalMatches = 0 Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ERROR in NavigateIntoContainerByTitle for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle, LogType.LOG_TYPE_ERROR)
                                    Exit Function
                                End If

                                Dim xmlData As XmlDocument = New XmlDocument
                                Try
                                    xmlData.LoadXml(OutArg(0).ToString)
                                Catch ex As Exception
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in NavigateIntoContainerByTitle for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                                    Exit Function
                                End Try

                                Dim ItemIndex As Integer = 0
                                Dim Found As Boolean = False

                                For ItemIndex = 1 To NumberReturned
                                    Try
                                        If xmlData.GetElementsByTagName("dc:title").Item(ItemIndex - 1).InnerText = ObjectTitle Then
                                            ' OK this is what we are looking for
                                            Try
                                                ObjectID = xmlData.GetElementsByTagName("container").Item(ItemIndex - 1).Attributes("id").Value
                                                If ObjectID <> "" Then
                                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle for UPnPDevice = " & MyUPnPDeviceName & " found container with ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " for NavigationString = " & NavigationString, LogType.LOG_TYPE_INFO)
                                                    IsContainer = True
                                                    Exit Do
                                                End If
                                            Catch ex As Exception
                                            End Try
                                            Try
                                                ObjectID = xmlData.GetElementsByTagName("item").Item(ItemIndex - 1).Attributes("id").Value
                                                If ObjectID <> "" Then
                                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle for UPnPDevice = " & MyUPnPDeviceName & " found Item with ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " for NavigationString = " & NavigationString, LogType.LOG_TYPE_INFO)
                                                    IsContainer = False
                                                    Try
                                                        ClassType = xmlData.GetElementsByTagName("upnp:class").Item(ItemIndex - 1).InnerText
                                                    Catch ex As Exception
                                                    End Try
                                                    Exit Do
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        End If
                                    Catch ex As Exception
                                    End Try
                                Next

                                StartIndex = StartIndex + NumberReturned
                                If StartIndex >= TotalMatches Then
                                    If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in NavigateIntoContainerByTitle for UPnPDevice = " & MyUPnPDeviceName & ", ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & " and NavigationString = " & NavigationString & " wasn't found", LogType.LOG_TYPE_ERROR)
                                    Exit Function
                                End If
                                wait(0.25)
                            Loop
                        Catch ex As Exception
                            Log("Error in NavigateIntoContainerByTitle for device - " & MyUPnPDeviceName & ", ObjectID = " & ObjectID.ToString & ", ObjectTitle = " & ObjectTitle & "  with NavigationString  = " & NavigationString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Function
                        End Try

                    End If
                End If
NextPart:
            Next
        Catch ex As Exception
            Log("Error in NavigateIntoContainerByTitle for device - " & MyUPnPDeviceName & "  with NavigationString  = " & NavigationString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ' Ok the navigation panned out, Objectid should now hold the latest Object
        If ObjectID <> "" Then
            ' we're at the end
            If IsContainer Then
                NavigateIntoContainerByTitle = GetContainerFromServer(ObjectID, False)
            Else
                Dim NewRecord = New DBRecord
                Dim ObjectClass As String() = Nothing
                Dim ClassInfo As String = ""
                NewRecord.Id = ObjectID
                NewRecord.ParentID = ""
                NewRecord.Title = ObjectTitle
                NewRecord.AlbumName = ""
                NewRecord.ArtistName = ""
                NewRecord.Genre = ""
                NewRecord.IconURL = ""
                NewRecord.ClassType = ClassType
                NewRecord.ItemOrContainer = "ITEM"
                Dim SList As DBRecord() = Nothing
                ReDim SList(0)
                SList(0) = NewRecord
                NavigateIntoContainerByTitle = SList
                SList = Nothing
            End If

            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigateIntoContainerByTitle for device - " & MyUPnPDeviceName & " reached end successfully", LogType.LOG_TYPE_INFO)
            Exit Function
        End If

    End Function

    Private Sub PrintOutXML(inXMLData As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(inXMLData)
            'log( "XML=" & Value.ToString) ' used for testing
        Catch ex As Exception
            Log("Error in PrintOutXML at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " loading XML. Error = " & ex.Message & ". XML = " & inXMLData.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            'Get a list of all the child elements
            Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
            For Each outerNode As XmlNode In nodelist
                For Each Attr As XmlAttribute In outerNode.Attributes
                    Log("PrintOutXML for Device = " & MyUPnPDeviceName & " (" & Attr.Name.ToString & ") = " & Attr.Value.ToString, LogType.LOG_TYPE_INFO)
                Next
                For Each InnerNode As XmlNode In outerNode.ChildNodes
                    Log("PrintOutXML for Device = " & MyUPnPDeviceName & " (" & InnerNode.Name.ToString & ") = " & InnerNode.InnerText.ToString, LogType.LOG_TYPE_INFO)
                    For Each Attr As XmlAttribute In InnerNode.Attributes
                        Log("PrintOutXML for Device = " & MyUPnPDeviceName & " (" & Attr.Name.ToString & ") = " & Attr.Value.ToString, LogType.LOG_TYPE_INFO)
                    Next
                Next
            Next
        Catch ex As Exception
            Log("Error in PrintOutXML for UPnPDevice = " & MyUPnPDeviceName & " processing Childnodes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function SearchObjects(ByVal ObjectID As String, SearchString As String, SearchOperator As SearchOperatorTypes, Optional ByVal DownToItems As Boolean = False) As DBRecord()
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchObjects called for device - " & MyUPnPDeviceName & " with ObjectID=" & ObjectID & " and SearchString = " & SearchString & " and Searchoperator = " & SearchOperator.ToString & " and DownToItems = " & DownToItems.ToString, LogType.LOG_TYPE_INFO)
        SearchObjects = Nothing
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchObjects called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If

        Dim SList As DBRecord() = Nothing
        Dim LoopIndex As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0

        Dim InArg(5) As String
        Dim OutArg(3) As String
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 1                        ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SearchObjects/Browse for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchObjects found " & TotalMatches.ToString & " entries for device = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_INFO)
        If TotalMatches = 0 Then Exit Function

        StartIndex = 0
        Dim RecordIndex As Integer = 0
        Do
            InArg(3) = StartIndex               ' Index         UI4
            InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in SearchObjects for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                SearchObjects = SList
                SList = Nothing
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            If NumberReturned = 0 Then
                SearchObjects = SList
                SList = Nothing
                Exit Function
            End If


            Dim xmlData As XmlDocument = New XmlDocument
            Try
                xmlData.LoadXml(OutArg(0).ToString)
                'log( "XML=" & Value.ToString) ' used for testing
            Catch ex As Exception
                Log("Error in SearchObjects at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                SearchObjects = SList
                SList = Nothing
                Exit Function
            End Try
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchObjects Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchObjects Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    If UCase(outerNode.Name) = "CONTAINER" Then
                        Dim RecordTitle As String = ""
                        Try
                            RecordTitle = outerNode.Item("dc:title").InnerText
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchObjects found a container with title: " & RecordTitle)
                        Catch ex As Exception
                        End Try
                        If SearchInString(RecordTitle, SearchString, SearchOperator) Then
                            Dim NewRecord = New DBRecord
                            Dim ObjectClass As String() = Nothing
                            Dim ClassInfo As String = ""
                            NewRecord.Id = ""
                            NewRecord.ParentID = ""
                            NewRecord.Title = RecordTitle
                            NewRecord.AlbumName = ""
                            NewRecord.ArtistName = ""
                            NewRecord.Genre = ""
                            NewRecord.IconURL = ""
                            Try
                                NewRecord.Id = outerNode.Attributes("id").Value
                            Catch ex As Exception
                            End Try
                            Try
                                NewRecord.ParentID = outerNode.Attributes("parentID").Value
                            Catch ex As Exception
                            End Try
                            Try
                                ClassInfo = outerNode.Item("upnp:class").InnerText
                                NewRecord.ClassType = ClassInfo
                            Catch ex As Exception
                            End Try
                            Try
                                NewRecord.ItemOrContainer = "CONTAINER"
                            Catch ex As Exception
                            End Try
                            If SList Is Nothing Then
                                ReDim SList(0)
                                SList(0) = NewRecord
                            Else
                                ReDim Preserve SList(RecordIndex)
                                SList(RecordIndex) = NewRecord
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("Record #" & RecordIndex.ToString & " with value " & NewRecord.Title, LogType.LOG_TYPE_INFO)
                            RecordIndex = RecordIndex + 1
                        End If
                    ElseIf UCase(outerNode.Name) = "ITEM" Then
                        Dim RecordTitle As String = ""
                        Try
                            RecordTitle = outerNode.Item("dc:title").InnerText
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchObjects found an Item with title: " & RecordTitle)
                        Catch ex As Exception
                        End Try
                        If SearchInString(RecordTitle, SearchString, SearchOperator) Then
                            Dim NewRecord = New DBRecord
                            Dim ObjectClass As String() = Nothing
                            NewRecord.Id = ""
                            NewRecord.ParentID = ""
                            NewRecord.Title = RecordTitle
                            NewRecord.AlbumName = ""
                            NewRecord.ArtistName = ""
                            NewRecord.Genre = ""
                            NewRecord.IconURL = ""
                            Dim ClassInfo As String = ""
                            Try
                                NewRecord.Id = outerNode.Attributes("id").Value
                            Catch ex As Exception
                            End Try
                            Try
                                NewRecord.ParentID = outerNode.Attributes("parentID").Value
                            Catch ex As Exception
                            End Try
                            Try
                                ClassInfo = outerNode.Item("upnp:class").InnerText
                                NewRecord.ClassType = ClassInfo
                            Catch ex As Exception
                            End Try
                            Try
                                NewRecord.ItemOrContainer = "ITEM"
                            Catch ex As Exception
                            End Try
                            If SList Is Nothing Then
                                ReDim SList(0)
                                SList(0) = NewRecord
                            Else
                                ReDim Preserve SList(RecordIndex)
                                SList(RecordIndex) = NewRecord
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("Record #" & RecordIndex.ToString & " with value " & NewRecord.Title, LogType.LOG_TYPE_INFO)
                            RecordIndex = RecordIndex + 1
                        End If
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SearchObjects for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes found node = " & outerNode.Name.ToString, LogType.LOG_TYPE_ERROR)
                    End If
                Next
            Catch ex As Exception
                Log("Error in SearchObjects at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            'hs.WaitEvents()
        Loop

        SearchObjects = SList
        SList = Nothing

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchObjects for device - " & MyUPnPDeviceName & " returned " & RecordIndex.ToString & " records ", LogType.LOG_TYPE_INFO)

    End Function

    Private Function SearchInString(inString As String, inSearchString As String, SearchOperator As SearchOperatorTypes) As Boolean
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchInString called for device - " & MyUPnPDeviceName & " with string = " & inString & " and SearchString = " & inSearchString & " and Searchoperator = " & SearchOperator.ToString)
        SearchInString = False
        If Trim(inString) = "" Or Trim(inSearchString) = "" Then Exit Function
        Dim SearchStrings As String()
        If inSearchString <> "" Then
            SearchStrings = Split(inSearchString, ",")
            For Each sString In SearchStrings
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchInString called for device - " & MyUPnPDeviceName & " is comparing string = " & inString.ToUpper & " and SearchString = " & sString.ToUpper)
                Select Case SearchOperator
                    Case SearchOperatorTypes.soContains
                        If inString.ToUpper.IndexOf(Trim(sString).ToUpper) <> -1 Then
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchInString called for device - " & MyUPnPDeviceName & " is comparing string = " & inString.ToUpper & " and SearchString = " & sString.ToUpper & " and concluded TRUE")
                            SearchInString = True
                        End If
                    Case SearchOperatorTypes.soIsEqual
                        inString = Trim(inString)
                        If inString.ToUpper = Trim(sString).ToUpper Then
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchInString called for device - " & MyUPnPDeviceName & " is comparing string = " & inString.ToUpper & " and SearchString = " & sString.ToUpper & " and concluded TRUE")
                            SearchInString = True
                        End If
                End Select
            Next
        End If
    End Function

    Public Function SearchForContainers(ByVal ObjectID As String, SearchString As String, SearchOperator As SearchOperatorTypes, ByRef Records As DBRecord(), Optional ByVal DownToItems As Boolean = False) As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchForContainers called for device - " & MyUPnPDeviceName & " with ObjectID=" & ObjectID & " and SearchString = " & SearchString & " and Searchoperator = " & SearchOperator.ToString & " and DownToItems = " & DownToItems.ToString & " and depth counter = " & MySearchMaxDepthCounter.ToString, LogType.LOG_TYPE_INFO)
        SearchForContainers = ""
        MySearchMaxDepthCounter = MySearchMaxDepthCounter - 1
        If MySearchMaxDepthCounter <= 0 Then
            MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
            Exit Function
        End If
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchForContainers called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_INFO)
            MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
            Exit Function
        End If

        'Dim SList As DBRecord() = Nothing
        Dim LoopIndex As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0

        Dim InArg(5)
        Dim OutArg(3)
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 1                        ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SearchForContainers/Browse for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchForContainers found " & TotalMatches.ToString & " entries for device = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_INFO)
        If TotalMatches = 0 Then
            MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
            Exit Function
        End If

        StartIndex = 0
        Dim RecordIndex As Integer = 0
        Dim ClassFound As Boolean = False
        Do
            InArg(3) = StartIndex               ' Index         UI4
            InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in SearchForContainers for device = " & MyUPnPDeviceName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                SearchForContainers = "NOK"
                MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            If NumberReturned = 0 Then
                SearchForContainers = "OK"
                MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
                Exit Function
            End If


            Dim xmlData As XmlDocument = New XmlDocument
            Try
                xmlData.LoadXml(OutArg(0).ToString)
                'log( "XML=" & Value.ToString) ' used for testing
            Catch ex As Exception
                Log("Error in SearchForContainers at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                SearchForContainers = "NOK"
                MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
                Exit Function
            End Try
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchForContainers Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO)
                If PIDebuglevel > DebugLevel.dlEvents Then Log("SearchForContainers Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    Dim RefObjectID As String = ""
                    Try
                        RefObjectID = outerNode.Attributes("refID").Value
                    Catch ex As Exception
                        RefObjectID = ""
                    End Try
                    If RefObjectID <> "" Then
                        SearchForContainers = "OK"
                        ClassFound = True
                        MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
                        Exit Function
                    End If
                    Dim ClassInfo As String = ""
                    Try
                        ClassInfo = outerNode.Item("upnp:class").InnerText
                    Catch ex As Exception
                    End Try
                    Dim ClassInfos As String()
                    ClassInfos = Split(ClassInfo, ".")
                    Dim Container As Boolean = False
                    If UBound(ClassInfos, 1) > 0 Then
                        If ClassInfos(1).ToUpper = "CONTAINER" Then
                            Container = True
                        End If
                    End If
                    Dim ID As String = ""
                    Try
                        ID = outerNode.Attributes("id").Value
                    Catch ex As Exception
                    End Try
                    If Container Then
                        If SearchInString(ClassInfo, SearchString, SearchOperator) Then
                            ClassFound = True
                            Dim RecordTitle As String = ""
                            Try
                                RecordTitle = outerNode.Item("dc:title").InnerText
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchForContainers found a container with title: " & RecordTitle)
                            Catch ex As Exception
                            End Try
                            Dim Artist As String = ""
                            Try
                                Artist = outerNode.Item("upnp:artist").InnerText
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchForContainers found a container with title: " & RecordTitle)
                            Catch ex As Exception
                            End Try
                            Dim Album As String = ""
                            Try
                                Album = outerNode.Item("upnp:album").InnerText
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchForContainers found a container with title: " & RecordTitle)
                            Catch ex As Exception
                            End Try
                            Dim Genre As String = ""
                            Try
                                Genre = outerNode.Item("upnp:genre").InnerText
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchForContainers found a container with title: " & RecordTitle)
                            Catch ex As Exception
                            End Try
                            Dim ObjectClass As String = ""
                            Try
                                ObjectClass = outerNode.Item("upnp:class").InnerText
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "SearchForContainers found a container with class: " & ObjectClass)
                            Catch ex As Exception
                            End Try
                            Dim NewRecord = New DBRecord
                            NewRecord.Id = ID
                            NewRecord.ParentID = ""
                            NewRecord.Title = RecordTitle
                            NewRecord.AlbumName = Album
                            NewRecord.ArtistName = Artist
                            NewRecord.Genre = Genre
                            NewRecord.IconURL = ""
                            NewRecord.ClassType = ObjectClass
                            Try
                                NewRecord.ItemOrContainer = "CONTAINER"
                            Catch ex As Exception
                            End Try
                            If Records Is Nothing Then
                                ReDim Records(0)
                                Records(0) = NewRecord
                            Else
                                ReDim Preserve Records(RecordIndex)
                                Records(RecordIndex) = NewRecord
                            End If
                            If PIDebuglevel > DebugLevel.dlEvents Then Log("Record #" & RecordIndex.ToString & " with value " & NewRecord.Title, LogType.LOG_TYPE_INFO)
                            RecordIndex = RecordIndex + 1
                        End If
                        If DownToItems And Not ClassFound Then
                            SearchForContainers(ID, SearchString, SearchOperator, Records, DownToItems)
                        End If
                    ElseIf UCase(outerNode.Name) = "ITEM" Then
                        SearchForContainers = "OK"
                        ClassFound = True
                        MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1
                        Exit Function
                    Else
                        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SearchForContainers for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes found node = " & outerNode.Name.ToString, LogType.LOG_TYPE_INFO)
                    End If
                Next
            Catch ex As Exception
                Log("Error in SearchForContainers at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            If ClassFound Then Exit Do
            'hs.WaitEvents()
        Loop

        SearchForContainers = "OK"

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SearchForContainers for device - " & MyUPnPDeviceName & " returned " & RecordIndex.ToString & " records ", LogType.LOG_TYPE_INFO)
        MySearchMaxDepthCounter = MySearchMaxDepthCounter + 1

    End Function

    Public Function GetDescriptionFromObject(ByVal ObjectID As String) As String
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetDescriptionFromObject called for device - " & MyUPnPDeviceName & " with ObjectID=" & ObjectID, LogType.LOG_TYPE_INFO)
        GetDescriptionFromObject = ""
        If ContentDirectory Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetDescriptionFromObject called for device - " & MyUPnPDeviceName & " but no handle to ContentDirectory", LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        If ObjectID = "" Then Exit Function

        Dim ReturnString As String = ""
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0

        Dim InArg(5) As String
        Dim OutArg(3) As String
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseMetadata"         ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 1                         ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in GetDescriptionFromObject/Browse for device = " & MyUPnPDeviceName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)

        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetDescriptionFromObject found " & TotalMatches.ToString & " entries for device = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("GetDescriptionFromObject found " & TotalMatches.ToString & " entries for device = " & MyUPnPDeviceName & " and MetaData = " & OutArg(0).ToString, LogType.LOG_TYPE_INFO)
        If TotalMatches = 0 Then Exit Function
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(OutArg(0).ToString)
        Catch ex As Exception
            Log("Error in GetDescriptionFromObject for UPnPDevice = " & MyUPnPDeviceName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim ItemClassInfo As String = ""
        Dim ItemClassInfos As String() = Nothing
        Try
            ItemClassInfo = xmlData.GetElementsByTagName("upnp:class").Item(0).InnerText
            ItemClassInfos = Split(ItemClassInfo, ".")
            If UCase(ItemClassInfos(1)) <> "ITEM" Then
                ' this should not be
                Log("Error in GetDescriptionFromObject for device = " & MyUPnPDeviceName & ". The Object is a container, it is = " & ItemClassInfo.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in GetDescriptionFromObject for device = " & MyUPnPDeviceName & " while searching XML for upnp:class. XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
        End Try
        Dim ItemURL As String = ""
        ItemURL = FindRightSizePictureURL(OutArg(0).ToString, PictureSize.psTiny)
        If ItemURL <> "" Then
            If Not (ItemURL.ToLower().StartsWith("http://") Or ItemURL.ToLower().StartsWith("https://") Or ItemURL.ToLower().StartsWith("file:")) Then ItemURL = "http://" & MyIPAddress & ":" & MyIPPort & ItemURL
            ReturnString = "<td><img src=" & ItemURL & " height=""100"" width=""100""></td><td>"
        End If


        ' "dc:creator"

        ' "upnp:author"
        ' "dc:date"
        ' "dc:description"
        ' "upnp:rating"
        ' "upnp:actor"
        ' "upnp:albumArtist"
        ' "pv:lastPlayedTime"
        ' "pv:playcount"
        ' "upnp:lastPlaybackTime"
        ' "upnp:playbackCount"
        Dim Title As String = ""
        Try
            Title = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If Title <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Title = " & Title.ToString & "</p>"
        End If

        Dim artist As String = ""
        Try
            artist = xmlData.GetElementsByTagName("upnp:artist").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If artist <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Artist = " & artist.ToString & "</p>"
        End If

        Dim album As String = ""
        Try
            album = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If album <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Album = " & album.ToString & "</p>"
        End If

        Dim genre As String = ""
        Try
            genre = xmlData.GetElementsByTagName("upnp:genre").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If genre <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Genre = " & genre.ToString & "</p>"
        End If

        Dim description As String = ""
        Try
            description = xmlData.GetElementsByTagName("dc:description").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If description <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Description = " & description.ToString & "</p>"
        End If
        description = ""
        Try
            description = xmlData.GetElementsByTagName("upnp:longDescription").Item(0).InnerText() ' XBMC
        Catch ex As Exception
        End Try
        If description <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Description = " & description.ToString & "</p>"
        End If

        Dim programTitle As String = ""
        Try
            programTitle = xmlData.GetElementsByTagName("upnp:programTitle").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If programTitle <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Program Title = " & programTitle.ToString & "</p>"
        End If

        Dim episodeNumber As String = ""
        Try
            episodeNumber = xmlData.GetElementsByTagName("upnp:episodeNumber").Item(0).InnerText()
        Catch ex As Exception
        End Try
        If episodeNumber <> "" Then
            If ReturnString = "" Then ReturnString = ReturnString & "<td>"
            ReturnString = ReturnString & "<p>Episode Number = " & episodeNumber.ToString & "</p>"
        End If


        If ReturnString <> "" Then ReturnString = ReturnString & "</td>"
        GetDescriptionFromObject = ReturnString
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetDescriptionFromObject for device - " & MyUPnPDeviceName & " returned " & GetDescriptionFromObject.ToString, LogType.LOG_TYPE_INFO)
        xmlData = Nothing

    End Function

End Class

