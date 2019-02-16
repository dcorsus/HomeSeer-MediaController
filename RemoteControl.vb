Imports System.IO

Partial Public Class HSPI

    Private Sub TreatSetIOExRemoteControl(ButtonValue As Integer)
        If g_bDebug Then Log("TreatSetIOExRemoteControl called for UPnPDevice = " & MyUPnPDeviceName & " and buttonvalue = " & ButtonValue, LogType.LOG_TYPE_INFO)
        Dim RemoteType As String = GetStringIniFile(MyUDN, DeviceInfoIndex.diRemoteType.ToString, "")
        Select Case RemoteType.ToUpper
            Case "SAMSUNGIAPP", "SAMSUNGWEBSOCKET", "SAMSUNGWEBSOCKETPIN"
                TreatSetIOExSamsung(ButtonValue)
            Case "SONYIRRC"
                TreatSetIOExSony(ButtonValue)
            Case "DIAL"
                TreatSetIOExRoku(ButtonValue)
            Case "ONKYO"
                TreatSetIOExOnkyo(ButtonValue)
            Case "LG"
                TreatSetIOExLG(ButtonValue)
        End Select
    End Sub

    Private Sub CreateRemoteButtons(HSRef As Integer)

        If g_bDebug Then Log("CreateRemoteButtons called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)

        Dim objRemoteFile As String = gRemoteControlPath
        Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
        Try
            RemoteButtons = GetIniSection(MyUDN, objRemoteFile) '  As Dictionary(Of String, String)
            If RemoteButtons Is Nothing Then
                Log("Error in CreateRemoteButtons for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
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
                            RemoteButtonName = RemoteButtonInfos(0)
                            RemoteButtonValue = Val(RemoteButton.Key)
                            If Not ((RemoteButtonInfos(1).IndexOf("LGlaunch") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetinput") = 0) Or (RemoteButtonInfos(1).IndexOf("LGsetchannel") = 0)) Then
                                Dim Pair As VSPair
                                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
                                Pair.PairType = VSVGPairType.SingleValue
                                Pair.Value = RemoteButtonValue
                                Pair.Status = RemoteButtonName
                                Pair.Render = Enums.CAPIControlType.Button
                                Pair.Render_Location.Row = 1 + Val(RemoteButtonInfos(2))
                                Pair.Render_Location.Column = Val(RemoteButtonInfos(3))
                                If RemoteButtonInfos(1).IndexOf("/launch/") = 0 Then
                                    Dim Appid As String = RemoteButtonInfos(1)
                                    Appid = Appid.Remove(0, 8) ' remove the /launch"
                                    Appid = Trim(Appid)
                                    If File.Exists(CurrentAppPath & "/html" & URLArtWorkPath & "RokuChannelImage_" & Appid.ToString & ".png") Then ' dcor tralala tobe fixed needs to be remoted!
                                        Dim vg As New VGPair
                                        vg.PairType = VSVGPairType.SingleValue
                                        vg.Graphic = ImagesPath & "Artwork/RokuChannelImage_" & Appid.ToString & ".png"
                                        vg.Set_Value = RemoteButtonValue
                                        hs.DeviceVGP_AddPair(HSRef, vg)
                                        Pair.PairButtonImageType = Enums.CAPIControlButtonImage.Use_Custom                          ' added v35 on 10/2/2018
                                        Pair.PairButtonImage = ImagesPath & "Artwork/RokuChannelImage_" & Appid.ToString & ".png"   ' added v35 on 10/2/2018
                                    End If
                                End If
                                hs.DeviceVSP_AddPair(HSRef, Pair)
                            End If
                        End If
                    End If
                Next
                RemoteButtons = Nothing
            End If
        Catch ex As Exception
            Log("Error in CreateRemoteButtons parsing the .ini file for remote button info for device - " & MyUPnPDeviceName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub CreateButtonInfoInInifile()
        If g_bDebug Then Log("CreateButtonInfoInInifile called for device - " & MyUPnPDeviceName, LogType.LOG_TYPE_INFO)
        Dim objRemoteFile As String = gRemoteControlPath
        Dim RemoteButtons As New System.Collections.Generic.Dictionary(Of String, String)()
        Dim RowIndex As Integer = 1
        Dim ColumnIndex As Integer = 1
        Try
            RemoteButtons = GetIniSection(MyUDN & " - Default Codes", objRemoteFile) '  As Dictionary(Of String, String)
            If RemoteButtons Is Nothing Then
                Log("Error in CreateButtonInfoInInifile for device - " & MyUPnPDeviceName & ". No buttons are specified in the RemoteControl.ini file", LogType.LOG_TYPE_ERROR)
                Exit Try
            Else
                For Each RemoteButton In RemoteButtons
                    'Key= Name  and Value = KeyCode ; Index 
                    'Store as Key=Index, Value =  Active ; Given Name ; Key Code ; Row Index ; Column Index -- RowIndex and Column Index start with 1
                    If RemoteButton.Key <> "" Then
                        Dim RemoteButtonInfos As String() = Split(RemoteButton.Value, ":;:-:")
                        If UBound(RemoteButtonInfos, 1) > 0 Then
                            WriteStringIniFile(MyUDN, RemoteButtonInfos(1).ToString, RemoteButton.Key & ":;:-:" & RemoteButtonInfos(0).ToString & ":;:-:" & RowIndex.ToString & ":;:-:" & ColumnIndex.ToString, objRemoteFile)
                            ColumnIndex = ColumnIndex + 1
                            If ColumnIndex > 5 Then
                                ColumnIndex = 1
                                RowIndex = RowIndex + 1
                            End If
                        Else
                            Log("Error in CreateButtonInfoInInifile for device - " & MyUPnPDeviceName & ". The Value = " & RemoteButton.Value & " for Key = " & RemoteButton.Key.ToString & " is too short, elements are missing", LogType.LOG_TYPE_ERROR)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            log("Error in CreateButtonInfoInInifile for device - " & MyUPnPDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


End Class

