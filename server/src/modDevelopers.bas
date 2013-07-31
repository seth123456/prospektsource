Attribute VB_Name = "modDevelopers"
' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Public Sub HandleDevLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            Password = MD5_string(Password)

            ' Check versions
            If Buffer.ReadLong <> App.Major Or Buffer.ReadLong <> App.Minor Or Buffer.ReadLong <> App.Revision Then
                Call AlertMsg(index, "Version outdated. Please run the auto-updater.")
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(index, Name)
            
            ' make sure they have enought access
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call AlertMsg(index, "Your access level is not high enough.")
                ClearPlayer index
                Exit Sub
            End If
            
            ' make sure they're not banned
            If isBanned_Account(index) Then
                Call AlertMsg(index, "Your account is banned from the game.")
                ClearPlayer index
                Exit Sub
            End If
            
            ' exit
            ClearBank index
            LoadBank index, Name
            
            Call SendMap(index, 1)
            'send Resource cache
            For i = 0 To ResourceCache(1).Resource_Count
                SendResourceCacheTo index, i
            Next
            Call SendClasses(index)
            Call SendItems(index)
            Call SendAnimations(index)
            Call SendNpcs(index)
            Call SendShops(index)
            Call SendSpells(index)
            Call SendResources(index)
            Call SendEffects(index)
            
            For i = 1 To MAX_EVENTS
                Call Events_SendEventData(index, i)
            Next
                
            ' send the login ok
            SendDevLoginOk index
            
            ' Show the player up on the socket status
            Call AddLog("Dev " & GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd("Dev " & GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

Sub HandleDevMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_CACHED_MAPS Then
        Exit Sub
    End If
    
    SendMap index, n
    'send Resource cache
    For i = 0 To ResourceCache(n).Resource_Count
        SendResourceCacheTo index, i
    Next
End Sub

Sub SendDevLoginOk(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SDevLoginOk
    Buffer.WriteLong index
    Buffer.WriteString Trim$(Player(index).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

