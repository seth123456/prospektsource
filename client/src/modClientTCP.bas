Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.Port

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal I As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = timeGetTime
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    SetStatus "Connecting to server..."
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (timeGetTime <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SendData(ByRef Data() As Byte)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString name
    buffer.WriteString Password
    buffer.WriteLong App.Major
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CLogin
    buffer.WriteString name
    buffer.WriteString Password
    buffer.WriteLong App.Major
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAddChar
    buffer.WriteString name
    buffer.WriteLong Sex
    buffer.WriteLong ClassNum
    buffer.WriteLong Sprite
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseChar
    buffer.WriteLong CharSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString MsgTo
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong Player(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong MapNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub SendMapRespawn()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub SendRequestEditMap()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingStart = timeGetTime
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestPlayerData()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte statNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestLevelUp()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong bankslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseBank
    SendData buffer.ToArray()
    Set buffer = Nothing
    InBank = False
    GUIWindow(GUI_BANK).visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChangeBankSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim buffer As clsBuffer, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X) = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub PlayerTarget(ByVal target As Long, ByVal TargetType As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If myTargetType = TargetType And myTarget = target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PlayerTarget", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendFinishTutorial()
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CFinishTutorial
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendFinishTutorial", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendSaveEvent(ByVal EIndex As Long)
    Dim buffer As clsBuffer
    Dim I As Long, d As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    Set buffer = New clsBuffer
    
    buffer.WriteLong CSaveEventData
    buffer.WriteLong EIndex
    buffer.WriteString Events(EIndex).name
    buffer.WriteByte Events(EIndex).chkSwitch
    buffer.WriteByte Events(EIndex).chkVariable
    buffer.WriteByte Events(EIndex).chkHasItem
    buffer.WriteLong Events(EIndex).SwitchIndex
    buffer.WriteByte Events(EIndex).SwitchCompare
    buffer.WriteLong Events(EIndex).VariableIndex
    buffer.WriteByte Events(EIndex).VariableCompare
    buffer.WriteLong Events(EIndex).VariableCondition
    buffer.WriteLong Events(EIndex).HasItemIndex
    If Events(EIndex).HasSubEvents Then
        buffer.WriteLong UBound(Events(EIndex).SubEvents)
        For I = 1 To UBound(Events(EIndex).SubEvents)
            With Events(EIndex).SubEvents(I)
                buffer.WriteLong .Type
                If .HasText Then
                    buffer.WriteLong UBound(.Text)
                    For d = 1 To UBound(.Text)
                        buffer.WriteString .Text(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
                If .HasData Then
                    buffer.WriteLong UBound(.Data)
                    For d = 1 To UBound(.Data)
                        buffer.WriteLong .Data(d)
                    Next d
                Else
                    buffer.WriteLong 0
                End If
            End With
        Next I
    Else
        buffer.WriteLong 0
    End If
    
    buffer.WriteByte Events(EIndex).Trigger
    buffer.WriteByte Events(EIndex).WalkThrought
    buffer.WriteByte Events(EIndex).Animated
    For I = 0 To 2
        buffer.WriteLong Events(EIndex).Graphic(I)
    Next
    buffer.WriteByte Events(EIndex).Layer
    
    SendData buffer.ToArray

    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Events_SendSaveEvent", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Events_SendRequestEditEvents()
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditEvents
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Events_SendRequestEditEvents", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub Events_SendRequestEventsData()
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEventsData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Events_SendRequestEventsData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub Events_SendChooseEventOption(ByVal I As Long)
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChooseEventOption
    buffer.WriteLong I
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Events_SendChooseEventOption", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub RequestSwitchesAndVariables()
    Dim I As Long, buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSwitchesAndVariables
    SendData buffer.ToArray
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "RequestSwitchesAndVariables", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub SendSwitchesAndVariables()
    Dim I As Long, buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For I = 1 To MAX_SWITCHES
        buffer.WriteString Switches(I)
    Next
    For I = 1 To MAX_VARIABLES
        buffer.WriteString Variables(I)
    Next
    SendData buffer.ToArray
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSwitchesAndVariables", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
