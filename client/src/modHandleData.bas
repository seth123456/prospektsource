Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetAddress = FunAddr
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYMap) = GetAddress(AddressOf HandlePlayerXYMap)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SBank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SStartTutorial) = GetAddress(AddressOf HandleStartTutorial)
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    HandleDataSub(SEventData) = GetAddress(AddressOf Events_HandleEventData)
    HandleDataSub(SEventUpdate) = GetAddress(AddressOf Events_HandleEventUpdate)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(SEventOpen) = GetAddress(AddressOf HandleEventOpen)
    HandleDataSub(SCreateProjectile) = GetAddress(AddressOf HandleCreateProjectile)
    HandleDataSub(SEventGraphic) = GetAddress(AddressOf HandleEventGraphic)
    HandleDataSub(SClientTime) = GetAddress(AddressOf HandleClientTime)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SUpdateEffect) = GetAddress(AddressOf HandleUpdateEffect)
    HandleDataSub(SEffect) = GetAddress(AddressOf HandleEffect)
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    HandleDataSub(SThreshold) = GetAddress(AddressOf HandleThreshold)
    HandleDataSub(SSwearFilter) = GetAddress(AddressOf HandleSwearFilter)
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        DestroyGame
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), MyIndex, buffer.ReadBytes(buffer.Length), 0, 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString 'Parse(1)
    
    Set buffer = Nothing
    'DestroyGame
    MsgBox Msg, vbOKOnly, Options.Game_Name
    logoutGame
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' save options
    Options.Username = sUser

    If Options.savePass = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = sPass
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' player high index
    Player_HighIndex = buffer.ReadLong
    
    Set buffer = Nothing
    Call SetStatus("Receiving game data...")
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = buffer.ReadString
            .Vital(Vitals.HP) = buffer.ReadLong
            .Vital(Vitals.MP) = buffer.ReadLong
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next
            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    newCharSprite = 0
    newCharClass = 1
    
    curMenu = MENU_CLASS
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim z As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For I = 1 To Max_Classes

        With Class(I)
            .name = buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = buffer.ReadLong 'CLng(Parse(n + 2))
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .MaleSprite(X) = buffer.ReadLong
            Next
            
            ' get array size
            z = buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For X = 0 To z
                .FemaleSprite(X) = buffer.ReadLong
            Next
                            
            For X = 1 To Stats.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    faderAlpha = 0
    faderState = 5
    canFade = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = 1

    For I = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, I, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, I, buffer.ReadLong)
        PlayerInv(I).bound = buffer.ReadByte
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong) 'CLng(Parse(3)))
    PlayerInv(n).bound = buffer.ReadByte
    
    ' changes, clear drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    playerNum = buffer.ReadLong
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, buffer.ReadLong, Shield)
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, PlayerIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerIndex = buffer.ReadLong
    Player(PlayerIndex).MaxVital(Vitals.HP) = buffer.ReadLong
    Call SetPlayerVital(PlayerIndex, Vitals.HP, buffer.ReadLong)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, PlayerIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerIndex = buffer.ReadLong
    Player(PlayerIndex).MaxVital(Vitals.MP) = buffer.ReadLong
    Call SetPlayerVital(PlayerIndex, Vitals.MP, buffer.ReadLong)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, I, buffer.ReadLong
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, buffer.ReadLong
    TNL = buffer.ReadLong
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long, X As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    Call SetPlayerName(I, buffer.ReadString)
    Call SetPlayerLevel(I, buffer.ReadLong)
    Call SetPlayerPOINTS(I, buffer.ReadLong)
    Call SetPlayerSprite(I, buffer.ReadLong)
    Call SetPlayerMap(I, buffer.ReadLong)
    Call SetPlayerX(I, buffer.ReadLong)
    Call SetPlayerY(I, buffer.ReadLong)
    Call SetPlayerDir(I, buffer.ReadLong)
    Call SetPlayerAccess(I, buffer.ReadLong)
    Call SetPlayerPK(I, buffer.ReadLong)
    Call SetPlayerClass(I, buffer.ReadLong)
    
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat I, X, buffer.ReadLong
    Next

    ' Check if the player is the client player
    If I = MyIndex Then
        ' Reset directions
        wDown = False
        aDown = False
        sDown = False
        dDown = False
        upDown = False
        leftDown = False
        downDown = False
        rightDown = False
    End If

    ' Make sure they aren't walking
    Player(I).Moving = 0
    Player(I).xOffset = 0
    Player(I).yOffset = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    n = buffer.ReadLong
    Call SetPlayerX(I, X)
    Call SetPlayerY(I, Y)
    Call SetPlayerDir(I, Dir)
    Player(I).xOffset = 0
    Player(I).yOffset = 0
    Player(I).Moving = n

    Select Case GetPlayerDir(I)
        Case DIR_UP
            Player(I).yOffset = PIC_Y
        Case DIR_DOWN
            Player(I).yOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(I).xOffset = PIC_X
        Case DIR_RIGHT
            Player(I).xOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim MapNpcNum As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim Movement As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerDir(I, Dir)

    With Player(I)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim Dir As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    Dir = buffer.ReadLong

    With MapNpc(I)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(MyIndex, X)
    Call SetPlayerY(MyIndex, Y)
    Call SetPlayerDir(MyIndex, Dir)
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim buffer As clsBuffer
Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    thePlayer = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Call SetPlayerX(thePlayer, X)
    Call SetPlayerY(thePlayer, Y)
    Call SetPlayerDir(thePlayer, Dir)
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).xOffset = 0
    Player(thePlayer).yOffset = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    ' Set player to attacking
    Player(I).Attacking = 1
    Player(I).AttackTimer = timeGetTime
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    I = buffer.ReadLong
    ' Set player to attacking
    MapNpc(I).Attacking = 1
    MapNpc(I).AttackTimer = timeGetTime
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim Y As Long
Dim I As Long
Dim instM As Long
Dim NeedMap As Byte
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GettingMap = True
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Erase all players except self
    For I = 1 To MAX_PLAYERS
        If I <> MyIndex Then
            Call SetPlayerMap(I, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For I = 1 To MAX_BYTE
        Blood(I).X = 0
        Blood(I).Y = 0
        Blood(I).Sprite = 0
        Blood(I).timer = 0
    Next
    
    ' Real map num
    instM = buffer.ReadLong
    If instM > 0 Then
        X = instM
        'Get instanced map num
        instM = buffer.ReadLong
    Else
        ' Get map num
        X = buffer.ReadLong
    End If
    ' Get revision
    Y = buffer.ReadLong

    If FileExist(App.path & MAP_PATH & "map" & X & MAP_EXT) Then
        Call LoadMap(X)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = Y Then
            ' We do so we dont need the map
            NeedMap = 0
            FMOD.CacheNewMapSounds
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim I As Long
Dim buffer As clsBuffer
Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    Map.BossNpc = buffer.ReadLong
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(I).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(I).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(I).Tileset = buffer.ReadLong
                Map.Tile(X, Y).Autotile(I) = buffer.ReadByte
            Next
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadLong
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.Npc(X) = buffer.ReadLong
    Next
    Map.Fog = buffer.ReadByte
    Map.FogSpeed = buffer.ReadByte
    Map.FogOpacity = buffer.ReadByte
    
    Map.Red = buffer.ReadByte
    Map.Green = buffer.ReadByte
    Map.Blue = buffer.ReadByte
    Map.Alpha = buffer.ReadByte
    
    Map.Panorama = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    Map.BGS = buffer.ReadString
    
    Set buffer = Nothing
    
    initAutotiles
    FMOD.CacheNewMapSounds
    
    ' Save the map
    Call SaveMap(MapNum)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer, tmpLong As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For I = 1 To MAX_MAP_ITEMS
        With MapItem(I)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            tmpLong = buffer.ReadLong
            If tmpLong = 0 Then
                .bound = False
            Else
                .bound = True
            End If
        End With
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For I = 1 To MAX_MAP_NPCS
        With MapNpc(I)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
        End With
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
Dim I As Long
Dim MusicFile As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' clear the action msgs
    For I = 1 To MAX_BYTE
        ClearActionMsg (I)
    Next I
    Action_HighIndex = 1
    
    ' player music
    If InGame Then
        MusicFile = Trim$(Map.Music)
        If Not MusicFile = "None." Then
            FMOD.Music_Play MusicFile
        Else
            FMOD.Music_Stop
        End If
    End If
    
    ' get the npc high index
    For I = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(I).Num > 0 Then
            Npc_HighIndex = I + 1
            Exit For
        End If
    Next
    
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS
    
    ' now cache the positions
    initAutotiles
    
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.Alpha
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity

    GettingMap = False
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String
Dim color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    color = buffer.ReadLong
    Call AddText(Msg, color)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer, tmpLong As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        tmpLong = buffer.ReadLong
        If tmpLong = 0 Then
            .bound = False
        Else
            .bound = True
        End If
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' changes to inventory, need to clear any drop menu
    sDialogue = vbNullString
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapNpc(n)
        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(Npc(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim ResourceNum As Long
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim shopnum As Long
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopnum = buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim spellnum As Long
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellnum = buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(I) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if in map editor, we cache shit ourselves
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For I = 0 To Resource_Index
            MapResource(I).ResourceState = buffer.ReadByte
            MapResource(I).X = buffer.ReadLong
            MapResource(I).Y = buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingEnd = timeGetTime
    Ping = PingEnd - PingStart
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Message As String, color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    Message = buffer.ReadString
    color = buffer.ReadLong
    tmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    CreateActionMsg Message, color, tmpType, X, Y
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, Sprite As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For I = 1 To MAX_BYTE
        If Blood(I).X = X And Blood(I).Y = Y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .X = X
        .Y = Y
        .Sprite = Sprite
        .timer = timeGetTime
        .Alpha = 255
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .LockIndex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound AnimInstance(AnimationIndex).X, AnimInstance(AnimationIndex).Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNpcNum = buffer.ReadLong
    For I = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(I) = buffer.ReadLong
    Next
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadLong
    SpellCD(Slot) = timeGetTime
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Access As Long
Dim name As String
Dim Message As String
Dim Colour As Long
Dim Header As String
Dim PK As Long
Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = buffer.ReadString
    Header = buffer.ReadString
    saycolour = buffer.ReadLong
    
    ' Check access level
    If Access > 0 Then
        Colour = Yellow
    Else
        Colour = White
    End If
    
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    'frmMain.txtChat.SelColor = colour
    'frmMain.txtChat.SelText = vbNewLine & Header & Name & ": "
    'frmMain.txtChat.SelColor = saycolour
    'frmMain.txtChat.SelText = message
    'frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    'AddText vbNewLine & Header & Name & ": ", colour, True
    AddText Header & name & ": " & Message, Colour
        
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim shopnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopnum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop shopnum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    StunDuration = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To MAX_BANK
        Bank.Item(I).Num = buffer.ReadLong
        Bank.Item(I).Value = buffer.ReadLong
    Next
    
    InBank = True
    GUIWindow(GUI_BANK).visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InTrade = buffer.ReadLong
     
    GUIWindow(GUI_TRADE).visible = True
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InTrade = 0
    GUIWindow(GUI_TRADE).visible = False
    TradeStatus = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim dataType As Byte
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    dataType = buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For I = 1 To MAX_INV
            TradeYourOffer(I).Num = buffer.ReadLong
            TradeYourOffer(I).Value = buffer.ReadLong
        Next
        YourWorth = buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For I = 1 To MAX_INV
            TradeTheirOffer(I).Num = buffer.ReadLong
            TradeTheirOffer(I).Value = buffer.ReadLong
        Next
        TheirWorth = buffer.ReadLong & "g"
    End If
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Status As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Status = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case Status
        Case 0 ' clear
            TradeStatus = vbNullString
        Case 1 ' they've accepted
            TradeStatus = "Other player has accepted."
        Case 2 ' you've accepted
            TradeStatus = "Waiting for other player to accept."
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    myTarget = buffer.ReadLong
    myTargetType = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
        
    For I = 1 To MAX_HOTBAR
        Hotbar(I).Slot = buffer.ReadLong
        Hotbar(I).sType = buffer.ReadByte
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Player_HighIndex = buffer.ReadLong
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    entityType = buffer.ReadLong
    entityNum = buffer.ReadLong
    
    PlayMapSound X, Y, entityType, entityNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    theName = buffer.ReadString
    
    Dialogue "Trade Request", theName & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim theName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    theName = buffer.ReadString
    
    Dialogue "Party Invitation", theName & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, I As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    inParty = buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = buffer.ReadLong
    For I = 1 To MAX_PARTY_MEMBERS
        Party.Member(I) = buffer.ReadLong
    Next
    Party.MemberCount = buffer.ReadLong
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim playerNum As Long
Dim buffer As clsBuffer, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' which player?
    playerNum = buffer.ReadLong
    ' set vitals
    For I = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(I) = buffer.ReadLong
        Player(playerNum).Vital(I) = buffer.ReadLong
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleStartTutorial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    ' set the first message
    Dim FileName As String
    Dim ShowTutorial As Boolean
    
    FileName = App.path & "\data files\tutorial.ini"
    ShowTutorial = Val(GetVar(FileName, "INIT", "ShowTutorial"))
    
    If ShowTutorial = True Then
        GUIWindow(GUI_TUTORIAL).visible = True
        GUIWindow(GUI_CHAT).visible = False
        SetTutorialState 1
    Else
        GUIWindow(GUI_TUTORIAL).visible = False
        GUIWindow(GUI_CHAT).visible = True
        SendFinishTutorial
    End If
    
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, TargetType As Long, target As Long, Message As String, Colour As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    target = buffer.ReadLong
    TargetType = buffer.ReadLong
    Message = buffer.ReadString
    Colour = buffer.ReadLong
    
    AddChatBubble target, TargetType, Message, Colour
    Set buffer = Nothing
End Sub
Public Sub Events_HandleEventUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    CurrentEventIndex = buffer.ReadLong
    With CurrentEvent
        .Type = buffer.ReadLong
        GUIWindow(GUI_EVENTCHAT).visible = Not (.Type = Evt_Quit)
        GUIWindow(GUI_CHAT).visible = (.Type = Evt_Quit)
        'Textz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .Text(1 To DCount)
            ReDim chatOptState(1 To DCount)
            .HasText = True
            For d = 1 To DCount
                .Text(d) = buffer.ReadString
            Next d
        Else
            Erase .Text
            .HasText = False
            ReDim chatOptState(1 To 1)
        End If
        'Dataz
        DCount = buffer.ReadLong
        If DCount > 0 Then
            ReDim .Data(1 To DCount)
            .HasData = True
            For d = 1 To DCount
                .Data(d) = buffer.ReadLong
            Next d
            Else
            Erase .Data
            .HasData = False
        End If
    End With
    
    Set buffer = Nothing
End Sub

Public Sub Events_HandleEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EIndex As Long, S As Long, SCount As Long, d As Long, DCount As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    EIndex = buffer.ReadLong
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    
    Events(EIndex).name = buffer.ReadString
    Events(EIndex).chkSwitch = buffer.ReadByte
    Events(EIndex).chkVariable = buffer.ReadByte
    Events(EIndex).chkHasItem = buffer.ReadByte
    Events(EIndex).SwitchIndex = buffer.ReadLong
    Events(EIndex).SwitchCompare = buffer.ReadByte
    Events(EIndex).VariableIndex = buffer.ReadLong
    Events(EIndex).VariableCompare = buffer.ReadByte
    Events(EIndex).VariableCondition = buffer.ReadLong
    Events(EIndex).HasItemIndex = buffer.ReadLong
    SCount = buffer.ReadLong
    If SCount > 0 Then
        ReDim Events(EIndex).SubEvents(1 To SCount)
        Events(EIndex).HasSubEvents = True
        For S = 1 To SCount
            With Events(EIndex).SubEvents(S)
                .Type = buffer.ReadLong
                'Textz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Text(1 To DCount)
                    .HasText = True
                    For d = 1 To DCount
                        .Text(d) = buffer.ReadString
                    Next d
                Else
                    Erase .Text
                    .HasText = False
                End If
                'Dataz
                DCount = buffer.ReadLong
                If DCount > 0 Then
                    ReDim .Data(1 To DCount)
                    .HasData = True
                    For d = 1 To DCount
                        .Data(d) = buffer.ReadLong
                    Next d
                Else
                    Erase .Data
                    .HasData = False
                End If
            End With
        Next S
    Else
        Events(EIndex).HasSubEvents = False
        Erase Events(EIndex).SubEvents
    End If
    
    Events(EIndex).Trigger = buffer.ReadByte
    Events(EIndex).WalkThrought = buffer.ReadByte
    Events(EIndex).Animated = buffer.ReadByte
    For S = 0 To 2
        Events(EIndex).Graphic(S) = buffer.ReadLong
    Next
    Events(EIndex).Layer = buffer.ReadByte
    
    Set buffer = Nothing
End Sub
Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To MAX_SWITCHES
        Switches(I) = buffer.ReadString
    Next
    
    For I = 1 To MAX_VARIABLES
        Variables(I) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub HandleEventOpen(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadByte
    eventNum = buffer.ReadLong
    Player(MyIndex).EventOpen(eventNum) = n
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapKey", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCreateProjectile(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
    Dim AttackerIndex As Long
    Dim AttackerType As Long
    Dim TargetIndex As Long
    Dim TargetType As Long
    Dim GrhIndex As Long
    Dim Rotate As Long
    Dim RotateSpeed As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    Call buffer.WriteBytes(Data())

    AttackerIndex = buffer.ReadLong
    AttackerType = buffer.ReadLong
    TargetIndex = buffer.ReadLong
    TargetType = buffer.ReadLong
    GrhIndex = buffer.ReadLong
    Rotate = buffer.ReadLong
    RotateSpeed = buffer.ReadLong
    
    'Create the projectile
    Call CreateProjectile(AttackerIndex, AttackerType, TargetIndex, TargetType, GrhIndex, Rotate, RotateSpeed)
    
End Sub

Private Sub HandleEventGraphic(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadByte
    eventNum = buffer.ReadLong
    Player(MyIndex).EventGraphic(eventNum) = n
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventGraphic", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim temp As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    GameTime.Minute = buffer.ReadByte
    GameTime.Hour = buffer.ReadByte
    GameTime.Day = buffer.ReadByte
    GameTime.Month = buffer.ReadByte
    GameTime.Year = buffer.ReadLong
    
    Set buffer = Nothing
End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString
    
    FMOD.Music_Stop
    FMOD.Music_Play str
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString

    FMOD.Sound_Play str
    
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'Need to learn how to fadeout :P
    'do later... way later.. like, after release, maybe never
    FMOD.Music_Stop
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleUpdateEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim buffer As clsBuffer
Dim EffectSize As Long
Dim EffectData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Effect
    EffectSize = LenB(Effect(n))
    ReDim EffectData(EffectSize - 1)
    EffectData = buffer.ReadBytes(EffectSize)
    CopyMemory ByVal VarPtr(Effect(n)), ByVal VarPtr(EffectData(0)), EffectSize
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleUpdateEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub HandleEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, X As Long, Y As Long, EffectNum As Long, I As Long, LockType As Byte, LockIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    EffectNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    LockType = buffer.ReadByte
    LockIndex = buffer.ReadLong
    If Effect(EffectNum).isMulti = YES Then
        For I = 1 To MAX_MULTIPARTICLE
            If Effect(EffectNum).MultiParticle(I) > 0 Then
                CastEffect Effect(EffectNum).MultiParticle(I), X, Y, LockType, LockIndex
            End If
        Next
    Else
        CastEffect EffectNum, X, Y, LockType, LockIndex
    End If
    PlayMapSound X, Y, SoundEntity.seEffect, EffectNum
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, EffectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    EffectType = buffer.ReadLong
    
    Select Case EffectType
        Case SEFFECT_TYPE_FADEIN
            FadeType = 0
            FadeAmount = 255
        Case SEFFECT_TYPE_FADEOUT
            FadeType = 1
            FadeAmount = 0
        Case SEFFECT_TYPE_FLASH
            FlashTimer = timeGetTime + 150
        Case SEFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case SEFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case SEFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    Set buffer = Nothing
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleThreshold(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim eventNum As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadByte
    Set buffer = Nothing
    Player(MyIndex).Threshold = n
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleThreshold", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwearFilter(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MaxSwearWords = buffer.ReadLong
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords)
    For I = 1 To MaxSwearWords
        SwearFilter(I).BadWord = buffer.ReadString
        SwearFilter(I).NewWord = buffer.ReadString
    Next
        
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleEventGraphic", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
