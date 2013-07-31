Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
' Actual input
Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetAsyncKeyState(VK_W) >= 0 Then wDown = False
    If GetAsyncKeyState(VK_S) >= 0 Then sDown = False
    If GetAsyncKeyState(VK_A) >= 0 Then aDown = False
    If GetAsyncKeyState(VK_D) >= 0 Then dDown = False
    
    If GetAsyncKeyState(VK_UP) >= 0 Then upDown = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then downDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then leftDown = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then rightDown = False
    
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    If GetAsyncKeyState(VK_TAB) >= 0 Then tabDown = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If

    'Move Up
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If
    
        ' move up
        If GetKeyState(vbKeyW) < 0 Then
            wDown = True
            sDown = False
            aDown = False
            dDown = False
            Exit Sub
        Else
            wDown = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            wDown = False
            sDown = False
            aDown = False
            dDown = True
            Exit Sub
        Else
            dDown = False
        End If
    
        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            wDown = False
            sDown = True
            aDown = False
            dDown = False
            Exit Sub
        Else
            sDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            wDown = False
            sDown = False
            aDown = True
            dDown = False
            Exit Sub
        Else
            aDown = False
        End If
        
        ' move up
        If GetKeyState(vbKeyUp) < 0 Then
            upDown = True
            leftDown = False
            downDown = False
            rightDown = False
            Exit Sub
        Else
            upDown = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            upDown = False
            leftDown = False
            downDown = False
            rightDown = True
            Exit Sub
        Else
            rightDown = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            upDown = False
            leftDown = False
            downDown = True
            rightDown = False
            Exit Sub
        Else
            downDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            upDown = False
            leftDown = True
            downDown = False
            rightDown = False
            Exit Sub
        Else
            leftDown = False
        End If
    Else
        wDown = False
        sDown = False
        aDown = False
        dDown = False
        upDown = False
        leftDown = False
        downDown = False
        rightDown = False
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyUp(ByVal keyCode As Long)
Dim I As Long

    If InGame Then
        ' admin pannel
        Select Case keyCode
            Case vbKeyEscape
                Dialogue "Quit Game", "Are you sure you want to quit game?", DIALOGUE_TYPE_QUIT, True
        End Select
        
        ' hotbar
        If Not chatOn Then
            For I = 1 To 9
                If keyCode = 48 + I Then
                    SendHotbarUse I
                End If
            Next
            If keyCode = 48 Then ' 0
                SendHotbarUse 10
            ElseIf keyCode = 189 Then ' -
                SendHotbarUse 11
            ElseIf keyCode = 187 Then ' =
                SendHotbarUse 12
            End If
        End If
    End If
    
    ' exit out of fade
    If inMenu Then
        If keyCode = vbKeyEscape Then
            If faderState < 4 Then
                faderState = 4
                faderAlpha = 0
            End If
        End If
    End If
End Sub

Public Sub HandleMenuKeyPresses(ByVal KeyAscii As Integer)
    If Not curMenu = MENU_LOGIN And Not curMenu = MENU_REGISTER And Not curMenu = MENU_NEWCHAR Then Exit Sub
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        Select Case curMenu
            Case MENU_LOGIN
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_LOGIN
                    End If
                End If
            Case MENU_REGISTER
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    curTextbox = 3
                ElseIf curTextbox = 3 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_NEWACCOUNT
                    End If
                End If
            Case MENU_NEWCHAR
                If KeyAscii = vbKeyReturn Then
                    MenuState MENU_STATE_ADDCHAR
                End If
        End Select
    End If
    
    Select Case curMenu
        Case MENU_LOGIN
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_REGISTER
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 3 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass2) > 0 Then sPass2 = Mid$(sPass2, 1, Len(sPass2) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass2 = sPass2 & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_NEWCHAR
            ' entering username
            If (KeyAscii = vbKeyBack) Then
                If LenB(sChar) > 0 Then sChar = Mid$(sChar, 1, Len(sChar) - 1)
            End If
        
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                sChar = sChar & ChrW$(KeyAscii)
            End If
    End Select
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim name As String
Dim I As Long
Dim n As Long
Dim Command() As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GUIWindow(GUI_CURRENCY).visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
    Else
        chatText = MyText
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        ' turn on/off the chat
        chatOn = Not chatOn

        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            Exit Sub
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(chatText)

                If Mid$(chatText, I, 1) <> Space(1) Then
                    name = name & Mid$(chatText, I, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, I, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - I > 0 Then
                MyText = Mid$(chatText, I + 1, Len(chatText) - I)
                ' Send the message to the player
                Call PlayerMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Global Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /who, /fps, /fpslock, /gui, /maps", HelpColor)
                Case "/maps"
                    ClearMapCache
                Case "/gui"
                    hideGUI = Not hideGUI
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' // Monitor Admin Commands //
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If
                    
                    GettingMap = True
                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If
                    
                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        GettingMap = True
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' // Creator Admin Commands //
                Case "/level"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    SendRequestLevelUp
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                Case "/spawnitem"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue
                    
                    If UBound(Command) < 2 Then
                        AddText "Usage: /spawnitem (item#) (value)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /spawnitem (item#) (value)", AlertColor
                        GoTo continue
                    End If
                    SendSpawnItem CLng(Command(1)), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    SendBanDestroy
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    
    If Not chatOn Then Exit Sub

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMouseMove(ByVal X As Long, ByVal Y As Long, ByVal Button As Long)
Dim I As Long

    ' Set the global cursor position
    GlobalX = (ScreenWidth / frmMain.ScaleWidth) * X
    GlobalY = (ScreenHeight / frmMain.ScaleHeight) * Y
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    
    ' GUI processing
    If Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (X >= GUIWindow(I).X And X <= GUIWindow(I).X + GUIWindow(I).Width) And (Y >= GUIWindow(I).Y And Y <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through them!
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((GlobalY + Camera.Top) \ PIC_Y)
End Sub

Public Sub HandleMouseDown(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_BARS, GUI_CHAT
                            ' nothing here so we can click through
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            Exit Sub
                        Case GUI_MENU
                            Menu_MouseDown Button
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_MAINMENU
                            MainMenu_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            Exit Sub
                        Case GUI_TUTORIAL
                            Tutorial_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If GUIWindow(GUI_CHAT).visible Then
            ChatScroll_MouseDown
        End If
    End If
    
    ' Handle events
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            'Call PlayerSearch(CurX, CurY)
            If FindTarget = False Then
                ' not in bank
                If InBank Then
                    'CanMove = False
                    'Exit Function
                    InBank = False
                    GUIWindow(GUI_BANK).visible = False
                End If
            End If
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim I As Long

    ' GUI processing
    If Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                        Case GUI_SPELLS
                            Spells_MouseUp
                        Case GUI_MENU
                            Menu_MouseUp
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_MAINMENU
                            MainMenu_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_TUTORIAL
                            Tutorial_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                        Case GUI_OPTIONS
                            Options_MouseUp
                        Case GUI_TRADE
                            Trade_MouseUp
                    End Select
                End If
            End If
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleDoubleClick()
Dim I As Long

    ' GUI processing
    If Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).X And GlobalX <= GUIWindow(I).X + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).Y And GlobalY <= GUIWindow(I).Y + GUIWindow(I).Height) Then
                If GUIWindow(I).visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                        Case GUI_SHOP
                            Shop_DoubleClick
                        Case GUI_BANK
                            Bank_DoubleClick
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).visible = Not GUIWindow(GUI_INVENTORY).visible
    Else
        GUIWindow(GUI_INVENTORY).visible = False
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).visible = Not GUIWindow(GUI_SPELLS).visible
    Else
        GUIWindow(GUI_SPELLS).visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).visible = Not GUIWindow(GUI_CHARACTER).visible
    Else
        GUIWindow(GUI_CHARACTER).visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_OPTIONS).visible = Not GUIWindow(GUI_OPTIONS).visible
    Else
        GUIWindow(GUI_OPTIONS).visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PARTY).visible = Not GUIWindow(GUI_PARTY).visible
    Else
        GUIWindow(GUI_PARTY).visible = False
    End If
End Sub

' Tutorial
Public Sub Tutorial_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    For I = 1 To 4
        If Len(Trim$(tutOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(tutOpt(I)) & "]")
            X = GUIWindow(GUI_TUTORIAL).X + 200 + (130 - (Width / 2))
            Y = GUIWindow(GUI_TUTORIAL).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                tutOptState(I) = 2 ' clicked
            End If
        End If
    Next
End Sub

Public Sub Tutorial_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    For I = 1 To 4
        If Len(Trim$(tutOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(tutOpt(I)) & "]")
            X = GUIWindow(GUI_TUTORIAL).X + 200 + (130 - (Width / 2))
            Y = GUIWindow(GUI_TUTORIAL).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' are we clicked?
                If tutOptState(I) = 2 Then
                    SetTutorialState tutorialState + 1
                    ' play sound
                    FMOD.Sound_Play Sound_ButtonClick
                End If
            End If
        End If
    Next
    
    For I = 1 To 4
        tutOptState(I) = 0 ' normal
    Next
End Sub

' Npc Chat
Public Sub Chat_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
    For I = 1 To UBound(CurrentEvent.Text) - 1
        If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
            X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
            Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                chatOptState(I) = 2 ' clicked
            End If
        End If
    Next
    Case Evt_Message
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    X = GUIWindow(GUI_EVENTCHAT).X + GUIWindow(GUI_EVENTCHAT).Width - Width - 10
    Y = GUIWindow(GUI_EVENTCHAT).Y + GUIWindow(GUI_EVENTCHAT).Height - 25
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End Select

End Sub
Public Sub Chat_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long

Select Case CurrentEvent.Type
    Case Evt_Menu
        For I = 1 To UBound(CurrentEvent.Text) - 1
            If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
                Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
                X = GUIWindow(GUI_EVENTCHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                Y = GUIWindow(GUI_EVENTCHAT).Y + 115 - ((I - 1) * 15)
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' are we clicked?
                    If chatOptState(I) = 2 Then
                        Events_SendChooseEventOption CurrentEvent.Data(I)
                        ' play sound
                        FMOD.Sound_Play Sound_ButtonClick
                    End If
                End If
            End If
        Next
        
        For I = 1 To UBound(CurrentEvent.Text) - 1
            chatOptState(I) = 0 ' normal
        Next
    Case Evt_Message
        Width = EngineGetTextWidth(Font_Default, "[Continue]")
        X = GUIWindow(GUI_EVENTCHAT).X + GUIWindow(GUI_EVENTCHAT).Width - Width - 10
        Y = GUIWindow(GUI_EVENTCHAT).Y + GUIWindow(GUI_EVENTCHAT).Height - 25
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' are we clicked?
            If chatContinueState = 2 Then
                Events_SendChooseEventOption CurrentEventIndex + 1
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        
        chatContinueState = 0
End Select
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = Button_ChatUp To Button_ChatDown
        X = GUIWindow(GUI_CHAT).X + Buttons(I).X
        Y = GUIWindow(GUI_CHAT).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            ' scroll the actual chat
            Select Case I
                Case Button_ChatUp ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case Button_ChatDown ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    I = Button_ShopExit
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case Button_ShopExit
                        ' exit
                        Set buffer = New clsBuffer
                        buffer.WriteLong CCloseShop
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        GUIWindow(GUI_SHOP).visible = False
                        InShop = 0
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    I = Button_ShopExit
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub

' Party
Public Sub Party_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = Button_PartyInvite To Button_PartyDisband
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case Button_PartyInvite ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Invalid invitation target.", BrightRed
                        End If
                    Case Button_PartyDisband ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "You are not in a party.", BrightRed
                        End If
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = Button_PartyInvite To Button_PartyDisband
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For I = Button_MusicOn To Button_FullscreenOff
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case Button_MusicOn ' music on
                        Options.Music = 1
                        FMOD.Music_Play Trim$(Map.Music)
                        SaveOptions
                        Buttons(Button_MusicOn).state = 2
                        Buttons(Button_MusicOff).state = 0
                    Case Button_MusicOff ' music off
                        FMOD.Music_Stop
                        Options.Music = 0
                        SaveOptions
                        Buttons(Button_MusicOn).state = 0
                        Buttons(Button_MusicOff).state = 2
                    Case Button_SoundOn ' sound on
                        Options.sound = 1
                        SaveOptions
                        Buttons(Button_SoundOn).state = 2
                        Buttons(Button_SoundOff).state = 0
                    Case Button_SoundOff ' sound off
                        FMOD.StopAllSounds
                        Options.sound = 0
                        SaveOptions
                        Buttons(Button_SoundOn).state = 0
                        Buttons(Button_SoundOff).state = 2
                    Case Button_DebugOn ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(Button_DebugOn).state = 2
                        Buttons(Button_DebugOff).state = 0
                    Case Button_DebugOff ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(Button_DebugOn).state = 0
                        Buttons(Button_DebugOff).state = 2
                    Case Button_AutotileOn ' Autotile on
                        Options.noAuto = 0
                        SaveOptions
                        Buttons(Button_AutotileOn).state = 2
                        Buttons(Button_AutotileOff).state = 0
                        ' cache render state
                        For X = 0 To Map.MaxX
                            For Y = 0 To Map.MaxY
                                For layerNum = 1 To MapLayer.Layer_Count - 1
                                    cacheRenderState X, Y, layerNum
                                Next
                            Next
                        Next
                    Case Button_AutotileOff ' Autotile off
                        Options.noAuto = 1
                        SaveOptions
                        Buttons(Button_AutotileOn).state = 0
                        Buttons(Button_AutotileOff).state = 2
                        ' cache render state
                        For X = 0 To Map.MaxX
                            For Y = 0 To Map.MaxY
                                For layerNum = 1 To MapLayer.Layer_Count - 1
                                    cacheRenderState X, Y, layerNum
                                Next
                            Next
                        Next
                    Case Button_FullscreenOn ' Fullscreen on
                        Options.Fullscreen = 1
                        SaveOptions
                        Buttons(Button_FullscreenOn).state = 2
                        Buttons(Button_FullscreenOff).state = 0
                    Case Button_FullscreenOff ' Fullscreen off
                        Options.Fullscreen = 0
                        SaveOptions
                        Buttons(Button_FullscreenOn).state = 0
                        Buttons(Button_FullscreenOff).state = 2
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = Button_MusicOn To Button_FullscreenOff
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = Button_Inventory To Button_Party
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case Button_Inventory
                        ' open window
                        OpenGuiWindow 1
                    Case Button_Spells
                        ' open window
                        OpenGuiWindow 2
                    Case Button_Character
                        ' open window
                        OpenGuiWindow 3
                    Case Button_Options
                        ' open window
                        OpenGuiWindow 4
                    Case Button_Trade
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendTradeRequest
                        Else
                            AddText "Invalid trade target.", BrightRed
                        End If
                    Case Button_Party
                        ' open window
                        OpenGuiWindow 6
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = Button_Inventory To Button_Party
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

' Main Menu
Public Sub MainMenu_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    If faderAlpha > 0 Then Exit Sub
    
    For I = Button_Login To Button_GenderRight
        X = GUIWindow(GUI_MAINMENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case Button_Login
                        ' login
                        DestroyTCP
                        curMenu = MENU_LOGIN
                        ' Load the username + pass
                        sUser = Trim$(Options.Username)
                        If Options.savePass = 1 Then
                            sPass = Trim$(Options.Password)
                        End If
                        curTextbox = 1
                    Case Button_Register
                        ' register
                        DestroyTCP
                        curMenu = MENU_REGISTER
                        ' clear the textbox
                        sUser = vbNullString
                        sPass = vbNullString
                        sPass2 = vbNullString
                        curTextbox = 1
                    Case Button_Credits
                        ' credits
                        DestroyTCP
                        curMenu = MENU_CREDITS
                    Case Button_Exit
                        ' exit
                        DestroyGame
                        Exit Sub
                    Case Button_LoginAccept
                        If curMenu = MENU_LOGIN Then
                            ' login accept
                            MenuState MENU_STATE_LOGIN
                        End If
                    Case Button_RegisterAccept
                        If curMenu = MENU_REGISTER Then
                            ' register accept
                            MenuState MENU_STATE_NEWACCOUNT
                        End If
                    Case Button_ClassAccept
                        If curMenu = MENU_CLASS Then
                            ' they've selected class - move on
                            sChar = vbNullString
                            curMenu = MENU_NEWCHAR
                        End If
                    Case Button_ClassNext
                        If curMenu = MENU_CLASS Then
                            ' next class
                            newCharClass = newCharClass + 1
                            If newCharClass > 3 Then
                                newCharClass = 1
                            End If
                        End If
                    Case Button_NewCharAccept
                        If curMenu = MENU_NEWCHAR Then
                            ' do eet
                            MenuState MENU_STATE_ADDCHAR
                        End If
                    Case Button_GenderLeft
                        If curMenu = MENU_NEWCHAR Then
                            ' do eet
                            ChangeGender
                        End If
                    Case Button_GenderRight
                        If curMenu = MENU_NEWCHAR Then
                            ' do eet
                            ChangeGender
                        End If
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next

    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub MainMenu_MouseDown(ByVal Button As Long)
Dim I As Long, X As Long, Y As Long

    If faderAlpha > 0 Then Exit Sub
    
    For I = Button_Login To Button_GenderRight
        X = GUIWindow(GUI_MAINMENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyText = "How many do you want to drop?"
                        tmpCurrencyItem = invNum
                        sDialogue = vbNullString
                        GUIWindow(GUI_CURRENCY).visible = True
                        GUIWindow(GUI_CHAT).visible = False
                        chatOn = True
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, Value As Long, multiplier As Double, I As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "How many do you want to deposit?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                GUIWindow(GUI_CHAT).visible = False
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).Num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)).Stackable = YES Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).Num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable = YES Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "How many do you want to trade?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).visible = True
                GUIWindow(GUI_CHAT).visible = False
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum) > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        ' drag it
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub

Public Sub Character_MouseDown()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = Button_AddStats1 To Button_AddStats5
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim I As Long, X As Long, Y As Long
    ' find out which button we're clicking
    For I = Button_AddStats1 To Button_AddStats5
        X = GUIWindow(GUI_CHARACTER).X + Buttons(I).X
        Y = GUIWindow(GUI_CHARACTER).Y + Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (I - (Button_AddStats1 - 1))
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    Next
End Sub

' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If Item(GetBankItemNum(bankNum)).Stackable = YES Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "How many do you want withdraw?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).visible = True
            GUIWindow(GUI_CHAT).visible = False
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim I As Long, X As Long, Y As Long

    ' find out which button we're clicking
    For I = Button_TradeAccept To Button_TradeDecline
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim I As Long, X As Long, Y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = Button_TradeAccept To Button_TradeDecline
        X = Buttons(I).X
        Y = Buttons(I).Y
        ' check if we're on the button
        If (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case Button_TradeAccept
                        AcceptTrade
                    Case Button_TradeDecline
                        DeclineTrade
                End Select
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Currency_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long, buffer As clsBuffer
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                End Select
            Else
                AddText "Please enter a valid amount.", BrightRed
            End If
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    End If
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    ' check if we're on the button
    If (GlobalX >= X And GlobalX <= X + Buttons(12).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            FMOD.Sound_Play Sound_ButtonClick
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).visible = False
    GUIWindow(GUI_CHAT).visible = True
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim I As Long, X As Long, Y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim I As Long, X As Long, Y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 105
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_CHAT).Y + 120
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                FMOD.Sound_Play Sound_ButtonClick
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).visible = False
    GUIWindow(GUI_CHAT).visible = True
    dialogueIndex = 0
End Sub
