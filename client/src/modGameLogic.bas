Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long, Tick As Long, TickFPS As Long, FPS As Long, I As Long, WalkTimer As Long, X As Long, Y As Long
Dim tmr25 As Long, tmr10000 As Long, mapTimer As Long, chatTmr As Long, targetTmr As Long, fogTmr As Long, barTmr As Long
Dim barDifference As Long, renderspeed As Long, tmr1000 As Long, fadeTmr As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' *** Start GameLoop ***
    Do While InGame
        Tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then
            ' check ping
            Call GetPing
            tmr10000 = Tick + 10000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If Count_Spellicon > 0 Then
                For I = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(I) > 0 Then
                        If SpellCD(I) > 0 Then
                            If SpellCD(I) + (Spell(PlayerSpells(I)).CDTime * 1000) < Tick Then
                                SpellCD(I) = 0
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If
            
            For I = 1 To MAX_BYTE
                CheckAnimInstance I
            Next
            
            tmr25 = Tick + 25
        End If
        
        If tmr1000 < Tick Then
            ' A second has passed, so process the time
            Call ProcessTime
            tmr1000 = Tick + 1000
        End If
        
        If chatTmr < Tick Then
            If ChatButtonUp Then
                ScrollChatBox 0
            End If
            If ChatButtonDown Then
                ScrollChatBox 1
            End If
            chatTmr = Tick + 50
        End If
        
        ' targetting
        If targetTmr < Tick Then
            If tabDown Then
                FindNearestTarget
            End If
            targetTmr = Tick + 50
        End If
        
        ' fog scrolling
        If fogTmr < Tick Then
            If CurrentFogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogTmr = Tick + 255 - CurrentFogSpeed
            End If
        End If
        
        ' ****** Parallax X ******
        If ParallaxX = -800 Then
            ParallaxX = 0
        Else
            ParallaxX = ParallaxX - 1
        End If
        
        ' ****** Parallax Y ******
        If ParallaxY = 0 Then
            ParallaxY = -600
        Else
            ParallaxY = ParallaxY + 1
        End If
        
        ' elastic bars
        If barTmr < Tick Then
            SetBarWidth BarWidth_GuiHP_Max, BarWidth_GuiHP
            SetBarWidth BarWidth_GuiSP_Max, BarWidth_GuiSP
            SetBarWidth BarWidth_GuiEXP_Max, BarWidth_GuiEXP
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    SetBarWidth BarWidth_NpcHP_Max(I), BarWidth_NpcHP(I)
                End If
            Next
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    SetBarWidth BarWidth_PlayerHP_Max(I), BarWidth_PlayerHP(I)
                    SetBarWidth BarWidth_PlayerMP_Max(I), BarWidth_PlayerMP(I)
                End If
            Next
            
            ' reset timer
            barTmr = Tick + 10
        End If
        
        ' Animations!
        If mapTimer < Tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            
            ' re-set timer
            mapTimer = Tick + 500
        End If
        
        ProcessWeather
        
        If fadeTmr < Tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                        
                    Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                    
                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            fadeTmr = Tick + 30
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    Call ProcessMovement(I)
                End If
            Next I

            ' Process npc movements (actually move them)
            For I = 1 To Npc_HighIndex
                If Map.Npc(I) > 0 Then
                    Call ProcessNpcMovement(I)
                End If
            Next I

            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If
        
        ' fader logic
        If canFade Then
            If faderAlpha <= 0 Then
                canFade = False
                faderAlpha = 0
            Else
                faderAlpha = faderAlpha - faderSpeed
            End If
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderspeed < Tick Then
            Call Render_Graphics
            renderspeed = timeGetTime + 15
        End If
        Call FMOD.UpdateSounds

        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < Tick + 20
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop

    frmMain.visible = False

    If isLogging Then
        isLogging = False
        MenuLoop
        GettingMap = True
        FMOD.Music_Stop
        FMOD.Music_Play Options.MenuMusic
    Else
        ' Shutdown the game
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MenuLoop()
Dim FrameTime As Long
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim faderTimer As Long
Dim tmr500 As Long, renderspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' *** Start GameLoop ***
    Do While inMenu
        Tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        ' fader logic
        ' 0, 1, 2, 3 = Fading in/out of intro
        ' 4 = fading in to main menu
        ' 5 = fading out of main menu
        ' 6 = fading in to game
        If canFade Then
            If faderTimer = 0 Then
                Select Case faderState
                    Case 0, 2, 4, 6 ' fading in
                        If faderAlpha <= 0 Then
                            faderTimer = Tick + 1000
                        Else
                            ' fade out a bit
                            faderAlpha = faderAlpha - faderSpeed
                        End If
                    Case 1, 3, 5 ' fading out
                        If faderAlpha >= 254 Then
                            If faderState < 5 Then
                                faderState = faderState + 1
                            ElseIf faderState = 5 Then
                                ' fading out to game - make game load during fade
                                faderAlpha = 254
                                HideMenu
                                ShowGame
                                Call GameInit
                                Call GameLoop
                                Exit Sub
                            End If
                        Else
                            ' fade in a bit
                            faderAlpha = faderAlpha + faderSpeed
                        End If
                End Select
            Else
                If faderTimer < Tick Then
                    ' change the speed
                    If faderState > 2 Then faderSpeed = 15
                    ' normal fades
                    If faderState < 4 Then
                        faderState = faderState + 1
                        faderTimer = 0
                    Else
                        faderTimer = 0
                    End If
                End If
            End If
        End If
        
        If tmr500 < Tick Then
            ' animate textbox
            If chatShowLine = "|" Then
                chatShowLine = vbNullString
            Else
                chatShowLine = "|"
            End If
            tmr500 = Tick + 500
        End If
        
        If Menu_Alert_Timer < Tick Then
            Menu_Alert_Message = vbNullString
            Menu_Alert_Colour = 0
            Menu_Alert_Timer = 0
            IsConnecting = False
        End If
        
        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderspeed < Tick Then
            Call Render_Menu
            renderspeed = timeGetTime + 15
        End If

        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < Tick + 20
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = WALK_SPEED '((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = RUN_SPEED ' ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        MovementSpeed = RUN_SPEED
    Else
        Exit Sub
    End If

    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed
            If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
            
        Case DIR_DOWN
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed
            If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
            
        Case DIR_LEFT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed
            If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
            
        Case DIR_RIGHT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed
            If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 1 Then
                    MapNpc(MapNpcNum).Step = 3
                Else
                    MapNpc(MapNpcNum).Step = 1
                End If
            End If
        Else
            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 1 Then
                    MapNpc(MapNpcNum).Step = 3
                Else
                    MapNpc(MapNpcNum).Step = 1
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim buffer As New clsBuffer, tmpIndex As Long, I As Long, X As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer

    If timeGetTime > Player(MyIndex).MapGetTimer + 250 Then
        ' find out if we want to pick it up
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).X = Player(MyIndex).X And MapItem(I).Y = Player(MyIndex).Y Then
                If MapItem(I).Num > 0 Then
                    If Item(MapItem(I).Num).BindType = 1 Then
                        ' make sure it's not a party drop
                        If Party.Leader > 0 Then
                            For X = 1 To MAX_PARTY_MEMBERS
                                tmpIndex = Party.Member(X)
                                If tmpIndex > 0 Then
                                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(I).playerName) Then
                                        If Item(MapItem(I).Num).ClassReq > 0 Then
                                            If Item(MapItem(I).Num).ClassReq <> Player(MyIndex).Class Then
                                                Dialogue "Loot Check", "This item is BoP and is not for your class. Are you sure you want to pick it up?", DIALOGUE_LOOT_ITEM, True
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Else
                        'not bound
                        Exit For
                    End If
                End If
            End If
        Next
        ' nevermind, pick it up
        Player(MyIndex).MapGetTimer = timeGetTime
        buffer.WriteLong CMapGetItem
        SendData buffer.ToArray()
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim buffer As clsBuffer
Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ControlDown Then
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < timeGetTime Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = timeGetTime
                End With

                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'If DirUp Or DirDown Or DirLeft Or DirRight Then
    If wDown Or sDown Or aDown Or dDown Or upDown Or leftDown Or downDown Or rightDown Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        GUIWindow(GUI_BANK).visible = False
    End If
    
    If GUIWindow(GUI_TUTORIAL).visible Then
        CanMove = False
        Exit Function
    End If

    d = GetPlayerDir(MyIndex)
    If wDown Or upDown Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If downDown Or sDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If aDown Or leftDown Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If dDown Or rightDown Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map.Tile(X, Y).Type = TILE_TYPE_EVENT Then
        If Map.Tile(X, Y).Data1 > 0 Then
            If Events(Map.Tile(X, Y).Data1).WalkThrought = NO Then
                If Player(MyIndex).EventOpen(Map.Tile(X, Y).Data1) = NO Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Check to see if a player is already on that tile
    If Map.Moral = 0 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next I
    End If

    ' Check to see if a npc is already on that tile
    For I = 1 To Npc_HighIndex
        If MapNpc(I).Num > 0 Then
            If MapNpc(I).X = X Then
                If MapNpc(I).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                GettingMap = True
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellSlot)).name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot) > 0 Then
        If timeGetTime > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong spellSlot
                SendData buffer.ToArray()
                Set buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = timeGetTime
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal Text As String, ByVal color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(Text, color)
        End If
    End If

    Debug.Print Text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .Type = MsgType
        .created = timeGetTime
        .Scroll = 1
        .X = X
        .Y = Y
        .Alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim LockIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = timeGetTime
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InShop = shopnum

    GUIWindow(GUI_SHOP).visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).Num
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(bankslot).Num = itemNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(bankslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    IsHotbarSlot = 0

    For I = 1 To MAX_HOTBAR
        Top = GUIWindow(GUI_HOTBAR).Y + HotbarTop
        Left = GUIWindow(GUI_HOTBAR).X + HotbarLeft + ((HotbarOffsetX + 32) * (((I - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = I
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)
        ' effects
        Case SoundEntity.seEffect
            If entityNum > MAX_EFFECTS Then Exit Sub
            soundName = Trim$(Effect(entityNum).sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    FMOD.Sound_Play soundName, X, Y
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    Dialogue_TitleCaption = diTitle
    Dialogue_TextCaption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        Dialogue_ButtonVisible(1) = False ' Yes button
        Dialogue_ButtonVisible(2) = True ' Okay button
        Dialogue_ButtonVisible(3) = False ' No button
    Else
        Dialogue_ButtonVisible(1) = True ' Yes button
        Dialogue_ButtonVisible(2) = False ' Okay button
        Dialogue_ButtonVisible(3) = True ' No button
    End If
    
    ' show the dialogue box
    GUIWindow(GUI_DIALOGUE).visible = True
    GUIWindow(GUI_CHAT).visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
Dim buffer As New clsBuffer
    Set buffer = New clsBuffer
    
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
            Case DIALOGUE_LOOT_ITEM
                ' send the packet
                Player(MyIndex).MapGetTimer = timeGetTime
                buffer.WriteLong CMapGetItem
                SendData buffer.ToArray()
            Case DIALOGUE_TYPE_QUIT
                dialogueIndex = 0
                logoutGame
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
        End Select
    End If
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top
End Function

Public Sub UpdateCamera()
Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1

    If EndX > Map.MaxX Then
        offsetX = 32

        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If

    If EndY > Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    CurX = TileView.Left + ((GlobalX + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((GlobalY + Camera.Top) \ PIC_Y)
    GlobalX_Map = GlobalX + (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
End Sub

Public Function IsBankItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsBankItem = 0
    
    For I = 1 To MAX_BANK
        If Not emptySlot Then
            If GetBankItemNum(I) <= 0 And GetBankItemNum(I) > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.bottom Then
                    
                    IsBankItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsBankItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim I As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
            Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))

            If X >= Left And X <= Left + 32 Then
                If Y >= Top And Y <= Top + 32 Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsShopItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsEqItem = 0

    For I = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, I) > 0 And GetPlayerEquipment(MyIndex, I) <= MAX_ITEMS Then

            With tempRec
                .Top = GUIWindow(GUI_CHARACTER).Y + EqTop
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_CHARACTER).X + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.bottom Then
                    IsEqItem = I
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsEqItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsInvItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsInvItem = 0

    For I = 1 To MAX_INV
        
        If Not emptySlot Then
            If GetPlayerInvItemNum(MyIndex, I) <= 0 Or GetPlayerInvItemNum(MyIndex, I) > MAX_ITEMS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_INVENTORY).Y + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.bottom Then
                    IsInvItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsInvItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsPlayerSpell = 0

    For I = 1 To MAX_PLAYER_SPELLS

        If Not emptySlot Then
            If PlayerSpells(I) <= 0 And PlayerSpells(I) > MAX_PLAYER_SPELLS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.bottom Then
                    IsPlayerSpell = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean, Optional ByVal emptySlot As Boolean = False) As Long
    Dim tempRec As RECT, skipThis As Boolean
    Dim I As Long
    Dim IsTradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsTradeItem = 0

    For I = 1 To MAX_INV
    
        If Yours Then
            IsTradeNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)
        Else
            IsTradeNum = TradeTheirOffer(I).Num
        End If
        
        If Not emptySlot Then
            If IsTradeNum <= 0 Or IsTradeNum > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
             With tempRec
                .Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.bottom Then
                    IsTradeItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsTradeItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String(Len(sString), "*")
End Function

Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                cacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub cacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > Count_Tileset Then
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Or Options.noAuto = 1 Then
            ' default to... default
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub calculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).X <> Map.Tile(X2, Y2).Layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).Y <> Map.Tile(X2, Y2).Layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub OpenNpcChat(ByVal npcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    ' set the shit
    chatNpc = npcNum
    chatText = mT
    tutOpt(1) = o1
    tutOpt(2) = o2
    tutOpt(3) = o3
    tutOpt(4) = o4
    ' we're in chat now boy
    GUIWindow(GUI_EVENTCHAT).visible = True
    GUIWindow(GUI_CHAT).visible = False
End Sub

Public Sub SetTutorialState(ByVal stateNum As Byte)
Dim FileName As String
Dim TutorialText(5) As String
Dim TutorialAnswer(5) As String
Dim TutorialIndex As Integer
Dim I As Long
    FileName = App.path & "\data files\tutorial.ini"

    For TutorialIndex = 1 To 5
        TutorialText(TutorialIndex) = GetVar(FileName, "TUTORIAL" & TutorialIndex, "Text")
        TutorialAnswer(TutorialIndex) = GetVar(FileName, "TUTORIAL" & TutorialIndex, "Answer")
    Next TutorialIndex


    Select Case stateNum
        Case 1 ' introduction
            chatText = TutorialText(1)
            tutOpt(1) = TutorialAnswer(1)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 2 ' next
            chatText = TutorialText(2)
            tutOpt(1) = TutorialAnswer(2)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 3 ' chatting
            chatText = TutorialText(3)
            tutOpt(1) = TutorialAnswer(3)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 4 ' combat
            chatText = TutorialText(4)
            tutOpt(1) = TutorialAnswer(4)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case 5 ' stats
            chatText = TutorialText(5)
            tutOpt(1) = TutorialAnswer(5)
            For I = 2 To 4
                tutOpt(I) = vbNullString
            Next
        Case Else ' goodbye
            chatText = vbNullString
            For I = 1 To 4
                tutOpt(I) = vbNullString
            Next
            SendFinishTutorial
            GUIWindow(GUI_TUTORIAL).visible = False
            GUIWindow(GUI_CHAT).visible = True
            AddText "Well done, you finished the tutorial.", BrightGreen
            Exit Sub
    End Select
    ' set the state
    tutorialState = stateNum
End Sub

Public Sub setOptionsState()
    ' music
    If Options.Music = 1 Then
        Buttons(Button_MusicOn).state = 2
        Buttons(Button_MusicOff).state = 0
    Else
        Buttons(Button_MusicOn).state = 0
        Buttons(Button_MusicOff).state = 2
    End If
    
    ' sound
    If Options.sound = 1 Then
        Buttons(Button_SoundOn).state = 2
        Buttons(Button_SoundOff).state = 0
    Else
        Buttons(Button_SoundOn).state = 0
        Buttons(Button_SoundOff).state = 2
    End If
    
    ' debug
    If Options.Debug = 1 Then
        Buttons(Button_DebugOn).state = 2
        Buttons(Button_DebugOff).state = 0
    Else
        Buttons(Button_DebugOn).state = 0
        Buttons(Button_DebugOff).state = 2
    End If
    
    ' autotile
    If Options.noAuto = 0 Then
        Buttons(Button_AutotileOn).state = 2
        Buttons(Button_AutotileOff).state = 0
    Else
        Buttons(Button_AutotileOn).state = 0
        Buttons(Button_AutotileOff).state = 2
    End If
    
    ' Fullscreen
    If Options.Fullscreen = 1 Then
        Buttons(Button_FullscreenOn).state = 2
        Buttons(Button_FullscreenOff).state = 0
    Else
        Buttons(Button_FullscreenOn).state = 0
        Buttons(Button_FullscreenOff).state = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If direction = 0 Then ' up
        ChatScroll = ChatScroll + 1
    Else ' down
        ChatScroll = ChatScroll - 1
    End If
    ' scrolling down
    If ChatScroll < 8 Then ChatScroll = 8
    ' scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    ' update the array
    UpdateChatArray
End Sub

Public Sub ClearMapCache()
Dim I As Long, FileName As String

    For I = 1 To MAX_MAPS
        FileName = App.path & "\data files\maps\map" & I & ".map"
        If FileExist(FileName) Then
            Kill FileName
        End If
    Next
    AddText "Map cache destroyed.", BrightGreen
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal Colour As Long)
Dim I As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For I = 1 To MAX_BYTE
        If chatBubble(I).TargetType = TargetType Then
            If chatBubble(I).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = I
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .TargetType = TargetType
        .Msg = SwearFilter_Replace(Msg)
        .Colour = Colour
        .timer = timeGetTime
        .active = True
    End With
End Sub

Public Sub FindNearestTarget()
Dim I As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, xDif As Long, yDif As Long
Dim bestX As Long, bestY As Long, bestIndex As Long

    X2 = GetPlayerX(MyIndex)
    Y2 = GetPlayerY(MyIndex)
    
    bestX = 255
    bestY = 255
    
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            X = MapNpc(I).X
            Y = MapNpc(I).Y
            ' find the difference - x
            If X < X2 Then
                xDif = X2 - X
            ElseIf X > X2 Then
                xDif = X - X2
            Else
                xDif = 0
            End If
            ' find the difference - y
            If Y < Y2 Then
                yDif = Y2 - Y
            ElseIf Y > Y2 Then
                yDif = Y - Y2
            Else
                yDif = 0
            End If
            ' best so far?
            If (xDif + yDif) < (bestX + bestY) Then
                bestX = xDif
                bestY = yDif
                bestIndex = I
            End If
        End If
    Next
    
    ' target the best
    If bestIndex > 0 And bestIndex <> myTarget Then PlayerTarget bestIndex, TARGET_TYPE_NPC
End Sub

Public Function FindTarget() As Boolean
Dim I As Long, X As Long, Y As Long

    ' check players
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            X = (GetPlayerX(I) * 32) + Player(I).xOffset + 32
            Y = (GetPlayerY(I) * 32) + Player(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget I, TARGET_TYPE_PLAYER
                    FindTarget = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' check npcs
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            X = (MapNpc(I).X * 32) + MapNpc(I).xOffset + 32
            Y = (MapNpc(I).Y * 32) + MapNpc(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    ' found our target!
                    PlayerTarget I, TARGET_TYPE_NPC
                    FindTarget = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub SetBarWidth(ByRef MaxWidth As Long, ByRef Width As Long)
Dim barDifference As Long
    If MaxWidth < Width Then
        ' find out the amount to increase per loop
        barDifference = ((Width - MaxWidth) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width - barDifference
    ElseIf MaxWidth > Width Then
        ' find out the amount to increase per loop
        barDifference = ((MaxWidth - Width) / 100) * 10
        ' if it's less than 1 then default to 1
        If barDifference < 1 Then barDifference = 1
        ' set the width
        Width = Width + barDifference
    End If
End Sub

' *****************
' ** Event Logic **
' *****************
Public Sub Events_SetSubEventType(ByVal EIndex As Long, ByVal SIndex As Long, ByVal EType As EventType)
    If EIndex <= 0 Or EIndex > MAX_EVENTS Then Exit Sub
    If SIndex < LBound(Events(EIndex).SubEvents) Or SIndex > UBound(Events(EIndex).SubEvents) Then Exit Sub
    
    'We are ok, allocate
    With Events(EIndex).SubEvents(SIndex)
        .Type = EType
        Select Case .Type
            Case Evt_Message
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Menu
                If Not .HasText Then ReDim .Text(1 To 2)
                If UBound(.Text) < 2 Then ReDim Preserve .Text(1 To 2)
                If Not .HasData Then ReDim .Data(1 To 1)
                .HasText = True
                .HasData = True
            Case Evt_OpenShop
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GOTO
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_GiveItem
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_PlayAnimation
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_Warp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_Switch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_Variable
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_AddText
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Chatbubble
                .HasText = True
                .HasData = True
                ReDim Preserve .Text(1 To 1)
                ReDim Preserve .Data(1 To 2)
            Case Evt_Branch
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 6)
            Case Evt_ChangeSkill
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeLevel
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_ChangeSprite
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangePK
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeClass
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeSex
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_ChangeExp
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 2)
            Case Evt_SetAccess
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_CustomScript
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 1)
            Case Evt_OpenEvent
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_ChangeGraphic
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 4)
            Case Evt_ChangeVitals
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 3)
            Case Evt_PlaySound
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_PlayBGM
                .HasText = True
                .HasData = False
                Erase .Data
                ReDim Preserve .Text(1 To 1)
            Case Evt_SpecialEffect
                .HasText = False
                .HasData = True
                Erase .Text
                ReDim Preserve .Data(1 To 5)
            Case Else
                .HasText = False
                .HasData = False
                Erase .Text
                Erase .Data
        End Select
    End With
End Sub


Public Function GetComparisonOperatorName(ByVal opr As ComparisonOperator) As String
    Select Case opr
        Case GEQUAL
            GetComparisonOperatorName = ">="
            Exit Function
        Case LEQUAL
            GetComparisonOperatorName = "<="
            Exit Function
        Case GREATER
            GetComparisonOperatorName = ">"
            Exit Function
        Case LESS
            GetComparisonOperatorName = "<"
            Exit Function
        Case EQUAL
            GetComparisonOperatorName = "="
            Exit Function
        Case NOTEQUAL
            GetComparisonOperatorName = "><"
            Exit Function
    End Select
    GetComparisonOperatorName = "Unknown"
End Function

Public Function GetEventTypeName(ByVal EventIndex As Long, SubIndex As Long) As String
Dim evtType As EventType
evtType = Events(EventIndex).SubEvents(SubIndex).Type
    Select Case evtType
        Case Evt_Message
            GetEventTypeName = "@Show Message: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_Menu
            GetEventTypeName = "@Show Choices"
            Exit Function
        Case Evt_Quit
            GetEventTypeName = "@Exit Event"
            Exit Function
        Case Evt_OpenShop
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "-" & Trim$(Shop(Events(EventIndex).SubEvents(SubIndex).Data(1)).name)
            Else
                GetEventTypeName = "@Open Shop: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "- None "
            End If
            Exit Function
        Case Evt_OpenBank
            GetEventTypeName = "@Open Bank"
            Exit Function
        Case Evt_GiveItem
            GetEventTypeName = "@Change Item"
            Exit Function
        Case Evt_ChangeLevel
            GetEventTypeName = "@Change Level"
            Exit Function
        Case Evt_PlayAnimation
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Play Animation: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & "." & Trim$(Animation(Events(EventIndex).SubEvents(SubIndex).Data(1)).name) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Else
                GetEventTypeName = "@Play Animation: None {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            End If
            Exit Function
        Case Evt_Warp
            GetEventTypeName = "@Warp to: " & Events(EventIndex).SubEvents(SubIndex).Data(1) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(2) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(3) & "}"
            Exit Function
        Case Evt_GOTO
            GetEventTypeName = "@GoTo: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_Switch
            If Events(EventIndex).SubEvents(SubIndex).Data(2) = 1 Then
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1) + 1) & " = True"
            Else
                GetEventTypeName = "@Change Switch: " & Events(EventIndex).SubEvents(SubIndex).Data(1) + 1 & "." & Switches(Events(EventIndex).SubEvents(SubIndex).Data(1) + 1) & " = False"
            End If
            Exit Function
        Case Evt_Variable
            GetEventTypeName = "@Change Variable: "
            Exit Function
        Case Evt_AddText
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(2)
                Case 0: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Player}"
                Case 1: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Map}"
                Case 2: GetEventTypeName = "@Add text: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "' {" & GetColorString(Events(EventIndex).SubEvents(SubIndex).Data(1)) & ", Global}"
            End Select
            Exit Function
        Case Evt_Chatbubble
            GetEventTypeName = "@Show chatbubble"
            Exit Function
        Case Evt_Branch
            GetEventTypeName = "@Conditional branch"
            Exit Function
        Case Evt_ChangeSkill
            GetEventTypeName = "@Change Spells"
            Exit Function
        Case Evt_ChangeSprite
            GetEventTypeName = "@Change Sprite: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_ChangePK
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Change PK: NO"
                Case 1: GetEventTypeName = "@Change PK: YES"
            End Select
            Exit Function
        Case Evt_ChangeClass
            If Events(EventIndex).SubEvents(SubIndex).Data(1) > 0 Then
                GetEventTypeName = "@Change Class: " & Trim$(Class(Events(EventIndex).SubEvents(SubIndex).Data(1)).name)
            Else
                GetEventTypeName = "@Change Class: None"
            End If
            Exit Function
        Case Evt_ChangeSex
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(1)
                Case 0: GetEventTypeName = "@Change Sex: MALE"
                Case 1: GetEventTypeName = "@Change Sex: FEMALE"
            End Select
            Exit Function
        Case Evt_ChangeExp
            GetEventTypeName = "@Change Exp"
            Exit Function
        Case Evt_SetAccess
            GetEventTypeName = "@Set Access: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_CustomScript
            GetEventTypeName = "@Custom Script: " & Events(EventIndex).SubEvents(SubIndex).Data(1)
            Exit Function
        Case Evt_OpenEvent
            Select Case Events(EventIndex).SubEvents(SubIndex).Data(3)
                Case 0: GetEventTypeName = "@Open Event: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
                Case 1: GetEventTypeName = "@Close Event: {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
            End Select
            Exit Function
        Case Evt_ChangeGraphic
            GetEventTypeName = "@Change graphic: " & Events(EventIndex).SubEvents(SubIndex).Data(3) & " {" & Events(EventIndex).SubEvents(SubIndex).Data(1) & ", " & Events(EventIndex).SubEvents(SubIndex).Data(2) & "}"
            Exit Function
        Case Evt_ChangeVitals
            GetEventTypeName = "@Change Vitals"
            Exit Function
        Case Evt_PlaySound
            GetEventTypeName = "@Play Sound: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_PlayBGM
            GetEventTypeName = "@Play Music: '" & Trim$(Events(EventIndex).SubEvents(SubIndex).Text(1)) & "'"
            Exit Function
        Case Evt_FadeoutBGM
            GetEventTypeName = "Stop Music"
            Exit Function
        Case Evt_SpecialEffect
            GetEventTypeName = "@Special Effect"
            Exit Function
    End Select
    GetEventTypeName = "Unknown"
End Function

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function
Public Sub ClearProjectile(ByVal ProjectileIndex As Integer)
 
    'Clear the selected index
    ProjectileList(ProjectileIndex).Graphic = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tx = 0
    ProjectileList(ProjectileIndex).ty = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
 
    'Update LastProjectile
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Graphic > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Note: Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

Exit Function

End Function

Public Sub CreateProjectile(ByVal AttackerIndex As Long, ByVal AttackerType As Long, ByVal TargetIndex As Long, ByVal TargetType As Long, ByVal Graphic As Long, ByVal Rotate As Long, ByVal RotateSpeed As Byte)
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Graphic > 0
    
    With ProjectileList(ProjectileIndex)
    
        ' ****** Initial Rotation Value ******
        .Rotate = Rotate
        
        ' ****** Set Values ******
        .Graphic = Graphic
        .RotateSpeed = RotateSpeed
    
        ' ****** Get Target Type ******
        Select Case AttackerType
            Case TARGET_TYPE_PLAYER
                .X = GetPlayerX(AttackerIndex) * PIC_X
                .Y = GetPlayerY(AttackerIndex) * PIC_Y
            Case TARGET_TYPE_NPC
                .X = MapNpc(AttackerIndex).X * PIC_X
                .Y = MapNpc(AttackerIndex).Y * PIC_Y
        End Select
        
        Select Case TargetType
            Case TARGET_TYPE_PLAYER
                .tx = Player(TargetIndex).X * PIC_X
                .ty = Player(TargetIndex).Y * PIC_Y
            Case TARGET_TYPE_NPC
                .tx = MapNpc(TargetIndex).X * PIC_X
                .ty = MapNpc(TargetIndex).Y * PIC_Y
        End Select
        
    End With
    
End Sub

Public Sub ProcessTime()
    With GameTime
        .Minute = .Minute + 1
        If .Minute >= 60 Then
            .Hour = .Hour + 1
            .Minute = 0
            
            If .Hour >= 24 Then
                .Day = .Day + 1
                .Hour = 0
                
                If .Day > GetMonthMax Then
                    .Month = .Month + 1
                    .Day = 1
                    
                    If .Month > 12 Then
                        .Year = .Year + 1
                        .Month = 1
                    End If
                End If
            End If
        End If
    End With
End Sub
Public Function GetMonthMax() As Byte
    Dim m As Byte
    m = GameTime.Month
    If m = 1 Or m = 3 Or m = 5 Or m = 7 Or m = 8 Or m = 10 Or m = 12 Then
        GetMonthMax = 31
    ElseIf m = 4 Or m = 6 Or m = 9 Or m = 11 Then
        GetMonthMax = 30
    ElseIf m = 2 Then
        GetMonthMax = 28
    End If
End Function

Public Sub CastEffect(ByVal EffectNum As Long, X As Long, Y As Long, LockType As Byte, LockIndex As Long)
    X = X * PIC_X + Half_PIC_X
    Y = Y * PIC_Y + Half_PIC_Y
    BeginEffect EffectNum, Effect(EffectNum).Type, ConvertMapX(X), ConvertMapY(Y), LockType, LockIndex
End Sub
Sub ClearEffect(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Effect(Index)), LenB(Effect(Index)))
    Effect(Index).name = vbNullString
    Effect(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearEffect", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearEffects()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_EFFECTS
        Call ClearEffect(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearEffects", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ProcessWeather()
Dim I As Long
    If CurrentWeather > 0 Then
        I = Rand(1, 101 - CurrentWeatherIntensity)
        If I = 1 Then
            'Add a new particle
            For I = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(I).InUse = False Then
                    If Rand(1, 2) = 1 Then
                        WeatherParticle(I).InUse = True
                        WeatherParticle(I).Type = CurrentWeather
                        WeatherParticle(I).Velocity = Rand(8, 14)
                        WeatherParticle(I).X = (TileView.Left * 32) - 32
                        WeatherParticle(I).Y = (TileView.Top * 32) + Rand(-32, frmMain.ScaleHeight)
                    Else
                        WeatherParticle(I).InUse = True
                        WeatherParticle(I).Type = CurrentWeather
                        WeatherParticle(I).Velocity = Rand(10, 15)
                        WeatherParticle(I).X = (TileView.Left * 32) + Rand(-32, frmMain.ScaleWidth)
                        WeatherParticle(I).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WEATHER_TYPE_STORM Then
        I = Rand(1, 400 - CurrentWeatherIntensity)
        If I = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            FMOD.Sound_Play Sound_Thunder
        End If
    End If
    
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).X > TileView.Right * 32 Or WeatherParticle(I).Y > TileView.bottom * 32 Then
                WeatherParticle(I).InUse = False
            Else
                WeatherParticle(I).X = WeatherParticle(I).X + WeatherParticle(I).Velocity
                WeatherParticle(I).Y = WeatherParticle(I).Y + WeatherParticle(I).Velocity
            End If
        End If
    Next
End Sub
