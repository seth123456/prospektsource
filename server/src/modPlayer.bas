Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendClientTimeTo(index)
    Call SendEffects(index)
    Call SendThreshold(index)
    Call SendSwearFilter(index)
    
    For i = 1 To MAX_EVENTS
        Call Events_SendEventData(index, i)
        Call SendEventOpen(index, Player(index).EventOpen(i), i)
        Call SendEventGraphic(index, Player(index).EventGraphic(i), i)
    Next
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    
    ' tell them to do the damn tutorial
    If Player(index).TutorialState = 0 Then SendStartTutorial index
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    Dim instanceMapID As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False
        ' Check if player was the only player on the map and stop npc processing if so
        If GetPlayerMap(index) > 0 Then
            instanceMapID = -1
            If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
                PlayersOnMap(GetPlayerMap(index)) = NO
                If IsInstancedMap(GetPlayerMap(index)) Then
                    instanceMapID = InstancedMaps(GetPlayerMap(index) - MAX_MAPS).OriginalMap
                    Call DestroyInstancedMap(GetPlayerMap(index) - MAX_MAPS)
                End If
            End If
            If instanceMapID > 0 Then
                SetPlayerMap index, instanceMapID
            End If
        End If
        ' Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = index Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' clear target
        For i = 1 To Player_HighIndex
            ' Prevent subscript out range
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                ' clear players target
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = index Then
                    TempPlayer(i).target = 0
                    TempPlayer(i).targetType = TARGET_TYPE_NONE
                    SendTarget i
                End If
            End If
        Next
        
        ' leave party.
        Party_PlayerLeave index

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_CACHED_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' clear target
    For i = 1 To Player_HighIndex
        ' Prevent subscript out range
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
            If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = index Then
                TempPlayer(i).target = 0
                TempPlayer(i).targetType = TARGET_TYPE_NONE
                SendTarget i
            End If
        End If
    Next
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If
    
    'Purge all NPCs targetting to this player
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(OldMap).Npc(i).targetType = TARGET_TYPE_PLAYER Then
            If MapNpc(OldMap).Npc(i).target = index Then
                MapNpc(OldMap).Npc(i).targetType = TARGET_TYPE_NONE
                MapNpc(OldMap).Npc(i).target = 0
            End If
        End If
    Next

    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        If IsInstancedMap(OldMap) Then
            Call DestroyInstancedMap(OldMap - MAX_MAPS)
        End If
        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    If IsInstancedMap(MapNum) Then
        Buffer.WriteLong InstancedMaps(MapNum - MAX_MAPS).OriginalMap
    Else
        Buffer.WriteLong 0
    End If
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, amount As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    MapNum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a event and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerY(index, GetPlayerY(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WalkThrought = YES Or (Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1) = YES) Then
                                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                                        SendPlayerMove index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerY(index, GetPlayerY(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WalkThrought = YES Or (Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1) = YES) Then
                                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                                        SendPlayerMove index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(index, GetPlayerX(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WalkThrought = YES Or (Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1) = YES) Then
                                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                                        SendPlayerMove index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(MapNum).MaxX Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_EVENT Then
                                Call SetPlayerX(index, GetPlayerX(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = YES
                            Else
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1 > 0 Then
                                    If Events(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WalkThrought = YES Or (Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1) = YES) Then
                                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                                        SendPlayerMove index, movement, sendToSelf
                                        Moved = YES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            If .Data4 = YES Then
                Call InstancedWarp(index, MapNum, x, y)
            Else
                Call PlayerWarp(index, MapNum, x, y)
            End If
            Moved = YES
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Moved = YES
        End If
        
        ' Event
        If .Type = TILE_TYPE_EVENT Then
            If .Data1 > 0 Then
                If Events(.Data1).Trigger = 0 Then InitEvent index, .Data1
            End If
            Moved = YES
        End If
        
        If .Type = TILE_TYPE_THRESHOLD Then
            If Player(index).Threshold = 1 Then
                Player(index).Threshold = 0
            Else
                Player(index).Threshold = 1
            End If
            ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
            SendThreshold index
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemnum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.shield

                    If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Stackable = YES Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Stackable = YES Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function HasItems(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Stackable = YES Then
                HasItems = GetPlayerInvItemValue(index, i)
            Else
                HasItems = HasItems + 1
            End If
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If
    
    If ItemVal = 0 Then ItemVal = 1

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Stackable = YES Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Player(index).Inv(i).Bound = 0
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    If Item(itemnum).Stackable = YES Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Player(index).Inv(invSlot).Bound = 0
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True, Optional ByVal forceBound As Boolean = False) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        ' force bound?
        If Not forceBound Then
            ' bind on pickup?
            If Item(itemnum).BindType = 1 Then ' bind on pickup
                Player(index).Inv(i).Bound = 1
                PlayerMsg index, "This item is now bound to your soul.", BrightRed
            Else
                Player(index).Inv(i).Bound = 0
            End If
        Else
            Player(index).Inv(i).Bound = 1
        End If
        ' send update
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Public Sub SetPlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
    If Item(itemID).Stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = itemID Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(index, i) = 0 Then
            Call SetPlayerInvItemNum(index, i, itemID)
            given = given + 1
            If Item(itemID).Stackable = YES Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(index, i)
        End If
    Next i
End Sub
Public Sub GivePlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
    Dim i As Long
    Dim given As Long
    If Item(itemID).Stackable = YES Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = itemID Then
                Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + itemCount)
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        If given >= itemCount Then Exit Sub
        If GetPlayerInvItemNum(index, i) = 0 Then
            Call SetPlayerInvItemNum(index, i, itemID)
            given = given + 1
            If Item(itemID).Stackable = YES Then
                Call SetPlayerInvItemValue(index, i, itemCount)
                given = itemCount
            End If
            Call SendInventoryUpdate(index, i)
        End If
    Next i
End Sub
Public Sub TakePlayerItems(ByVal index As Long, ByVal itemID As Long, ByVal itemCount As Long)
Dim i As Long
    If HasItems(index, itemID) >= itemCount Then
        If Item(itemID).Stackable = YES Then
            TakeInvItem index, itemID, itemCount
        Else
            For i = 1 To MAX_INV
                If HasItems(index, itemID) >= itemCount Then
                    If GetPlayerInvItemNum(index, i) = itemID Then
                        SetPlayerInvItemNum index, i, 0
                        SetPlayerInvItemValue index, i, 0
                        SendInventoryUpdate index, i
                    End If
                End If
            Next
        End If
    Else
        PlayerMsg index, "You need [" & itemCount & "] of [" & Trim$(Item(itemID).Name) & "]", AlertColor
    End If
End Sub

Function HasSpell(ByVal index As Long, ByVal spellnum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If Player(index).Spell(i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    MapNum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(index)) Then
                    If (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Stackable = YES Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                                Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
                            
                            ' is it bind on pickup?
                            Player(index).Inv(n).Bound = 0
                            If Item(GetPlayerInvItemNum(index, n)).BindType = 1 Or MapItem(MapNum, i).Bound Then
                                Player(index).Inv(n).Bound = 1
                                If Not Trim$(MapItem(MapNum, i).playerName) = Trim$(GetPlayerName(index)) Then
                                    PlayerMsg index, "This item is now bound to your soul.", BrightRed
                                End If
                            End If

                            ' Erase item from the map
                            ClearMapItem i, MapNum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim MapNum As Long, tmpIndex As Long, i As Long

    MapNum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).playerName = vbNullString Or MapItem(MapNum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    ' if in party show their party member's drops
    If TempPlayer(index).inParty > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            tmpIndex = Party(TempPlayer(index).inParty).Member(i)
            If tmpIndex > 0 Then
                If Trim$(GetPlayerName(tmpIndex)) = MapItem(MapNum, mapItemNum).playerName Then
                    If MapItem(MapNum, mapItemNum).Bound = 0 Then
                        CanPlayerPickupItem = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    ' exit out
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            ' make sure it's not bound
            If Item(GetPlayerInvItemNum(index, invNum)).BindType > 0 Then
                If Player(index).Inv(invNum).Bound = 1 Then
                    PlayerMsg index, "This item is soulbound and cannot be picked up by other players.", BrightRed
                End If
            End If
            
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
                If Player(index).Inv(invNum).Bound > 0 Then
                    MapItem(GetPlayerMap(index), i).Bound = True
                Else
                    MapItem(GetPlayerMap(index), i).Bound = False
                End If

                If Item(GetPlayerInvItemNum(index, invNum)).Stackable = YES Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                        Player(index).Inv(invNum).Bound = 0
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                    Player(index).Inv(invNum).Bound = 0
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn, MapItem(GetPlayerMap(index), i).Bound)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        For i = 1 To Vitals.Vital_Count - 1
            SendVital index, i
        Next
        SendEXP index
        SendPlayerData index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then
        Player(index).Level = MAX_LEVELS
        Exit Function
    End If
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = 100 + (((GetPlayerLevel(index) ^ 2) * 10) * 2)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).exp = exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
    x = Player(index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_CACHED_MAPS Then
        Player(index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).Inv(invSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub
Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
    Player(index).Spell(spellslot) = spellnum
End Sub
Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    ' Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                    If TempPlayer(i).target = index Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next

    ' Drop all worn items
    For i = 1 To MAX_INV
        PlayerMapDropItem index, i, GetPlayerInvItemValue(index, i)
    Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, Options.StartMap, Options.StartX, Options.StartY)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    'Purge all NPCs targetting to this player
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).Npc(i).targetType = TARGET_TYPE_PLAYER Then
            If MapNpc(GetPlayerMap(index)).Npc(i).target = index Then
                MapNpc(GetPlayerMap(index)).Npc(i).targetType = TARGET_TYPE_NONE
                MapNpc(GetPlayerMap(index)).Npc(i).target = 0
            End If
        End If
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    TempPlayer(index).InBank = False
    TempPlayer(index).InShop = 0
    If TempPlayer(index).InTrade > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Num = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
        Next
        
        TempPlayer(index).InTrade = 0
        TempPlayer(TempPlayer(index).InTrade).InTrade = 0
        
        SendCloseTrade index
        SendCloseTrade TempPlayer(index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim CanCut As Boolean
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1
        If Resource_index = 0 Then Exit Sub

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If
        Next
        
        ' Check if can be cut
        CanCut = False
        If Resource_num > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then
                    CanCut = True
                End If
            End If
            If Resource(Resource_index).ToolRequired = 0 Then CanCut = True
        End If
        If CanCut Then
            ' inv space?
            If Resource(Resource_index).ItemReward > 0 Then
                If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                    PlayerMsg index, "You have no inventory space.", BrightRed
                    Exit Sub
                End If
            End If

            ' check if already cut down
            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
                rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
                Damage = GetPlayerDamage(index)

                ' check if damage is more than health
                If Damage > 0 Then
                    SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                    SendEffect GetPlayerMap(index), Resource(Resource_index).Effect, rX, rY
                    SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                    SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                    ' cut it down!
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = timeGetTime
                        SendResourceCacheToMap GetPlayerMap(index), Resource_num
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        ' carry on
                        GiveInvItem index, Resource(Resource_index).ItemReward, 1
                    Else
                        ' just do the damage
                        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                    End If
                Else
                    ' too weak
                    SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                End If
            Else
                ' send message if it exists
                If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            End If
        Else
            PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    If BankSlot = 0 Then Exit Sub
    Bank(index).Item(BankSlot).Num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    If BankSlot = 0 Then Exit Sub
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(index, invSlot)).Stackable = YES Then
        If amount < 1 Then Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Stackable = YES Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerBankItemNum(index, BankSlot)).Stackable = YES Then
        If amount < 1 Then Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Stackable = YES Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, itemnum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ITEM_TYPE_ARMOR
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemnum, Armor
                
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemnum, 0

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If Item(itemnum).isTwoHanded > 0 Then
                    If GetPlayerEquipment(index, shield) > 0 Then
                        PlayerMsg index, "This is 2Handed weapon! Please unequip shield first.", BrightRed
                        Exit Sub
                    End If
                End If

                If GetPlayerEquipment(index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, Weapon)
                End If

                SetPlayerEquipment index, itemnum, Weapon
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemnum, 1
                
                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_HELMET
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemnum, Helmet
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SHIELD
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).isTwoHanded > 0 Then
                        PlayerMsg index, "You have 2Handed weapon equipped! Please unequip it first.", BrightRed
                        Exit Sub
                    End If
                End If

                If GetPlayerEquipment(index, shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, shield)
                End If

                SetPlayerEquipment index, itemnum, shield
                PlayerMsg index, "You equip " & CheckGrammar(Item(itemnum).Name), BrightGreen
                
                ' tell them if it's soulbound
                If Item(itemnum).BindType = 2 Then ' BoE
                    If Player(index).Inv(invNum).Bound = 0 Then
                        PlayerMsg index, "This item is now bound to your soul.", BrightRed
                    End If
                End If
                
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    If Item(tempItem).BindType > 0 Then
                        GiveInvItem index, tempItem, 0, , True ' give back the stored item
                        tempItem = 0
                    Else
                        GiveInvItem index, tempItem, 0
                        tempItem = 0
                    End If
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                
                Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 1)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                    
                        ' make sure they don't already know it
                        For i = 1 To MAX_PLAYER_SPELLS
                            If Player(index).Spell(i) > 0 Then
                                If Player(index).Spell(i) = n Then
                                    PlayerMsg index, "You already know this spell.", BrightRed
                                    Exit Sub
                                End If
                            End If
                        Next
                    
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq


                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Player(index).Spell(i) = n
                                    Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemnum, 0)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    SendPlayerSpells index
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_EVENT
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then
                    InitEvent index, n
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        End Select
    End If
End Sub

' *****************
' ** Event Logic **
' *****************
Private Function IsForwardingEvent(ByVal EType As EventType) As Boolean
    Select Case EType
        Case Evt_Menu, Evt_Message
            IsForwardingEvent = False
        Case Else
            IsForwardingEvent = True
    End Select
End Function

Public Sub InitEvent(ByVal index As Long, ByVal EventIndex As Long)
    If TempPlayer(index).CurrentEvent > 0 And TempPlayer(index).CurrentEvent <= MAX_EVENTS Then Exit Sub
    If Events(EventIndex).chkVariable > 0 Then
        If Not CheckComparisonOperator(Player(index).Variables(Events(EventIndex).VariableIndex), Events(EventIndex).VariableCondition, Events(EventIndex).VariableCompare) = True Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkSwitch > 0 Then
        If Not Player(index).Switches(Events(EventIndex).SwitchIndex) = Events(EventIndex).SwitchCompare Then
            Exit Sub
        End If
    End If
    
    If Events(EventIndex).chkHasItem > 0 Then
        If HasItem(index, Events(EventIndex).HasItemIndex) = 0 Then
            Exit Sub
        End If
    End If
    
    TempPlayer(index).CurrentEvent = EventIndex
    Call DoEventLogic(index, 1)
End Sub

Public Function CheckComparisonOperator(ByVal numOne As Long, ByVal numTwo As Long, ByVal opr As ComparisonOperator) As Boolean
    CheckComparisonOperator = False
    Select Case opr
        Case GEQUAL
            If numOne >= numTwo Then CheckComparisonOperator = True
        Case LEQUAL
            If numOne <= numTwo Then CheckComparisonOperator = True
        Case GREATER
            If numOne > numTwo Then CheckComparisonOperator = True
        Case LESS
            If numOne < numTwo Then CheckComparisonOperator = True
        Case EQUAL
            If numOne = numTwo Then CheckComparisonOperator = True
        Case NOTEQUAL
            If Not (numOne = numTwo) Then CheckComparisonOperator = True
    End Select
End Function

Public Sub DoEventLogic(ByVal index As Long, ByVal Opt As Long)
Dim x As Long, y As Long, i As Long, Buffer As clsBuffer
    
    If TempPlayer(index).CurrentEvent <= 0 Or TempPlayer(index).CurrentEvent > MAX_EVENTS Then GoTo EventQuit
    If Not (Events(TempPlayer(index).CurrentEvent).HasSubEvents) Then GoTo EventQuit
    If Opt <= 0 Or Opt > UBound(Events(TempPlayer(index).CurrentEvent).SubEvents) Then GoTo EventQuit
    
        With Events(TempPlayer(index).CurrentEvent).SubEvents(Opt)
            Select Case .Type
                Case Evt_Quit
                    GoTo EventQuit
                Case Evt_OpenShop
                    Call SendOpenShop(index, .Data(1))
                    TempPlayer(index).InShop = .Data(1)
                    GoTo EventQuit
                Case Evt_OpenBank
                    SendBank index
                    TempPlayer(index).InBank = True
                    GoTo EventQuit
                Case Evt_GiveItem
                    If .Data(1) > 0 And .Data(1) <= MAX_ITEMS Then
                        Select Case .Data(3)
                            Case 0: Call TakePlayerItems(index, .Data(1), .Data(2))
                            Case 1: Call SetPlayerItems(index, .Data(1), .Data(2))
                            Case 2: Call GivePlayerItems(index, .Data(1), .Data(2))
                        End Select
                    End If
                    SendInventory index
                Case Evt_ChangeLevel
                    Select Case .Data(2)
                        Case 0: Call SetPlayerLevel(index, .Data(1))
                        Case 1: Call SetPlayerLevel(index, GetPlayerLevel(index) + .Data(1))
                        Case 2: Call SetPlayerLevel(index, GetPlayerLevel(index) - .Data(1))
                    End Select
                    SendPlayerData index
                Case Evt_PlayAnimation
                    x = .Data(2)
                    y = .Data(3)
                    If x < 0 Then x = GetPlayerX(index)
                    If y < 0 Then y = GetPlayerY(index)
                    If x >= 0 And y >= 0 And x <= Map(GetPlayerMap(index)).MaxX And y <= Map(GetPlayerMap(index)).MaxY Then Call SendAnimation(GetPlayerMap(index), .Data(1), x, y)
                Case Evt_Warp
                    If .Data(4) = YES Then
                        Call InstancedWarp(index, .Data(1), .Data(2), .Data(3))
                    Else
                        Call PlayerWarp(index, .Data(1), .Data(2), .Data(3))
                    End If
                Case Evt_GOTO
                    Call DoEventLogic(index, .Data(1))
                    Exit Sub
                Case Evt_Switch
                    Player(index).Switches(.Data(1)) = .Data(2)
                Case Evt_Variable
                    Select Case .Data(2)
                        Case 0: Player(index).Variables(.Data(1)) = .Data(3)
                        Case 1: Player(index).Variables(.Data(1)) = Player(index).Variables(.Data(1)) + .Data(3)
                        Case 2: Player(index).Variables(.Data(1)) = Player(index).Variables(.Data(1)) - .Data(3)
                        Case 3: Player(index).Variables(.Data(1)) = Random(.Data(3), .Data(4))
                    End Select
                Case Evt_AddText
                    Select Case .Data(2)
                        Case 0: PlayerMsg index, Trim$(.Text(1)), .Data(1)
                        Case 1: MapMsg GetPlayerMap(index), Trim$(.Text(1)), .Data(1)
                        Case 2: GlobalMsg Trim$(.Text(1)), .Data(1)
                    End Select
                Case Evt_Chatbubble
                    Select Case .Data(1)
                        Case 0: SendChatBubble GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Trim$(.Text(1)), DarkBrown
                        Case 1: SendChatBubble GetPlayerMap(index), .Data(2), TARGET_TYPE_NPC, Trim$(.Text(1)), DarkBrown
                    End Select
                Case Evt_Branch
                    Select Case .Data(1)
                        Case 0
                            If CheckComparisonOperator(Player(index).Variables(.Data(6)), .Data(2), .Data(5)) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                        Case 1
                            If Player(index).Switches(.Data(5)) = .Data(2) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                        Case 2
                            If HasItems(index, .Data(2)) >= .Data(5) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                        Case 3
                            If GetPlayerClass(index) = .Data(2) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                        Case 4
                            If HasSpell(index, .Data(2)) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                        Case 5
                            If CheckComparisonOperator(GetPlayerLevel(index), .Data(2), .Data(5)) Then
                                Call DoEventLogic(index, .Data(3))
                                Exit Sub
                            Else
                                Call DoEventLogic(index, .Data(4))
                                Exit Sub
                            End If
                    End Select
                Case Evt_ChangeSkill
                    If .Data(2) = 0 Then
                        If FindOpenSpellSlot(index) > 0 Then
                            If HasSpell(index, .Data(1)) = False Then
                                SetPlayerSpell index, FindOpenSpellSlot(index), .Data(1)
                            End If
                        End If
                    Else
                        If HasSpell(index, .Data(1)) = True Then
                            For i = 1 To MAX_PLAYER_SPELLS
                                If Player(index).Spell(i) = .Data(1) Then
                                    SetPlayerSpell index, i, 0
                                End If
                            Next
                        End If
                    End If
                    SendPlayerSpells index
                Case Evt_ChangeSprite
                    SetPlayerSprite index, .Data(1)
                    SendPlayerData index
                Case Evt_ChangePK
                    SetPlayerPK index, .Data(1)
                    SendPlayerData index
                Case Evt_ChangeClass
                    SetPlayerClass index, .Data(1)
                    SendPlayerData index
                Case Evt_ChangeSex
                    Player(index).Sex = .Data(1)
                    SendPlayerData index
                Case Evt_ChangeExp
                    Select Case .Data(2)
                        Case 0: Call SetPlayerExp(index, .Data(1))
                        Case 1: Call SetPlayerExp(index, GetPlayerExp(index) + .Data(1))
                        Case 2: Call SetPlayerExp(index, GetPlayerExp(index) - .Data(1))
                    End Select
                    CheckPlayerLevelUp index
                    SendEXP index
                Case Evt_SetAccess
                    SetPlayerAccess index, .Data(1)
                    SendPlayerData index
                Case Evt_CustomScript
                    CustomScript index, .Data(1)
                Case Evt_OpenEvent
                    x = .Data(1)
                    y = .Data(2)
                    If .Data(3) = 0 Then
                        If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_EVENT And Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(x, y).Data1) = NO Then
                            Select Case .Data(4)
                                Case 0
                                    Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(x, y).Data1) = YES
                                    SendEventOpen index, YES, Map(GetPlayerMap(index)).Tile(x, y).Data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = YES
                                            SendEventOpen i, YES, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = YES
                                            SendEventOpen i, YES, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                            End Select
                        End If
                    Else
                        If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_EVENT And Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(x, y).Data1) = YES Then
                            Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(x, y).Data1) = NO
                            Select Case .Data(4)
                                Case 0
                                    Player(index).EventOpen(Map(GetPlayerMap(index)).Tile(x, y).Data1) = NO
                                    SendEventOpen index, NO, Map(GetPlayerMap(index)).Tile(x, y).Data1
                                Case 1
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = NO
                                            SendEventOpen i, NO, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                                Case 2
                                    For i = 1 To Player_HighIndex
                                        If IsPlaying(i) Then
                                            Player(i).EventOpen(Map(GetPlayerMap(i)).Tile(x, y).Data1) = NO
                                            SendEventOpen i, NO, Map(GetPlayerMap(i)).Tile(x, y).Data1
                                        End If
                                    Next i
                            End Select
                        End If
                    End If
                Case Evt_Changegraphic
                    x = .Data(1)
                    y = .Data(2)
                    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_EVENT Then
                        Select Case .Data(4)
                            Case 0
                                Player(index).EventGraphic(Map(GetPlayerMap(index)).Tile(x, y).Data1) = .Data(3)
                                SendEventGraphic index, .Data(3), Map(GetPlayerMap(index)).Tile(x, y).Data1
                            Case 1
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) Then
                                        Player(i).EventGraphic(Map(GetPlayerMap(i)).Tile(x, y).Data1) = .Data(3)
                                        SendEventGraphic i, .Data(3), Map(GetPlayerMap(i)).Tile(x, y).Data1
                                    End If
                                Next i
                            Case 2
                                For i = 1 To Player_HighIndex
                                    If IsPlaying(i) Then
                                        Player(i).EventGraphic(Map(GetPlayerMap(i)).Tile(x, y).Data1) = .Data(3)
                                        SendEventGraphic i, .Data(3), Map(GetPlayerMap(i)).Tile(x, y).Data1
                                    End If
                                Next i
                        End Select
                    End If
                Case Evt_ChangeVitals
                    Select Case .Data(3)
                        Case 0: Call SetPlayerVital(index, .Data(2), .Data(1))
                        Case 1: Call SetPlayerVital(index, .Data(2), GetPlayerVital(index, .Data(2)) + .Data(1))
                        Case 2: Call SetPlayerVital(index, .Data(2), GetPlayerVital(index, .Data(2)) - .Data(1))
                    End Select
                    SendVital index, .Data(2)
                Case Evt_PlaySound
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SPlaySound
                        Buffer.WriteString Trim$(.Text(1))
                        SendDataTo index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_PlayBGM
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SPlayBGM
                        Buffer.WriteString Trim$(.Text(1))
                        SendDataTo index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_FadeoutBGM
                    Set Buffer = New clsBuffer
                        Buffer.WriteLong SFadeoutBGM
                        SendDataTo index, Buffer.ToArray
                    Set Buffer = Nothing
                Case Evt_SpecialEffect
                    Select Case .Data(1)
                        Case 0: SendSpecialEffect index, SEFFECT_TYPE_FADEOUT
                        Case 1: SendSpecialEffect index, SEFFECT_TYPE_FADEIN
                        Case 2: SendSpecialEffect index, SEFFECT_TYPE_FLASH
                        Case 3: SendSpecialEffect index, SEFFECT_TYPE_FOG, .Data(2), .Data(3), .Data(4)
                        Case 4: SendSpecialEffect index, SEFFECT_TYPE_WEATHER, .Data(2), .Data(3)
                        Case 5: SendSpecialEffect index, SEFFECT_TYPE_TINT, .Data(2), .Data(3), .Data(4), .Data(5)
                    End Select
            End Select
        End With
    
    'Make sure this is last
    If IsForwardingEvent(Events(TempPlayer(index).CurrentEvent).SubEvents(Opt).Type) Then
        Call DoEventLogic(index, Opt + 1)
    Else
        Call Events_SendEventUpdate(index, TempPlayer(index).CurrentEvent, Opt)
    End If
    
Exit Sub
EventQuit:
    TempPlayer(index).CurrentEvent = -1
    TempPlayer(index).inEventWith = 0
    Events_SendEventQuit index
    Exit Sub
End Sub

Sub CheckEvent(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Event_num As Long
    Dim Event_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_EVENT Then
        Event_index = Map(GetPlayerMap(index)).Tile(x, y).Data1
    End If
    
    If Event_index > 0 Then
        If Events(Event_index).Trigger > 0 Then
            InitEvent index, Event_index
        End If
    End If
End Sub

