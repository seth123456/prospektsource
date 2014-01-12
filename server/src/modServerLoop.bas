Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long, LastUpdatePlayerTime As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = timeGetTime
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If timeGetTime > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If timeGetTime > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < timeGetTime Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                End If
            Next
            tmr25 = timeGetTime + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = timeGetTime + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            Call ProcessTime
            tmr1000 = timeGetTime + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = timeGetTime + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = timeGetTime + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = timeGetTime + 300000
        End If
        
        ' Checks to update player time every 5 minutes - Can be tweaked
        If Tick > LastUpdatePlayerTime Then
            SendClientTime
            SaveTime
            LastUpdatePlayerTime = timeGetTime + 300000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_CACHED_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next

            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NPCNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim targetX As Long, targetY As Long, target_verify As Boolean

    For MapNum = 1 To MAX_CACHED_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                If MapItem(MapNum, i).playerName <> vbNullString Then
                    ' make item public?
                    If Not MapItem(MapNum, i).Bound Then
                        If MapItem(MapNum, i).playerTimer < timeGetTime Then
                            ' make it public
                            MapItem(MapNum, i).playerName = vbNullString
                            MapItem(MapNum, i).playerTimer = 0
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                    ' despawn item?
                    If MapItem(MapNum, i).canDespawn Then
                        If MapItem(MapNum, i).despawnTimer < timeGetTime Then
                            ' despawn it
                            ClearMapItem i, MapNum
                            ' send updates to everyone
                            SendMapItemsToAll MapNum
                        End If
                    End If
                End If
            End If
        Next
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(i).Num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc MapNum, i, x
                    HandleHoT_Npc MapNum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < timeGetTime Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = timeGetTime
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = timeGetTime
            
            For x = 1 To MAX_MAP_NPCS
                NPCNum = MapNpc(MapNum).Npc(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = Npc(NPCNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(x).x - GetPlayerX(i)
                                        DistanceY = MapNpc(MapNum).Npc(x).y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                                                    Call SendChatBubble(MapNum, x, TARGET_TYPE_NPC, Trim$(Npc(NPCNum).AttackSay), DarkBrown)
                                                End If
                                                MapNpc(MapNum).Npc(x).targetType = TARGET_TYPE_PLAYER ' player
                                                MapNpc(MapNum).Npc(x).target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ' Check if target was found for NPC targetting
                            If MapNpc(MapNum).Npc(x).target = 0 Then
                                For i = 1 To MAX_MAP_NPCS
                                    ' exist?
                                    If MapNpc(MapNum).Npc(i).Num > 0 Then
                                        n = Npc(NPCNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(x).x - CLng(MapNpc(MapNum).Npc(i).x)
                                        DistanceY = MapNpc(MapNum).Npc(x).y - CLng(MapNpc(MapNum).Npc(i).y)
                                        
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
                                            
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(MapNpc(MapNum).Npc(x).Num).Moral > NPC_MORAL_NONE Then
                                                If Npc(MapNpc(MapNum).Npc(i).Num).Moral > NPC_MORAL_NONE Then
                                                    If Npc(MapNpc(MapNum).Npc(x).Num).Moral <> Npc(MapNpc(MapNum).Npc(i).Num).Moral Then
                                                        If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                                                            Call SendChatBubble(MapNum, x, TARGET_TYPE_NPC, Trim$(Npc(NPCNum).AttackSay), DarkBrown)
                                                        End If
                                                        MapNpc(MapNum).Npc(x).targetType = TARGET_TYPE_NPC
                                                        MapNpc(MapNum).Npc(x).target = i
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    If MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If timeGetTime > MapNpc(MapNum).Npc(x).StunTimer + (MapNpc(MapNum).Npc(x).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(x).StunDuration = 0
                            MapNpc(MapNum).Npc(x).StunTimer = 0
                        End If
                    Else
                        ' check if in conversation
                        If MapNpc(MapNum).Npc(x).inEventWith > 0 Then
                            ' check if we can stop having conversation
                            If Not TempPlayer(MapNpc(MapNum).Npc(x).inEventWith).inEventWith = x Then
                                MapNpc(MapNum).Npc(x).inEventWith = 0
                                MapNpc(MapNum).Npc(x).Dir = MapNpc(MapNum).Npc(x).e_lastDir
                                NpcDir MapNum, x, MapNpc(MapNum).Npc(x).Dir
                            End If
                        Else
                            target = MapNpc(MapNum).Npc(x).target
                            targetType = MapNpc(MapNum).Npc(x).targetType
        
                            ' Check to see if its time for the npc to walk
                            If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                                If targetType = 1 Then ' player
        
                                    ' Check to see if we are following a player or not
                                    If target > 0 Then
            
                                        ' Check if the player is even playing, if so follow'm
                                        If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                            DidWalk = False
                                            target_verify = True
                                            targetY = GetPlayerY(target)
                                            targetX = GetPlayerX(target)
                                        Else
                                            MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                            MapNpc(MapNum).Npc(x).target = 0
                                        End If
                                    End If
                                
                                ElseIf targetType = 2 Then 'npc
                                    
                                    If target > 0 Then
                                        
                                        If MapNpc(MapNum).Npc(target).Num > 0 Then
                                            DidWalk = False
                                            target_verify = True
                                            targetY = MapNpc(MapNum).Npc(target).y
                                            targetX = MapNpc(MapNum).Npc(target).x
                                        Else
                                            MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                            MapNpc(MapNum).Npc(x).target = 0
                                        End If
                                    End If
                                End If
                                
                                If target_verify Then
                                    
                                    ' if map has a working matrix, use a* pathfinding
                                    If mapMatrix(MapNum).created Then
                                        If GetPlayerX(target) <> MapNpc(MapNum).Npc(x).targetX Or GetPlayerY(target) <> MapNpc(MapNum).Npc(x).targetY Then
                                            ' player has moved, re-find the path
                                            MapNpc(MapNum).Npc(x).hasPath = APlus(MapNum, CLng(MapNpc(MapNum).Npc(x).x), CLng(MapNpc(MapNum).Npc(x).y), CLng(GetPlayerX(target)), CLng(GetPlayerY(target)), Void, MapNpc(MapNum).Npc(x).arPath)
                                            ' set the npc's cur path location
                                            If MapNpc(MapNum).Npc(x).hasPath Then MapNpc(MapNum).Npc(x).pathLoc = UBound(MapNpc(MapNum).Npc(x).arPath)
                                        End If
                                        ' if has path, follow it
                                        If MapNpc(MapNum).Npc(x).hasPath Then
                                                ' follow path
                                            NpcMoveAlongPath MapNum, x
                                            DidWalk = True
                                        End If
                                    End If
        
                                    ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(MapNum).Npc(x).x - 1 = targetX And MapNpc(MapNum).Npc(x).y = targetY Then
                                            If MapNpc(MapNum).Npc(x).Dir <> DIR_LEFT Then
                                                Call NpcDir(MapNum, x, DIR_LEFT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(MapNum).Npc(x).x + 1 = targetX And MapNpc(MapNum).Npc(x).y = targetY Then
                                            If MapNpc(MapNum).Npc(x).Dir <> DIR_RIGHT Then
                                                Call NpcDir(MapNum, x, DIR_RIGHT)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(MapNum).Npc(x).x = targetX And MapNpc(MapNum).Npc(x).y - 1 = targetY Then
                                            If MapNpc(MapNum).Npc(x).Dir <> DIR_UP Then
                                                Call NpcDir(MapNum, x, DIR_UP)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        If MapNpc(MapNum).Npc(x).x = targetX And MapNpc(MapNum).Npc(x).y + 1 = targetY Then
                                            If MapNpc(MapNum).Npc(x).Dir <> DIR_DOWN Then
                                                Call NpcDir(MapNum, x, DIR_DOWN)
                                            End If
        
                                            DidWalk = True
                                        End If
        
                                        ' We could not move so Target must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)
        
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
        
                                                If CanNpcMove(MapNum, x, i) Then
                                                    Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
        
                                Else
                                    i = Int(Rnd * 4)
        
                                    If i = 1 Then
                                        i = Int(Rnd * 4)
        
                                        If CanNpcMove(MapNum, x, i) Then
                                            Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    target = MapNpc(MapNum).Npc(x).target
                    targetType = MapNpc(MapNum).Npc(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If targetType = 1 Then ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                If Npc(MapNpc(MapNum).Npc(x).Num).Projectile > 0 Then
                                    TryNpcShootPlayer x, target
                                Else
                                    TryNpcAttackPlayer x, target
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(x).target = 0
                                MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                            End If
                        ElseIf targetType = TARGET_TYPE_NPC Then
                        ' Is the target NPC alive?
                            If MapNpc(MapNum).Npc(target).Num > 0 Then
                                If Npc(MapNpc(MapNum).Npc(x).Num).Projectile > 0 Then
                                    TryNpcShootNPC MapNum, x, target
                                Else
                                    TryNpcAttackNPC MapNum, x, target
                                End If
                            Else
                                ' npc is dead or non-existant, set target to 0
                                MapNpc(MapNum).Npc(x).target = 0
                                MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                            End If
                        End If
                    End If
                    
                    ' check for spells
                    If MapNpc(MapNum).Npc(x).spellBuffer.Spell = 0 Then
                        ' loop through and try and cast our spells
                        For i = 1 To MAX_NPC_SPELLS
                            If Npc(NPCNum).Spell(i) > 0 Then
                                NpcBufferSpell MapNum, x, i
                            End If
                        Next
                    Else
                        ' check the timer
                        If MapNpc(MapNum).Npc(x).spellBuffer.Timer + (Spell(Npc(NPCNum).Spell(MapNpc(MapNum).Npc(x).spellBuffer.Spell)).CastTime * 1000) < timeGetTime Then
                            ' cast the spell
                            NpcCastSpell MapNum, x, MapNpc(MapNum).Npc(x).spellBuffer.Spell, MapNpc(MapNum).Npc(x).spellBuffer.target, MapNpc(MapNum).Npc(x).spellBuffer.tType
                            ' clear the buffer
                            MapNpc(MapNum).Npc(x).spellBuffer.Spell = 0
                            MapNpc(MapNum).Npc(x).spellBuffer.target = 0
                            MapNpc(MapNum).Npc(x).spellBuffer.Timer = 0
                            MapNpc(MapNum).Npc(x).spellBuffer.tType = 0
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).Npc(x).stopRegen Then
                    If MapNpc(MapNum).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = MapNpc(MapNum).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(NPCNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(NPCNum, Vitals.HP) Then
                                MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(x).Num = 0 And Map(MapNum).Npc(x) > 0 Then
                    If TickCount > MapNpc(MapNum).Npc(x).SpawnWait + (Npc(Map(MapNum).Npc(x)).SpawnSecs * 1000) Then
                        ' if it's a boss chamber then don't let them respawn
                        If Map(MapNum).Moral = MAP_MORAL_BOSS Then
                            ' make sure the boss is alive
                            If Map(MapNum).BossNpc > 0 Then
                                If Map(MapNum).Npc(Map(MapNum).BossNpc) > 0 Then
                                    If x <> Map(MapNum).BossNpc Then
                                        If MapNpc(MapNum).Npc(Map(MapNum).BossNpc).Num > 0 Then
                                            Call SpawnNpc(x, MapNum)
                                        End If
                                    Else
                                        SpawnNpc x, MapNum
                                    End If
                                End If
                            End If
                        Else
                            Call SpawnNpc(x, MapNum)
                        End If
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If timeGetTime > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = timeGetTime
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

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
    Dim M As Byte
    M = GameTime.Month
    If M = 1 Or M = 3 Or M = 5 Or M = 7 Or M = 8 Or M = 10 Or M = 12 Then
        GetMonthMax = 31
    ElseIf M = 4 Or M = 6 Or M = 9 Or M = 11 Then
        GetMonthMax = 30
    ElseIf M = 2 Then
        GetMonthMax = 28
    End If
End Function

