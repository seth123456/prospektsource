Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
        Case MP
            GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = Item(weaponNum).Data2 + (((Item(weaponNum).Data2 / 100) * 5) * GetPlayerStat(index, Strength))
    Else
        GetPlayerDamage = 1 + (((0.01) * 5) * GetPlayerStat(index, Strength))
    End If

End Function

Function GetPlayerDefence(ByVal index As Long) As Long
Dim Defence As Long, i As Long, itemnum As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' base defence
    For i = 1 To Equipment.Equipment_Count - 1
        If i <> Equipment.Weapon Then
            itemnum = GetPlayerEquipment(index, i)
            If itemnum > 0 Then
                If Item(itemnum).Data2 > 0 Then
                    Defence = Defence + Item(itemnum).Data2
                End If
            End If
        End If
    Next
    
    ' divide by 3
    Defence = Defence / 3
    
    ' floor it at 1
    If Defence < 1 Then Defence = 1
    
    ' add in a player's agility
    GetPlayerDefence = Defence + (((Defence / 100) * 2.5) * (GetPlayerStat(index, Agility) / 2))
End Function

Function GetPlayerSpellDamage(ByVal index As Long, ByVal spellnum As Long, ByVal Vital As Vitals) As Long
Dim Damage As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' return damage
    Damage = Spell(spellnum).Vital(Vital)
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetPlayerSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcSpellDamage(ByVal NPCNum As Long, ByVal spellnum As Long, ByVal Vital As Vitals) As Long
Dim Damage As Long

    ' Check for subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then Exit Function
    
    ' return damage
    Damage = Spell(spellnum).Vital(Vital)
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetNpcSpellDamage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(NPCNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(NPCNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(NPCNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Npc(NPCNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    ' return the calculation
    GetNpcDamage = Npc(NPCNum).Damage + (((Npc(NPCNum).Damage / 100) * 5) * Npc(NPCNum).Stat(Stats.Strength))
End Function

Function GetNpcDefence(ByVal NPCNum As Long) As Long
Dim Defence As Long
    
    ' base defence
    Defence = 2
    
    ' add in a player's agility
    GetNpcDefence = Defence + (((Defence / 100) * 2.5) * (Npc(NPCNum).Stat(Stats.Agility) / 2))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(NPCNum).Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNpcNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNpcNum) Then
    
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        'damage = damage - RAND(1, (Npc(NpcNum).Stat(Stats.Agility) * 2))
        Damage = Damage - RAND((GetNpcDefence(NPCNum) / 100) * 10, (GetNpcDefence(NPCNum) / 100) * 10)
        ' randomise from 1 to max hit
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, MapNpcNum, Damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' exit out early
        If isSpell Then
             If NPCNum > 0 Then
                If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                    TempPlayer(Attacker).target = MapNpcNum
                    SendTarget Attacker
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NPCNum > 0 And timeGetTime > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                        TempPlayer(Attacker).target = MapNpcNum
                        SendTarget Attacker
                        CanPlayerAttackNpc = True
                    Else
                        ' init conversation if it's friendly
                        If Npc(NPCNum).Event > 0 Then
                            With MapNpc(MapNum).Npc(MapNpcNum)
                                .inEventWith = Attacker
                                .e_lastDir = .Dir
                                If GetPlayerY(Attacker) = .y - 1 Then
                                    .Dir = DIR_UP
                                ElseIf GetPlayerY(Attacker) = .y + 1 Then
                                    .Dir = DIR_DOWN
                                ElseIf GetPlayerX(Attacker) = .x - 1 Then
                                    .Dir = DIR_LEFT
                                ElseIf GetPlayerX(Attacker) = .x + 1 Then
                                    .Dir = DIR_RIGHT
                                End If
                                ' Set chat value to Npc
                                TempPlayer(Attacker).inEventWith = MapNpcNum
                                TempPlayer(Attacker).inEventMap = MapNum
                                ' send NPC's dir to the map
                                NpcDir MapNum, MapNpcNum, .Dir
                            End With
                            InitEvent Attacker, Npc(NPCNum).Event
                            Exit Function
                        End If
                        If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(MapNum, MapNpcNum, TARGET_TYPE_NPC, Trim$(Npc(NPCNum).AttackSay), DarkBrown)
                        End If
                        ' Reset attack timer
                        TempPlayer(Attacker).AttackTimer = timeGetTime
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    Name = Trim$(Npc(NPCNum).Name)
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    If spellnum > 0 Then
        Call SendAnimation(MapNum, Spell(spellnum).SpellAnim, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, TARGET_TYPE_NPC, MapNpcNum)
        Call SendEffect(MapNum, Spell(spellnum).Effect, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, TARGET_TYPE_NPC, MapNpcNum)
        SendMapSound Attacker, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, spellnum
    Else
        If n > 0 Then
            Call SendAnimation(MapNum, Item(n).Animation, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, TARGET_TYPE_NPC, MapNpcNum)
            Call SendEffect(MapNum, Item(n).Effect, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, TARGET_TYPE_NPC, MapNpcNum)
        End If
    End If
    
    SendActionMsg GetPlayerMap(Attacker), "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
    SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y

    If Damage >= MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) Then
        ' Calculate exp to give attacker
        exp = Npc(NPCNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, Npc(NPCNum).Level
        Else
            ' no party - keep exp for self
            GivePlayerEXP Attacker, exp, Npc(NPCNum).Level
        End If
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(NPCNum).DropItem(n) = 0 Then Exit For
            If Rnd <= Npc(NPCNum).DropChance(n) Then
                Call SpawnItem(Npc(NPCNum).DropItem(n), Npc(NPCNum).DropItemValue(n), MapNum, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, GetPlayerName(Attacker))
            End If
        Next
        
        If Npc(NPCNum).Event > 0 Then InitEvent Attacker, Npc(NPCNum).Event
        
        ' destroy map npcs
        If Map(MapNum).Moral = MAP_MORAL_BOSS Then
            If MapNpcNum = Map(MapNum).BossNpc Then
                ' kill all the other npcs
                For i = 1 To MAX_MAP_NPCS
                    If Map(MapNum).Npc(i) > 0 Then
                        ' only kill dangerous npcs
                        If Npc(Map(MapNum).Npc(i)).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(Map(MapNum).Npc(i)).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' kill!
                            MapNpc(MapNum).Npc(i).Num = 0
                            MapNpc(MapNum).Npc(i).SpawnWait = timeGetTime
                            MapNpc(MapNum).Npc(i).Vital(Vitals.HP) = 0
                            ' send kill command
                            SendNpcDeath MapNum, i
                        End If
                    End If
                Next
            End If
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNpcNum).Num = 0
        MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = timeGetTime
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        SendNpcDeath MapNum, MapNpcNum
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) - Damage
        
        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 1 ' player
        MapNpc(MapNum).Npc(MapNpcNum).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNpcNum).Num Then
                    MapNpc(MapNum).Npc(i).target = Attacker
                    MapNpc(MapNum).Npc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC MapNpcNum, MapNum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNpcNum, spellnum, Attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
        
        ' set the player's target if they don't have one
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
            TempPlayer(Attacker).target = MapNpcNum
            SendTarget Attacker
        End If
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
Dim MapNum As Long, NPCNum As Long, blockAmount As Long, Damage As Long, Defence As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, index) Then
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerDefence(index)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if crit hit
        If CanNpcCrit(NPCNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long, Optional ByVal isSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(index)
    NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If
    
    ' exit out early if it's a spell
    If isSpell Then
        If IsPlaying(index) Then
            If NPCNum > 0 Then
                CanNpcAttackPlayer = True
                Exit Function
            End If
        End If
    End If
    
    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim MapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
    MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = timeGetTime
    
    ' send the sound
    If spellnum > 0 Then
        Call SendAnimation(MapNum, Spell(spellnum).SpellAnim, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        Call SendEffect(MapNum, Spell(spellnum).Effect, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        SendMapSound Victim, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, spellnum
    Else
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).Animation, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        Call SendEffect(MapNum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).Effect, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
    End If
        
    ' if stunning spell, stun the npc
    If spellnum > 0 Then
        If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
        ' DoT
        If Spell(spellnum).Duration > 0 Then
            ' TODO: Add Npc vs Player DOTs
        End If
    End If
    
    ' Say damage
    SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
    SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If spellnum > 0 Then
            SendMapSound Victim, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, spellnum
        Else
            SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        End If
        
        ' send animation
        If Not overTime Then
            If spellnum = 0 Then Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Animation, GetPlayerX(Victim), GetPlayerY(Victim))
        End If
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = timeGetTime
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long, NPCNum As Long, MapNum As Long, Damage As Long, Defence As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Defence = GetPlayerDefence(Victim)
        If Defence > 0 Then
            Damage = Damage - RAND(Defence - ((Defence / 100) * 10), Defence + ((Defence / 100) * 10))
        End If
        
        ' randomise for up to 10% lower than max hit
        If Damage <= 0 Then Damage = 1
        Damage = RAND(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal isSpell As Boolean = False) As Boolean
Dim partynum As Long, i As Long

    If Not isSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function
    
    ' make sure it's not you
    If Victim = Attacker Then
        PlayerMsg Attacker, "Cannot attack yourself.", BrightRed
        Exit Function
    End If
    
    ' check co-ordinates if not spell
    If Not isSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 5 Then
        Call PlayerMsg(Attacker, "You are below level 5, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 5 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 5, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' make sure not in your party
    partynum = TempPlayer(Attacker).inParty
    If partynum > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partynum).Member(i) > 0 Then
                If Victim = Party(partynum).Member(i) Then
                    PlayerMsg Attacker, "Cannot attack party members.", BrightRed
                    Exit Function
                End If
            End If
        Next
    End If
    
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = Victim
    SendTarget Attacker
    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = timeGetTime
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If

    If spellnum > 0 Then
        Call SendAnimation(GetPlayerMap(Victim), Spell(spellnum).SpellAnim, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        Call SendEffect(GetPlayerMap(Victim), Spell(spellnum).Effect, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, spellnum
    Else
        If n > 0 Then
            Call SendAnimation(GetPlayerMap(Victim), Item(n).Animation, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
            Call SendEffect(GetPlayerMap(Victim), Item(n).Effect, GetPlayerX(Victim), GetPlayerY(Victim), TARGET_TYPE_PLAYER, Victim)
        End If
    End If
        
    SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
    SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - exp)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(Attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).inParty, exp, Attacker, GetPlayerLevel(Victim)
            Else
                ' not in party, get exp for self
                GivePlayerEXP Attacker, exp, GetPlayerLevel(Victim)
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = Victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = timeGetTime
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player Victim, spellnum, Attacker
            End If
        End If
        
        ' change target if need be
        If TempPlayer(Attacker).target = 0 Then
            TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
            TempPlayer(Attacker).target = Victim
            SendTarget Attacker
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
Dim spellnum As Long, mpCost As Long, LevelReq As Long, MapNum As Long, spellCastType As Long, ClassReq As Long
Dim AccessReq As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = Player(index).Spell(spellslot)
    MapNum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' make sure we're not buffering already
    If TempPlayer(index).spellBuffer.Spell = spellslot Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    mpCost = Spell(spellnum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < mpCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            spellCastType = 2 ' targetted
        Else
            spellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            spellCastType = 0 ' self-cast
        Else
            spellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    Range = Spell(spellnum).Range
    HasBuffered = False
    
    Select Case spellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    If Spell(spellnum).VitalType(Vitals.HP) = 0 Or Spell(spellnum).VitalType(Vitals.MP) = 0 Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    Else
                        HasBuffered = True
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if beneficial magic then self-cast it instead
                If Spell(spellnum).VitalType(Vitals.HP) = 1 Or Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    target = index
                    targetType = TARGET_TYPE_PLAYER
                    HasBuffered = True
                Else
                    ' if have target, check in range
                    If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(MapNum).Npc(target).x, MapNpc(MapNum).Npc(target).y) Then
                        PlayerMsg index, "Target not in range.", BrightRed
                        HasBuffered = False
                    Else
                        If Spell(spellnum).VitalType(Vitals.HP) = 0 Or Spell(spellnum).VitalType(Vitals.MP) = 0 Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                HasBuffered = True
                            End If
                        Else
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = timeGetTime
        TempPlayer(index).spellBuffer.target = target
        TempPlayer(index).spellBuffer.tType = targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub NpcBufferSpell(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal npcSpellSlot As Long)
Dim spellnum As Long, mpCost As Long, Range As Long, HasBuffered As Boolean, targetType As Byte, target As Long, spellCastType As Long, i As Long

    ' prevent rte9
    If npcSpellSlot <= 0 Or npcSpellSlot > MAX_NPC_SPELLS Then Exit Sub
    
    With MapNpc(MapNum).Npc(MapNpcNum)
        ' set the spell number
        spellnum = Npc(.Num).Spell(npcSpellSlot)
        
        ' prevent rte9
        If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
        
        ' make sure we're not already buffering
        If .spellBuffer.Spell > 0 Then Exit Sub
        
        ' see if cooldown as finished
        If .SpellCD(npcSpellSlot) > timeGetTime Then Exit Sub
        
        ' Set the MP Cost
        mpCost = Spell(spellnum).mpCost
        
        ' have they got enough mp?
        If .Vital(Vitals.MP) < mpCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(spellnum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellnum).IsAoE Then
                spellCastType = 2 ' targetted
            Else
                spellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellnum).IsAoE Then
                spellCastType = 0 ' self-cast
            Else
                spellCastType = 1 ' self-cast AoE
            End If
        End If
        
        targetType = .targetType
        target = .target
        Range = Spell(spellnum).Range
        HasBuffered = False
        
        ' make sure on the map
        If GetPlayerMap(target) <> MapNum Then Exit Sub
        
        Select Case spellCastType
            Case 0, 1 ' self-cast & self-cast AOE
                HasBuffered = True
            Case 2, 3 ' targeted & targeted AOE
                ' if it's a healing spell then heal a friend
                If Spell(spellnum).VitalType(Vitals.HP) = 1 Or Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    ' find a friend who needs healing
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            targetType = TARGET_TYPE_NPC
                            target = i
                            HasBuffered = True
                        End If
                    Next
                Else
                    ' check if have target
                    If Not target > 0 Then Exit Sub
                    ' make sure it's a player
                    If targetType = TARGET_TYPE_PLAYER Then
                        ' if have target, check in range
                        If Not isInRange(Range, .x, .y, GetPlayerX(target), GetPlayerY(target)) Then
                            Exit Sub
                        Else
                            If CanNpcAttackPlayer(MapNpcNum, target, True) Then
                                HasBuffered = True
                            End If
                        End If
                    End If
                End If
        End Select
        
        If HasBuffered Then
            SendAnimation MapNum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum
            .spellBuffer.Spell = npcSpellSlot
            .spellBuffer.Timer = timeGetTime
            .spellBuffer.target = target
            .spellBuffer.tType = targetType
        End If
    End With
End Sub

Public Sub NpcCastSpell(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Long)
Dim spellnum As Long, mpCost As Long, Vital As Long, DidCast As Boolean, i As Long, AoE As Long, Range As Long, VitalType As Byte, increment As Boolean, x As Long, y As Long, spellCastType As Long

    DidCast = False
    
    ' rte9
    If spellslot <= 0 Or spellslot > MAX_NPC_SPELLS Then Exit Sub
    
    With MapNpc(MapNum).Npc(MapNpcNum)
        ' cache spell num
        spellnum = Npc(.Num).Spell(spellslot)
        
        ' cache mp cost
        mpCost = Spell(spellnum).mpCost
        
        ' make sure still got enough mp
        If .Vital(Vitals.MP) < mpCost Then Exit Sub
        
        ' find out what kind of spell it is! self cast, target or AOE
        If Spell(spellnum).Range > 0 Then
            ' ranged attack, single target or aoe?
            If Not Spell(spellnum).IsAoE Then
                spellCastType = 2 ' targetted
            Else
                spellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(spellnum).IsAoE Then
                spellCastType = 0 ' self-cast
            Else
                spellCastType = 1 ' self-cast AoE
            End If
        End If
        
        ' store data
        AoE = Spell(spellnum).AoE
        Range = Spell(spellnum).Range
        
        Select Case spellCastType
            Case 0 ' self-cast target
                If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                    SpellNpc_Effect Vitals.HP, True, MapNpcNum, Vital, spellnum, MapNum
                End If
                If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                    SpellNpc_Effect Vitals.MP, True, MapNpcNum, Vital, spellnum, MapNum
                End If
            Case 1, 3 ' self-cast AOE & targetted AOE
                If spellCastType = 1 Then
                    x = .x
                    y = .y
                ElseIf spellCastType = 3 Then
                    If targetType = 0 Then Exit Sub
                    If target = 0 Then Exit Sub
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        x = GetPlayerX(target)
                        y = GetPlayerY(target)
                    Else
                        x = MapNpc(MapNum).Npc(target).x
                        y = MapNpc(MapNum).Npc(target).y
                    End If
                    
                    If Not isInRange(Range, .x, .y, x, y) Then Exit Sub
                End If
                If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If isInRange(AoE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNpcNum, i, True) Then
                                        Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                                        NpcAttackPlayer MapNpcNum, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If Spell(spellnum).VitalType(Vitals.MP) = 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If isInRange(AoE, .x, .y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanNpcAttackPlayer(MapNpcNum, i, True) Then
                                        Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                                        NpcAttackPlayer MapNpcNum, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                                    SpellNpc_Effect Vitals.HP, True, i, Vital, spellnum, MapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
                If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(MP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                                    SpellNpc_Effect Vitals.MP, True, i, Vital, spellnum, MapNum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                End If
            Case 2 ' targetted
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(MapNum).Npc(target).x
                    y = MapNpc(MapNum).Npc(target).y
                End If
                    
                If Not isInRange(Range, .x, .y, x, y) Then Exit Sub
                
                If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanNpcAttackPlayer(MapNpcNum, target, True) Then
                            Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                            NpcAttackPlayer MapNpcNum, target, Vital, spellnum
                            DidCast = True
                        End If
                    End If
                End If
                
                If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetNpcSpellDamage(.Num, spellnum, HP)
                        SpellNpc_Effect Vitals.HP, True, target, Vital, spellnum, MapNum
                        DidCast = True
                    End If
                End If
                
                If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                    If targetType = TARGET_TYPE_NPC Then
                        Vital = GetNpcSpellDamage(.Num, spellnum, MP)
                        SpellNpc_Effect Vitals.MP, True, target, Vital, spellnum, MapNum
                        DidCast = True
                    End If
                End If
        End Select
        
        If DidCast Then
            .Vital(Vitals.MP) = .Vital(Vitals.MP) - mpCost
            .SpellCD(spellslot) = timeGetTime + (Spell(spellnum).CDTime * 1000)
        End If
    End With
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
Dim spellnum As Long, mpCost As Long, LevelReq As Long, MapNum As Long, Vital As Long, DidCast As Boolean, ClassReq As Long
Dim AccessReq As Long, i As Long, AoE As Long, Range As Long, VitalType As Byte, increment As Boolean, x As Long, y As Long
Dim Buffer As clsBuffer, spellCastType As Long, Rotate As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = Player(index).Spell(spellslot)
    MapNum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub

    mpCost = Spell(spellnum).mpCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < mpCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            spellCastType = 2 ' targetted
        Else
            spellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            spellCastType = 0 ' self-cast
        Else
            spellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' store data
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
    
    Select Case spellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_VITALCHANGE
                    If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                        Vital = GetPlayerSpellDamage(index, spellnum, HP)
                        SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                        DidCast = True
                    End If
                    If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                        Vital = GetPlayerSpellDamage(index, spellnum, MP)
                        SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                        DidCast = True
                    End If
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SendEffect MapNum, Spell(spellnum).Effect, GetPlayerX(index), GetPlayerY(index), TARGET_TYPE_PLAYER, index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If spellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf spellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(MapNum).Npc(target).x
                    y = MapNpc(MapNum).Npc(target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If

                Rotate = Engine_GetAngle(GetPlayerX(index), GetPlayerY(index), x, y)
            
                ' ****** Set Player Direction Based On Angle ******
                If Rotate >= 315 And Rotate <= 360 Then
                    Call SetPlayerDir(index, DIR_UP)
                ElseIf Rotate >= 0 And Rotate <= 45 Then
                    Call SetPlayerDir(index, DIR_UP)
                ElseIf Rotate >= 225 And Rotate <= 315 Then
                    Call SetPlayerDir(index, DIR_LEFT)
                ElseIf Rotate >= 135 And Rotate <= 225 Then
                    Call SetPlayerDir(index, DIR_DOWN)
                ElseIf Rotate >= 45 And Rotate <= 135 Then
                    Call SetPlayerDir(index, DIR_RIGHT)
                End If
            End If
            
            If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                Vital = GetPlayerSpellDamage(index, spellnum, HP)
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If i <> index Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    If CanPlayerAttackPlayer(index, i, True) Then
                                        PlayerAttackPlayer index, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        PlayerAttackNpc index, i, Vital, spellnum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    Next
            End If
            
            If Spell(spellnum).VitalType(Vitals.MP) = 0 Then
                Vital = GetPlayerSpellDamage(index, spellnum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, False, i, Vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
                    
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(MapNum).Npc(i).Num > 0 Then
                                If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                    If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                        SpellNpc_Effect Vitals.MP, False, i, Vital, spellnum, MapNum
                                        DidCast = True
                                    End If
                                End If
                            End If
                        Next
            End If
            
            If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                Vital = GetPlayerSpellDamage(index, spellnum, HP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.HP, True, i, Vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
            
            If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                Vital = GetPlayerSpellDamage(index, spellnum, MP)
                For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect Vitals.MP, True, i, Vital, spellnum
                                    DidCast = True
                                End If
                            End If
                        End If
                    Next
            End If
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(MapNum).Npc(target).x
                y = MapNpc(MapNum).Npc(target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            Rotate = Engine_GetAngle(GetPlayerX(index), GetPlayerY(index), x, y)
            
            ' ****** Set Player Direction Based On Angle ******
            If Rotate >= 315 And Rotate <= 360 Then
                Call SetPlayerDir(index, DIR_UP)
            ElseIf Rotate >= 0 And Rotate <= 45 Then
                Call SetPlayerDir(index, DIR_UP)
            ElseIf Rotate >= 225 And Rotate <= 315 Then
                Call SetPlayerDir(index, DIR_LEFT)
            ElseIf Rotate >= 135 And Rotate <= 225 Then
                Call SetPlayerDir(index, DIR_DOWN)
            ElseIf Rotate >= 45 And Rotate <= 135 Then
                Call SetPlayerDir(index, DIR_RIGHT)
            End If
            
            If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                Vital = GetPlayerSpellDamage(index, spellnum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If Vital > 0 Then
                                PlayerAttackPlayer index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                PlayerAttackNpc index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
            End If
            
            If Spell(spellnum).VitalType(Vitals.MP) = 0 Then
                Vital = GetPlayerSpellDamage(index, spellnum, MP)
                    If targetType = TARGET_TYPE_PLAYER Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect Vitals.MP, False, target, Vital, spellnum
                                DidCast = True
                            End If
                    Else
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect Vitals.MP, False, target, Vital, spellnum, MapNum
                                DidCast = True
                            End If
                    End If
            End If
            
            If Spell(spellnum).VitalType(Vitals.HP) = 1 Then
                Vital = GetPlayerSpellDamage(index, spellnum, HP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.HP, True, target, Vital, spellnum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.HP, True, target, Vital, spellnum, MapNum
                    DidCast = True
                End If
            End If
            
            If Spell(spellnum).VitalType(Vitals.MP) = 1 Then
                Vital = GetPlayerSpellDamage(index, spellnum, MP)
                If targetType = TARGET_TYPE_PLAYER Then
                    SpellPlayer_Effect Vitals.MP, True, target, Vital, spellnum
                    DidCast = True
                Else
                    SpellNpc_Effect Vitals.MP, True, target, Vital, spellnum, MapNum
                    DidCast = True
                End If
            End If
    End Select
    
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - mpCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
        TempPlayer(index).SpellCD(spellslot) = timeGetTime + (Spell(spellnum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
    End If
End Sub
Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendEffect GetPlayerMap(index), Spell(spellnum).Effect, GetPlayerX(index), GetPlayerY(index), TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Player index, spellnum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
        
        ' send update
        SendVital index, Vital
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
        
        SendAnimation MapNum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendEffect MapNum, Spell(spellnum).Effect, MapNpc(MapNum).Npc(index).x, MapNpc(MapNum).Npc(index).y, TARGET_TYPE_NPC, index
        SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
        
        ' send the sound
        SendMapSound index, MapNpc(MapNum).Npc(index).x, MapNpc(MapNum).Npc(index).y, SoundEntity.seSpell, spellnum
        
        If increment Then
            MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) + Damage
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Npc MapNum, index, spellnum
            End If
        ElseIf Not increment Then
            MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = spellnum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = spellnum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).DoT(i)
            If .Spell = spellnum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellnum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).HoT(i)
            If .Spell = spellnum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, GetPlayerSpellDamage(.Caster, .Spell, HP)
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(index).Map, "+" & GetPlayerSpellDamage(.Caster, .Spell, HP), BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell, HP)
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).Npc(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, GetPlayerSpellDamage(.Caster, .Spell, HP), , True
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(MapNum).Npc(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg MapNum, "+" & GetPlayerSpellDamage(.Caster, .Spell, HP), BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
                MapNpc(MapNum).Npc(index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(index).Vital(Vitals.HP) + GetPlayerSpellDamage(.Caster, .Spell, HP)
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(spellnum).StunDuration
        TempPlayer(index).StunTimer = timeGetTime
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal MapNum As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(index).StunDuration = Spell(spellnum).StunDuration
        MapNpc(MapNum).Npc(index).StunTimer = timeGetTime
    End If
End Sub

Public Sub TryPlayerShootNpc(ByVal index As Long, ByVal MapNpcNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long
Dim n As Long
Dim Stat As Stats

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootNpc(index, MapNpcNum) Then
    
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
        Call CreateProjectile(MapNum, index, TARGET_TYPE_PLAYER, MapNpcNum, TARGET_TYPE_NPC, Item(GetPlayerEquipment(index, Weapon)).Projectile, Item(GetPlayerEquipment(index, Weapon)).Rotation)
        ' check if NPC cafn avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (Npc(NPCNum).Stat(Stats.Endurance) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, MapNpcNum, Damage, -1)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Public Function CanPlayerShootNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If Not isInRange(Item(GetPlayerEquipment(Attacker, Weapon)).Range, GetPlayerX(Attacker), GetPlayerY(Attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y) Then Exit Function
            attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NPCNum > 0 And timeGetTime > TempPlayer(Attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
            End Select
            
            If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
                TempPlayer(Attacker).target = MapNpcNum
                SendTarget Attacker
                CanPlayerShootNpc = True
            Else
                If NpcX = GetPlayerX(Attacker) Then
                    If NpcY = GetPlayerY(Attacker) Then
                         If Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            
                            If Npc(NPCNum).Event > 0 Then
                                InitEvent Attacker, Npc(NPCNum).Event
                                Exit Function
                            End If
                        End If
                        If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                            Call SendChatBubble(MapNum, MapNpcNum, TARGET_TYPE_NPC, Trim$(Npc(NPCNum).AttackSay), DarkBrown)
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub TryNpcShootNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal Victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, blockAmount As Long, Damage As Long
    
    ' Can the npc attack the player?
    If CanNpcShootNPC(MapNum, Attacker, Victim) Then
        aNpcNum = MapNpc(MapNum).Npc(Attacker).Num
        vNpcNum = MapNpc(MapNum).Npc(Victim).Num
        Call CreateProjectile(MapNum, Attacker, TARGET_TYPE_NPC, Victim, TARGET_TYPE_NPC, Npc(aNpcNum).Projectile, Npc(aNpcNum).Rotation)
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(Victim).x * 32), (MapNpc(MapNum).Npc(Victim).y * 32)
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(Victim).x * 32), (MapNpc(MapNum).Npc(Victim).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (Npc(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(Attacker).x * 32), (MapNpc(MapNum).Npc(Attacker).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(MapNum, Attacker, Victim, Damage)
        End If
    End If
End Sub

Public Sub TryNpcShootPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
Dim MapNum As Long, NPCNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcShootPlayer(MapNpcNum, index) Then
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num
        Call CreateProjectile(MapNum, MapNpcNum, TARGET_TYPE_NPC, index, TARGET_TYPE_PLAYER, Npc(NPCNum).Projectile, Npc(NPCNum).Rotation)
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(index, Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(NPCNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, Damage)
        End If
    End If
End Sub

Public Sub TryPlayerShootPlayer(ByVal Attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long
Dim n As Long
Dim Stat As Stats

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerShootPlayer(Attacker, Victim) Then
    
        MapNum = GetPlayerMap(Attacker)
        Call CreateProjectile(MapNum, Attacker, TARGET_TYPE_PLAYER, Victim, TARGET_TYPE_PLAYER, Item(GetPlayerEquipment(Attacker, Weapon)).Projectile, Item(GetPlayerEquipment(Attacker, Weapon)).Rotation)
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(Victim, Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(Attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage, -1)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub
Function CanPlayerShootPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean

    ' Check attack timer
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
    Else
        If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function
    TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(Attacker).target = Victim
    SendTarget Attacker
    CanPlayerShootPlayer = True
End Function
Function CanNpcShootNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim aNpcNum As Long, vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    ' Check for subscript out of range
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Function
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Function

    aNpcNum = MapNpc(MapNum).Npc(Attacker).Num
    vNpcNum = MapNpc(MapNum).Npc(Victim).Num
    
    ' Check for subscript out of range
    If aNpcNum <= 0 Then
        Exit Function
    End If
    
    If vNpcNum <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(MapNum).Npc(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(MapNum).Npc(Attacker).AttackTimer = timeGetTime

    AttackerX = MapNpc(MapNum).Npc(Attacker).x
    AttackerY = MapNpc(MapNum).Npc(Attacker).y
    VictimX = MapNpc(MapNum).Npc(Victim).x
    VictimY = MapNpc(MapNum).Npc(Victim).y
    
    If isInRange(Npc(aNpcNum).ProjectileRange, AttackerX, AttackerY, VictimX, VictimY) Then
        CanNpcShootNPC = True
    End If
End Function

Function CanNpcShootPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim MapNum As Long
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(index)
    NPCNum = MapNpc(MapNum).Npc(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NPCNum > 0 Then
            If isInRange(Npc(NPCNum).ProjectileRange, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, GetPlayerX(index), GetPlayerY(index)) Then
                CanNpcShootPlayer = True
            End If
        End If
    End If
End Function

Sub CreateProjectile(ByVal MapNum As Long, ByVal Attacker As Long, ByVal AttackerType As Long, ByVal Victim As Long, ByVal targetType As Long, ByVal Graphic As Long, ByVal RotateSpeed As Long)
Dim Rotate As Long
Dim Buffer As clsBuffer
    
    If AttackerType = TARGET_TYPE_PLAYER Then
        ' ****** Initial Rotation Value ******
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), GetPlayerX(Victim), GetPlayerY(Victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(GetPlayerX(Attacker), GetPlayerY(Attacker), MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y)
        End Select
    
        ' ****** Set Player Direction Based On Angle ******
        If Rotate >= 315 And Rotate <= 360 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 0 And Rotate <= 45 Then
            Call SetPlayerDir(Attacker, DIR_UP)
        ElseIf Rotate >= 225 And Rotate <= 315 Then
            Call SetPlayerDir(Attacker, DIR_LEFT)
        ElseIf Rotate >= 135 And Rotate <= 225 Then
            Call SetPlayerDir(Attacker, DIR_DOWN)
        ElseIf Rotate >= 45 And Rotate <= 135 Then
            Call SetPlayerDir(Attacker, DIR_RIGHT)
        End If
        
        Set Buffer = New clsBuffer
        Buffer.WriteLong SPlayerDir
        Buffer.WriteLong Attacker
        Buffer.WriteLong GetPlayerDir(Attacker)
        Call SendDataToMap(MapNum, Buffer.ToArray())
        Set Buffer = Nothing
    ElseIf AttackerType = TARGET_TYPE_NPC Then
        Select Case targetType
            Case TARGET_TYPE_PLAYER
                Rotate = Engine_GetAngle(MapNpc(MapNum).Npc(Attacker).x, MapNpc(MapNum).Npc(Attacker).y, GetPlayerX(Victim), GetPlayerY(Victim))
            Case TARGET_TYPE_NPC
                Rotate = Engine_GetAngle(MapNpc(MapNum).Npc(Attacker).x, MapNpc(MapNum).Npc(Attacker).y, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y)
        End Select
    End If

    Call SendProjectile(MapNum, Attacker, AttackerType, Victim, targetType, Graphic, Rotate, RotateSpeed)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetX As Integer, ByVal targetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targetY Then
        'Check for going right (90 degrees)
        If CenterX < targetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetX Then
        'Check for going up (360 degrees)
        If CenterY > targetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetX - CenterX) ^ 2 + Abs(targetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetX - CenterX) ^ 2 + targetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function
    Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function
End Function

Sub NpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim i As Long, n As Long
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Sub
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Sub

    aNpcNum = MapNpc(MapNum).Npc(Attacker).Num
    vNpcNum = MapNpc(MapNum).Npc(Victim).Num
    
    ' Check for subscript out of range
    If aNpcNum <= 0 Then
        Exit Sub
    End If
    
    If vNpcNum <= 0 Then
        Exit Sub
    End If
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong Attacker
    SendDataToMap MapNum, Buffer.ToArray()
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(Attacker).stopRegen = True
    MapNpc(MapNum).Npc(Attacker).stopRegenTimer = timeGetTime
    
    ' Say damage
    SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(Victim).x * 32), (MapNpc(MapNum).Npc(Victim).y * 32)
    SendBlood MapNum, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y
    
    Call SendAnimation(MapNum, Npc(MapNpc(MapNum).Npc(Attacker).Num).Animation, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y, TARGET_TYPE_NPC, Victim)
    Call SendEffect(MapNum, Npc(MapNpc(MapNum).Npc(Attacker).Num).Effect, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y, TARGET_TYPE_NPC, Victim)
    
    ' send the sound
    SendMapSound Victim, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y, SoundEntity.seNpc, MapNpc(MapNum).Npc(Attacker).Num
    
    If Damage >= MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) Then
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(vNpcNum).DropItem(n) = 0 Then Exit For
        
            If Rnd <= Npc(vNpcNum).DropChance(n) Then
                Call SpawnItem(Npc(vNpcNum).DropItem(n), Npc(vNpcNum).DropItemValue(n), MapNum, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y)
            End If
        Next
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(Victim).Num = 0
        MapNpc(MapNum).Npc(Victim).SpawnWait = timeGetTime
        MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(Victim).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(Victim).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong Victim
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = Victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) = MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) - Damage
        
        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(Victim).targetType = TARGET_TYPE_NPC
        MapNpc(MapNum).Npc(Victim).target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(Victim).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(Victim).Num Then
                    MapNpc(MapNum).Npc(i).target = Attacker
                    MapNpc(MapNum).Npc(i).targetType = TARGET_TYPE_NPC
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(Victim).stopRegen = True
        MapNpc(MapNum).Npc(Victim).stopRegenTimer = timeGetTime
        
        SendMapNpcVitals MapNum, Victim
    End If
    MapNpc(MapNum).Npc(Attacker).AttackTimer = timeGetTime
End Sub

Function CanNpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim aNpcNum As Long, vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long

    ' Check for subscript out of range
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Function
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Function

    aNpcNum = MapNpc(MapNum).Npc(Attacker).Num
    vNpcNum = MapNpc(MapNum).Npc(Victim).Num
    
    ' Check for subscript out of range
    If aNpcNum <= 0 Then
        Exit Function
    End If
    
    If vNpcNum <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(Attacker).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    If MapNpc(MapNum).Npc(Victim).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(MapNum).Npc(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If

    MapNpc(MapNum).Npc(Attacker).AttackTimer = timeGetTime

    AttackerX = MapNpc(MapNum).Npc(Attacker).x
    AttackerY = MapNpc(MapNum).Npc(Attacker).y
    VictimX = MapNpc(MapNum).Npc(Victim).x
    VictimY = MapNpc(MapNum).Npc(Victim).y
    ' Check if at same coordinates
    If (VictimY + 1 = AttackerY) And (VictimX = AttackerX) Then
        CanNpcAttackNPC = True
    Else
        If (VictimY - 1 = AttackerY) And (VictimX = AttackerX) Then
            CanNpcAttackNPC = True
        Else
            If (VictimY = AttackerY) And (VictimX + 1 = AttackerX) Then
                CanNpcAttackNPC = True
            Else
                If (VictimY = AttackerY) And (VictimX - 1 = AttackerX) Then
                    CanNpcAttackNPC = True
                End If
            End If
        End If
    End If
End Function

Public Sub TryNpcAttackNPC(ByVal MapNum As Long, ByVal Attacker As Long, ByVal Victim As Long)
Dim aNpcNum As Long, vNpcNum As Long, blockAmount As Long, Damage As Long
    
    ' Can the npc attack the player?
    If CanNpcAttackNPC(MapNum, Attacker, Victim) Then
        aNpcNum = MapNpc(MapNum).Npc(Attacker).Num
        vNpcNum = MapNpc(MapNum).Npc(Victim).Num
        
        ' check if NPC can avoid the attack
        If CanNpcDodge(vNpcNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(Victim).x * 32), (MapNpc(MapNum).Npc(Victim).y * 32)
            Exit Sub
        End If
        If CanNpcParry(vNpcNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(Victim).x * 32), (MapNpc(MapNum).Npc(Victim).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(aNpcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanNpcBlock(vNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (Npc(vNpcNum).Stat(Stats.Endurance) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(aNpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(Attacker).x * 32), (MapNpc(MapNum).Npc(Attacker).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackNPC(MapNum, Attacker, Victim, Damage)
        End If
    End If
End Sub
