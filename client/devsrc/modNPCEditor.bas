Attribute VB_Name = "modNPCEditor"
Option Explicit

Public Npc(1 To MAX_NPCS) As NpcRec
Public TempNpc As NpcRec

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

Public Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    
    ' Casting
    Spell(1 To MAX_NPC_SPELLS) As Long
    Event As Long
    Projectile As Long
    ProjectileRange As Byte
    Rotation As Integer
    Moral As Byte
    Effect As Long
End Type

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1

    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .scrlSprite.Max = Count_Char
        .scrlAnimation.Max = MAX_ANIMATIONS
        .scrlEvent.Max = MAX_EVENTS
        .scrlProjectilePic.Max = Count_Projectile
        .scrlEffect.Max = MAX_EFFECTS
        .scrlNum.Max = MAX_ITEMS
        .scrlDrop.Max = MAX_NPC_DROPS
        .scrlSpell.Max = MAX_NPC_SPELLS
        .txtName.Text = Trim$(Npc(EditorIndex).name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.Max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.Value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtHP.Text = Npc(EditorIndex).HP
        .txtEXP.Text = Npc(EditorIndex).EXP
        .txtLevel.Text = Npc(EditorIndex).Level
        .txtDamage.Text = Npc(EditorIndex).Damage
        .scrlEvent.Value = Npc(EditorIndex).Event
        .scrlProjectilePic.Value = Npc(EditorIndex).Projectile
        .scrlProjectileRange.Value = Npc(EditorIndex).ProjectileRange
        .scrlProjectileRotation.Value = Npc(EditorIndex).Rotation
        .cmbMoral.ListIndex = Npc(EditorIndex).Moral
        .scrlEffect.Value = Npc(EditorIndex).Effect
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Npc(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For I = 1 To Stats.Stat_Count - 1
            .scrlStat(I).Value = Npc(EditorIndex).Stat(I)
        Next
        
        ' show 1 data
        .scrlDrop.Value = 1
        .scrlSpell.Value = 1
    End With
    
    Npc_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_NPCS
        If Npc_Changed(I) Then
            Call SendSaveNpc(I)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Npc_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).name = vbNullString
    Npc(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_NPCS
        Call ClearNPC(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
