Attribute VB_Name = "modSpellEditor"
Option Explicit

Public Spell(1 To MAX_SPELLS) As SpellRec
Public TempSpell As SpellRec

' Spell constants
Public Const SPELL_TYPE_VITALCHANGE As Byte = 0
Public Const SPELL_TYPE_WARP As Byte = 1

Public Type SpellRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    VitalType(1 To Vitals.Vital_Count - 1) As Byte
    Effect As Long
End Type


' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1

    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        .scrlEffect.Max = MAX_EFFECTS
        .scrlIcon.Max = Count_Spellicon
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For I = 1 To MAX_CLASSES
            .cmbClass.AddItem Trim$(Class(I).name)
        Next
        .cmbClass.ListIndex = 0
        
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).Dir
        .txtHPVital.Text = Spell(EditorIndex).Vital(Vitals.HP)
        .txtMPVital.Text = Spell(EditorIndex).Vital(Vitals.MP)
        .optHPVital(Spell(EditorIndex).VitalType(Vitals.HP)).Value = True
        .optMPVital(Spell(EditorIndex).VitalType(Vitals.MP)).Value = True
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        .scrlEffect.Value = Spell(EditorIndex).Effect
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_SPELLS
        If Spell_Changed(I) Then
            Call SendSaveSpell(I)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
