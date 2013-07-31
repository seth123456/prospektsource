Attribute VB_Name = "modAnimationEditor"
Option Explicit

Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public TempAnimation As AnimationRec

' Animation
Public Const AnimColumns As Long = 5

Public Type AnimationRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    For I = 0 To 1
        frmEditor_Animation.scrlSprite(I).Max = Count_Anim
    Next

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For I = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(I) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For I = 0 To 1
            frmEditor_Animation.scrlSprite(I).Value = .Sprite(I)
            frmEditor_Animation.scrlFrameCount(I).Value = .Frames(I)
            frmEditor_Animation.scrlLoopCount(I).Value = .LoopCount(I)
            
            If .looptime(I) > 0 Then
                frmEditor_Animation.scrlLoopTime(I).Value = .looptime(I)
            Else
                frmEditor_Animation.scrlLoopTime(I).Value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_ANIMATIONS
        If Animation_Changed(I) Then
            Call SendSaveAnimation(I)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).name = vbNullString
    Animation(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_ANIMATIONS
        Call ClearAnimation(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
