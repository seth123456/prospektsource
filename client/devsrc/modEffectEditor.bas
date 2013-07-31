Attribute VB_Name = "modEffectEditor"
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public TempEffect As EffectRec
Public EditorEffectData As Effect

'Constants With The Order Number For Each Effect
Public Const EFFECT_TYPE_HEAL As Byte = 1
Public Const EFFECT_TYPE_PROTECTION As Byte = 2
Public Const EFFECT_TYPE_STRENGTHEN As Byte = 3
Public Const EFFECT_TYPE_SUMMON As Byte = 4

Public Type EffectRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    isMulti As Byte
    MultiParticle(1 To MAX_MULTIPARTICLE) As Long
    Type As Long
    Sprite As Long
    Particles As Long
    Size As Single
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Duration As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    Modifier As Byte
End Type

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean 'If the effect is in use
    Alpha As Single
    Decay As Single
    Red As Single
    Green As Single
    Blue As Single
    XSpeed As Single
    YSpeed As Single
    XAcc As Single
    YAcc As Single
    Effectnum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Particles() As clsParticle  'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles ' Cant use in .NET maybe change
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindType As Byte
    BindIndex As Long       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
End Type

' /////////////////
' // Effect Editor //
' /////////////////
Public Sub EffectEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If frmEditor_Effect.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Effect.lstIndex.ListIndex + 1

    ' add the array to the combo
    frmEditor_Effect.cmbSound.Clear
    frmEditor_Effect.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_Effect.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    ' set max values
    frmEditor_Effect.scrlSprite.Max = Count_Particle
    frmEditor_Effect.scrlMultiParticle.Max = MAX_MULTIPARTICLE
    frmEditor_Effect.scrlEffect.Max = MAX_EFFECTS

    With Effect(EditorIndex)
        frmEditor_Effect.txtName.Text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Effect.cmbSound.ListCount >= 0 Then
            For I = 0 To frmEditor_Effect.cmbSound.ListCount
                If frmEditor_Effect.cmbSound.List(I) = Trim$(.sound) Then
                    frmEditor_Effect.cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Effect.cmbSound.ListIndex = -1 Then frmEditor_Effect.cmbSound.ListIndex = 0
        End If
        frmEditor_Effect.scrlSprite.Value = .Sprite
        frmEditor_Effect.cmbType.ListIndex = .Type - 1
        frmEditor_Effect.scrlParticles.Value = .Particles
        frmEditor_Effect.scrlSize.Value = .Size
        frmEditor_Effect.scrlAlpha.Value = .Alpha
        frmEditor_Effect.scrlDecay.Value = .Decay
        frmEditor_Effect.scrlRed.Value = .Red
        frmEditor_Effect.scrlGreen.Value = .Green
        frmEditor_Effect.scrlBlue.Value = .Blue
        frmEditor_Effect.scrlXSpeed.Value = .XSpeed
        frmEditor_Effect.scrlYSpeed.Value = .YSpeed
        frmEditor_Effect.scrlXAcc.Value = .XAcc
        frmEditor_Effect.scrlYAcc.Value = .YAcc
        frmEditor_Effect.scrlModifier.Value = .Modifier
        frmEditor_Effect.optEffectType(.isMulti) = True
        If .isMulti = 1 Then
            frmEditor_Effect.fraMultiParticle.Visible = True
            frmEditor_Effect.fraEffect.Visible = False
            frmEditor_Effect.scrlEffect.Value = .MultiParticle(1)
        Else
            frmEditor_Effect.fraEffect.Visible = True
            frmEditor_Effect.fraMultiParticle.Visible = False
        End If
        frmEditor_Effect.scrlDuration = .Duration
        GDIBeginEffect
        EditorIndex = frmEditor_Effect.lstIndex.ListIndex + 1
    End With
    
    Effect_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub EffectEditorOk()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_EFFECTS
        If Effect_Changed(I) Then
            Call SendSaveEffect(I)
        End If
    Next
    
    Unload frmEditor_Effect
    Editor = 0
    ClearChanged_Effect
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub EffectEditorCancel()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Effect
    ClearChanged_Effect
    ClearEffects
    SendRequestEffects
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EffectEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Effect()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ZeroMemory Effect_Changed(1), MAX_EFFECTS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Effect", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearEffect(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Effect(Index)), LenB(Effect(Index)))
    Effect(Index).name = vbNullString
    Effect(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEffect", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearEffects()
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    For I = 1 To MAX_EFFECTS
        Call ClearEffect(I)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearEffects", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub GDIUpdateEffectOffset(ByVal EffectIndex As Integer)
    EditorEffectData.X = EditorEffectData.X + (LastOffsetX - ParticleOffsetX)
    EditorEffectData.Y = EditorEffectData.Y + (LastOffsetY - ParticleOffsetY)
End Sub

Sub GDIEffect_Kill(ByVal EffectIndex As Integer)
    'Stop The Selected Effect
    EditorEffectData.Used = False
End Sub

Public Sub GDIRenderEditorEffectData(Optional ByVal SetRenderStates As Boolean = True)
Dim sRect As GeomRec
    
    'Check if we have the device
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    If EditorEffectData.Gfx <= 0 Or EditorEffectData.Gfx > Count_Particle Then Exit Sub
    
    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene

    'Set the render state for the size of the particle
    Call D3DDevice8.SetRenderState(D3DRS_POINTSIZE, EditorEffectData.FloatSize)
    
    'Set the render state to point blitting
    If SetRenderStates Then D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the texture
    Directx8.SetTexture Tex_Particle(EditorEffectData.Gfx)
    D3DDevice8.SetTexture 0, gTexture(Tex_Particle(EditorEffectData.Gfx)).Texture

    'Draw all the particles at once
    D3DDevice8.DrawPrimitiveUP D3DPT_POINTLIST, EditorEffectData.ParticleCount, EditorEffectData.PartVertex(0), Len(EditorEffectData.PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    sRect.Top = 0
    sRect.Left = 0
    sRect.Width = frmEditor_Effect.picEffect.Width
    sRect.Height = frmEditor_Effect.picEffect.Height
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, frmEditor_Effect.picEffect.hWnd, ByVal 0)
End Sub

Sub GDIUpdateEffectAll()
'Updates all of the effects and renders them

    'Make sure the effect is in use
    If EditorEffectData.Used Then
            
        If EditorEffectData.Progression = 0 Then GDIEffect_Kill I
        'Find out which effect is selected, then update it
        GDIUpdateEffect
        'Render the effect
        Call GDIRenderEditorEffectData
    End If
End Sub

Public Sub GDIBeginEffect()
Dim EffectIndex As Integer
Dim I As Long

    'Clear the old information from the effect
    Erase EditorEffectData.Particles()
    Erase EditorEffectData.PartVertex()
    ZeroMemory EditorEffectData, LenB(EditorEffectData)
    EditorEffectData.GoToX = -30000
    EditorEffectData.GoToY = -30000

    'Set The Effect's Variables
    EditorEffectData.Effectnum = frmEditor_Effect.cmbType.ListIndex + 1 'Set the effect number
    EditorEffectData.Used = True 'Enabled the effect
    EditorEffectData.X = frmEditor_Effect.picEffect.ScaleWidth / 2 'Set the effect's X coordinate
    EditorEffectData.Y = frmEditor_Effect.picEffect.ScaleHeight / 2 'Set the effect's Y coordinate
    EditorEffectData.ParticleCount = frmEditor_Effect.scrlParticles 'Set the number of particles
    EditorEffectData.Gfx = frmEditor_Effect.scrlSprite 'Set the graphic
    EditorEffectData.Alpha = frmEditor_Effect.scrlAlpha
    EditorEffectData.Decay = frmEditor_Effect.scrlDecay
    EditorEffectData.Red = frmEditor_Effect.scrlRed
    EditorEffectData.Green = frmEditor_Effect.scrlGreen
    EditorEffectData.Blue = frmEditor_Effect.scrlBlue
    EditorEffectData.XSpeed = frmEditor_Effect.scrlXSpeed
    EditorEffectData.YSpeed = frmEditor_Effect.scrlYSpeed
    EditorEffectData.XAcc = frmEditor_Effect.scrlXAcc
    EditorEffectData.YAcc = frmEditor_Effect.scrlYAcc
    EditorEffectData.Modifier = frmEditor_Effect.scrlModifier 'How large the circle is
    EditorEffectData.Progression = frmEditor_Effect.scrlDuration 'How long the effect will last
    EditorEffectData.FloatSize = Effect_FToDW(frmEditor_Effect.scrlSize) 'Size of the particles
    'Set the number of particles left to the total avaliable
    EditorEffectData.ParticlesLeft = EditorEffectData.ParticleCount
    'Redim the number of particles
    ReDim EditorEffectData.Particles(0 To EditorEffectData.ParticleCount)
    ReDim EditorEffectData.PartVertex(0 To EditorEffectData.ParticleCount)
    'Create the particles
    For I = 0 To EditorEffectData.ParticleCount
        Set EditorEffectData.Particles(I) = New clsParticle
        EditorEffectData.Particles(I).Used = True
        EditorEffectData.PartVertex(I).RHW = 1
        GDIResetEffect I
    Next I
    'Set The Initial Time
    EditorEffectData.PreviousFrame = timeGetTime
End Sub

Private Sub GDIUpdateEffect()
Dim ElapsedTime As Single
Dim I As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - EditorEffectData.PreviousFrame) * 0.01
    EditorEffectData.PreviousFrame = timeGetTime
    If EditorEffectData.Progression > 0 Then EditorEffectData.Progression = EditorEffectData.Progression - ElapsedTime
    'Go through the particle loop
    For I = 0 To EditorEffectData.ParticleCount
        'Check If Particle Is In Use
        If EditorEffectData.Particles(I).Used Then
            'Update The Particle
            EditorEffectData.Particles(I).UpdateParticle ElapsedTime
            'Check if the particle is ready to die
            If EditorEffectData.Particles(I).sngA <= 0 Then
                'Check if the effect is ending
                If EditorEffectData.Progression > 0 Then
                    'Reset the particle
                    GDIResetEffect I
                Else
                     'Disable the particle
                    EditorEffectData.Particles(I).Used = False
                    'Subtract from the total particle count
                    EditorEffectData.ParticlesLeft = EditorEffectData.ParticlesLeft - 1
                    'Check if the effect is out of particles
                    If EditorEffectData.ParticlesLeft = 0 Then GDIBeginEffect
                    'Clear the color (dont leave behind any artifacts)
                    EditorEffectData.PartVertex(I).color = 0
                End If
            Else
                'Set the particle information on the particle vertex
                EditorEffectData.PartVertex(I).color = D3DColorMake(EditorEffectData.Particles(I).sngR, EditorEffectData.Particles(I).sngG, EditorEffectData.Particles(I).sngB, EditorEffectData.Particles(I).sngA)
                EditorEffectData.PartVertex(I).X = EditorEffectData.Particles(I).sngX
                EditorEffectData.PartVertex(I).Y = EditorEffectData.Particles(I).sngY
            End If
        End If
    Next I
End Sub

Private Sub GDIResetEffect(ByVal Index As Long)
Dim A As Single
Dim X As Single
Dim Y As Single

    Select Case EditorEffectData.Effectnum
        Case EFFECT_TYPE_HEAL
            'Reset the particle
            X = EditorEffectData.X - 10 + Rnd * 20
            Y = EditorEffectData.Y - 10 + Rnd * 20
            EditorEffectData.Particles(Index).ResetIt X, Y, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), EditorEffectData.XAcc, EditorEffectData.YAcc
            EditorEffectData.Particles(Index).ResetColor EditorEffectData.Red / 100, EditorEffectData.Green / 100, EditorEffectData.Blue / 100, EditorEffectData.Alpha / 100 + (Rnd * 0.2), EditorEffectData.Decay / 100 + (Rnd * 0.5)
        Case EFFECT_TYPE_PROTECTION
            'Get the positions
            A = Rnd * 360 * DegreeToRadian
            X = EditorEffectData.X - (Sin(A) * EditorEffectData.Modifier)
            Y = EditorEffectData.Y + (Cos(A) * EditorEffectData.Modifier)
            'Reset the particle
            EditorEffectData.Particles(Index).ResetIt X, Y, EditorEffectData.XSpeed, Rnd * EditorEffectData.YSpeed, EditorEffectData.XAcc, EditorEffectData.YAcc
            EditorEffectData.Particles(Index).ResetColor EditorEffectData.Red / 100, EditorEffectData.Green / 100, EditorEffectData.Blue / 100, EditorEffectData.Alpha / 100 + (Rnd * 0.4), EditorEffectData.Decay / 100 + (Rnd * 0.2)
        Case EFFECT_TYPE_STRENGTHEN
            'Get the positions
            A = Rnd * 360 * DegreeToRadian
            X = EditorEffectData.X - (Sin(A) * EditorEffectData.Modifier)
            Y = EditorEffectData.Y + (Cos(A) * EditorEffectData.Modifier)
            'Reset the particle
            EditorEffectData.Particles(Index).ResetIt X, Y, EditorEffectData.XSpeed, Rnd * EditorEffectData.YSpeed, EditorEffectData.XAcc, EditorEffectData.YAcc
            EditorEffectData.Particles(Index).ResetColor EditorEffectData.Red / 100, EditorEffectData.Green / 100, EditorEffectData.Blue / 100, EditorEffectData.Alpha / 100 + (Rnd * 0.4), EditorEffectData.Decay / 100 + (Rnd * 0.2)
        Case EFFECT_TYPE_SUMMON
            A = (Index / 30) * EXP(Index / (EditorEffectData.Progression + 10 + (EditorEffectData.Modifier * 10)))
            X = EditorEffectData.X + (A * Cos(Index))
            Y = EditorEffectData.Y + (A * Sin(Index))
            'Reset the particle
            EditorEffectData.Particles(Index).ResetIt X, Y, EditorEffectData.XSpeed, EditorEffectData.YSpeed, EditorEffectData.XAcc, EditorEffectData.YAcc
            EditorEffectData.Particles(Index).ResetColor EditorEffectData.Red / 100, EditorEffectData.Green / 100 + (Rnd * 0.2), EditorEffectData.Blue / 100, EditorEffectData.Alpha / 100, EditorEffectData.Decay / 100 + (Rnd * 0.2)
    End Select
End Sub

Public Function Effect_FToDW(F As Single) As Long
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = Direct3DX8.CreateBuffer(4)
    Direct3DX8.BufferSetData Buf, 0, 4, 1, F
    Direct3DX8.BufferGetData Buf, 0, 4, 1, Effect_FToDW
End Function
