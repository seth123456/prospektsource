Attribute VB_Name = "modParticles"
Option Explicit

'Constants With The Order Number For Each Effect
Public Const EFFECT_TYPE_HEAL As Byte = 1
Public Const EFFECT_TYPE_PROTECTION As Byte = 2
Public Const EFFECT_TYPE_STRENGTHEN As Byte = 3
Public Const EFFECT_TYPE_SUMMON As Byte = 4

Private Sub UpdateEffectBinding(ByVal EffectIndex As Integer)
'Updates the binding of a particle effect to a target, if the effect is bound to a character
Dim TargetA As Single
 
    'Update position through character binding
    If EffectData(EffectIndex).BindIndex > 0 Then
        'Calculate the X and Y positions
        Select Case EffectData(EffectIndex).BindType
            Case TARGET_TYPE_PLAYER
                EffectData(EffectIndex).GoToX = ConvertMapX((Player(EffectData(EffectIndex).BindIndex).X * 32) + Player(EffectData(EffectIndex).BindIndex).xOffset + Half_PIC_X)
                EffectData(EffectIndex).GoToY = ConvertMapY((Player(EffectData(EffectIndex).BindIndex).Y * 32) + Player(EffectData(EffectIndex).BindIndex).yOffset + Half_PIC_Y)
            Case TARGET_TYPE_NPC
                EffectData(EffectIndex).GoToX = ConvertMapX((MapNpc(EffectData(EffectIndex).BindIndex).X * 32 + MapNpc(EffectData(EffectIndex).BindIndex).xOffset) + Half_PIC_X)
                EffectData(EffectIndex).GoToY = ConvertMapY((MapNpc(EffectData(EffectIndex).BindIndex).Y * 32 + MapNpc(EffectData(EffectIndex).BindIndex).yOffset) + Half_PIC_Y)
        End Select
    End If
    'Move to the new position if needed
    If EffectData(EffectIndex).GoToX <> EffectData(EffectIndex).X Or EffectData(EffectIndex).GoToY <> EffectData(EffectIndex).Y Then
        If EffectData(EffectIndex).X > 0 And EffectData(EffectIndex).X < (MAX_MAPX * 32) Then
        If EffectData(EffectIndex).Y > 0 And EffectData(EffectIndex).Y < (MAX_MAPY * 32) Then
            'Calculate the angle
            TargetA = Engine_GetAngle((EffectData(EffectIndex).X), (EffectData(EffectIndex).Y), EffectData(EffectIndex).GoToX, EffectData(EffectIndex).GoToY) - 180
            'Update the position of the effect
            EffectData(EffectIndex).X = EffectData(EffectIndex).X - Sin(TargetA * DegreeToRadian) * EffectData(EffectIndex).BindSpeed
            EffectData(EffectIndex).Y = EffectData(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * EffectData(EffectIndex).BindSpeed
 
            'Check if the effect is close enough to the target to just stick it at the target
            If Abs(EffectData(EffectIndex).X - EffectData(EffectIndex).GoToX) < 6 Then EffectData(EffectIndex).X = EffectData(EffectIndex).GoToX
            If Abs(EffectData(EffectIndex).Y - EffectData(EffectIndex).GoToY) < 6 Then EffectData(EffectIndex).Y = EffectData(EffectIndex).GoToY
            
            'Check if the position of the effect is equal to that of the target
            If EffectData(EffectIndex).X = EffectData(EffectIndex).GoToX Then
                If EffectData(EffectIndex).Y = EffectData(EffectIndex).GoToY Then
                    'For some effects, if the position is reached, we want to end the effect
                    If EffectData(EffectIndex).KillWhenAtTarget Then
                        EffectData(EffectIndex).BindIndex = 0
                        EffectData(EffectIndex).Progression = 0
                        EffectData(EffectIndex).GoToX = EffectData(EffectIndex).X
                        EffectData(EffectIndex).GoToY = EffectData(EffectIndex).Y
                    End If
                    Exit Sub    'The effect is at the right position, don't update
                End If
            End If
        End If
        End If
    End If
End Sub

Public Function Effect_FToDW(F As Single) As Long
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = Direct3DX8.CreateBuffer(4)
    Direct3DX8.BufferSetData Buf, 0, 4, 1, F
    Direct3DX8.BufferGetData Buf, 0, 4, 1, Effect_FToDW
End Function

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'Kills (stops) a single effect or all effects
Dim I As Long

    'Check If To Kill All Effects
    If KillAll = True Then
        'Loop Through Every Effect
        For I = 1 To MAX_BYTE
            'Stop The Effect
            EffectData(I).Used = False
        Next
    Else
        'Stop The Selected Effect
        EffectData(EffectIndex).Used = False
    End If
End Sub

Private Function Effect_NextOpenSlot() As Integer
'Finds the next open effects index
Dim EffectIndex As Integer

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1 'Check The Next Slot
        If EffectIndex > MAX_BYTE Then 'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While EffectData(EffectIndex).Used = True 'Check Next If Effect Is In Use
    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex
    'Clear the old information from the effect
    Erase EffectData(EffectIndex).Particles()
    Erase EffectData(EffectIndex).PartVertex()
    ZeroMemory EffectData(EffectIndex), LenB(EffectData(EffectIndex))
    EffectData(EffectIndex).GoToX = -30000
    EffectData(EffectIndex).GoToY = -30000
End Function

Public Sub RenderEffectData(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
    
    'Check if we have the device
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    If EffectData(EffectIndex).Gfx <= 0 Or EffectData(EffectIndex).Gfx > Count_Particle Then Exit Sub

    'Set the render state for the size of the particle
    Call D3DDevice8.SetRenderState(D3DRS_POINTSIZE, EffectData(EffectIndex).FloatSize)
    
    'Set the render state to point blitting
    If SetRenderStates Then D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the texture
    Directx8.SetTexture Tex_Particle(EffectData(EffectIndex).Gfx)
    D3DDevice8.SetTexture 0, gTexture(Tex_Particle(EffectData(EffectIndex).Gfx)).Texture

    'Draw all the particles at once
    D3DDevice8.DrawPrimitiveUP D3DPT_POINTLIST, EffectData(EffectIndex).ParticleCount, EffectData(EffectIndex).PartVertex(0), Len(EffectData(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Sub UpdateEffectAll(ByVal I As Long)
    'Make sure the effect is in use
    If EffectData(I).Used Then
        'Update the effect position if it is binded
        Call UpdateEffectBinding(I)
        If EffectData(I).Progression = 0 Then Effect_Kill I
        'Find out which effect is selected, then update it
        UpdateEffect I, EffectData(I).EffectNum
        'Render the effect
        Call RenderEffectData(I)
    End If
End Sub

Public Sub BeginEffect(ByVal EffectNum As Long, ByVal EffectType As Long, ByVal X As Single, ByVal Y As Single, ByVal LockType As Byte, ByVal LockIndex As Long)
Dim EffectIndex As Integer
Dim I As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Sub

    'Set The Effect's Variables
    EffectData(EffectIndex).EffectNum = EffectType 'Set the effect number
    EffectData(EffectIndex).Used = True 'Enabled the effect
    EffectData(EffectIndex).X = X 'Set the effect's X coordinate
    EffectData(EffectIndex).Y = Y 'Set the effect's Y coordinate
    EffectData(EffectIndex).BindType = LockType
    EffectData(EffectIndex).BindIndex = LockIndex
    EffectData(EffectIndex).ParticleCount = Effect(EffectNum).Particles 'Set the number of particles
    EffectData(EffectIndex).Gfx = Effect(EffectNum).Sprite 'Set the graphic
    EffectData(EffectIndex).Alpha = Effect(EffectNum).Alpha
    EffectData(EffectIndex).Decay = Effect(EffectNum).Decay
    EffectData(EffectIndex).Red = Effect(EffectNum).Red
    EffectData(EffectIndex).Green = Effect(EffectNum).Green
    EffectData(EffectIndex).Blue = Effect(EffectNum).Blue
    EffectData(EffectIndex).XSpeed = Effect(EffectNum).XSpeed
    EffectData(EffectIndex).YSpeed = Effect(EffectNum).YSpeed
    EffectData(EffectIndex).XAcc = Effect(EffectNum).XAcc
    EffectData(EffectIndex).YAcc = Effect(EffectNum).YAcc
    EffectData(EffectIndex).Modifier = Effect(EffectNum).Modifier 'How large the circle is
    EffectData(EffectIndex).Progression = Effect(EffectNum).Duration 'How long the effect will last
    EffectData(EffectIndex).FloatSize = Effect_FToDW(Effect(EffectNum).Size) 'Size of the particles
    'Set the number of particles left to the total avaliable
    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticleCount
    'Redim the number of particles
    ReDim EffectData(EffectIndex).Particles(0 To EffectData(EffectIndex).ParticleCount)
    ReDim EffectData(EffectIndex).PartVertex(0 To EffectData(EffectIndex).ParticleCount)
    'Create the particles
    For I = 0 To EffectData(EffectIndex).ParticleCount
        Set EffectData(EffectIndex).Particles(I) = New clsParticle
        EffectData(EffectIndex).Particles(I).Used = True
        EffectData(EffectIndex).PartVertex(I).RHW = 1
        ResetEffect EffectIndex, I, EffectType
    Next I
    'Set The Initial Time
    EffectData(EffectIndex).PreviousFrame = timeGetTime
End Sub

Private Sub UpdateEffect(ByVal EffectIndex As Long, ByVal EffectNum As Long)
Dim ElapsedTime As Single
Dim I As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - EffectData(EffectIndex).PreviousFrame) * 0.01
    EffectData(EffectIndex).PreviousFrame = timeGetTime
    If EffectData(EffectIndex).Progression > 0 Then EffectData(EffectIndex).Progression = EffectData(EffectIndex).Progression - ElapsedTime
    'Go through the particle loop
    For I = 0 To EffectData(EffectIndex).ParticleCount
        'Check If Particle Is In Use
        If EffectData(EffectIndex).Particles(I).Used Then
            'Update The Particle
            EffectData(EffectIndex).Particles(I).UpdateParticle ElapsedTime
            'Check if the particle is ready to die
            If EffectData(EffectIndex).Particles(I).sngA <= 0 Then
                'Check if the effect is ending
                If EffectData(EffectIndex).Progression > 0 Then
                    'Reset the particle
                    ResetEffect EffectIndex, I, EffectNum
                Else
                    'Disable the particle
                    EffectData(EffectIndex).Particles(I).Used = False
                    'Subtract from the total particle count
                    EffectData(EffectIndex).ParticlesLeft = EffectData(EffectIndex).ParticlesLeft - 1
                    'Check if the effect is out of particles
                    If EffectData(EffectIndex).ParticlesLeft = 0 Then EffectData(EffectIndex).Used = False
                    'Clear the color (dont leave behind any artifacts)
                    EffectData(EffectIndex).PartVertex(I).color = 0
                End If
            Else
                'Set the particle information on the particle vertex
                EffectData(EffectIndex).PartVertex(I).color = D3DColorMake(EffectData(EffectIndex).Particles(I).sngR, EffectData(EffectIndex).Particles(I).sngG, EffectData(EffectIndex).Particles(I).sngB, EffectData(EffectIndex).Particles(I).sngA)
                EffectData(EffectIndex).PartVertex(I).X = EffectData(EffectIndex).Particles(I).sngX
                EffectData(EffectIndex).PartVertex(I).Y = EffectData(EffectIndex).Particles(I).sngY
            End If
        End If
    Next I
End Sub

Private Sub ResetEffect(ByVal EffectIndex As Long, ByVal Index As Long, ByVal EffectNum As Long)
Dim A As Single
Dim X As Single
Dim Y As Single

    Select Case EffectNum
        Case EFFECT_TYPE_HEAL
            'Reset the particle
            X = EffectData(EffectIndex).X - 10 + Rnd * 20
            Y = EffectData(EffectIndex).Y - 10 + Rnd * 20
            EffectData(EffectIndex).Particles(Index).ResetIt X, Y, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
            EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.2), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.5)
        Case EFFECT_TYPE_PROTECTION
            'Get the positions
            A = Rnd * 360 * DegreeToRadian
            X = EffectData(EffectIndex).X - (Sin(A) * EffectData(EffectIndex).Modifier)
            Y = EffectData(EffectIndex).Y + (Cos(A) * EffectData(EffectIndex).Modifier)
            'Reset the particle
            EffectData(EffectIndex).Particles(Index).ResetIt X, Y, EffectData(EffectIndex).XSpeed, Rnd * EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
            EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.4), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)
        Case EFFECT_TYPE_STRENGTHEN
            'Get the positions
            A = Rnd * 360 * DegreeToRadian
            X = EffectData(EffectIndex).X - (Sin(A) * EffectData(EffectIndex).Modifier)
            Y = EffectData(EffectIndex).Y + (Cos(A) * EffectData(EffectIndex).Modifier)
            'Reset the particle
            EffectData(EffectIndex).Particles(Index).ResetIt X, Y, EffectData(EffectIndex).XSpeed, Rnd * EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
            EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100, EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100 + (Rnd * 0.4), EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)
        Case EFFECT_TYPE_SUMMON
            A = (Index / 30) * EXP(Index / (EffectData(EffectIndex).Progression + 10 + (EffectData(EffectIndex).Modifier * 10)))
            X = EffectData(EffectIndex).X + (A * Cos(Index))
            Y = EffectData(EffectIndex).Y + (A * Sin(Index))
            'Reset the particle
            EffectData(EffectIndex).Particles(Index).ResetIt X, Y, EffectData(EffectIndex).XSpeed, EffectData(EffectIndex).YSpeed, EffectData(EffectIndex).XAcc, EffectData(EffectIndex).YAcc
            EffectData(EffectIndex).Particles(Index).ResetColor EffectData(EffectIndex).Red / 100, EffectData(EffectIndex).Green / 100 + (Rnd * 0.2), EffectData(EffectIndex).Blue / 100, EffectData(EffectIndex).Alpha / 100, EffectData(EffectIndex).Decay / 100 + (Rnd * 0.2)
    End Select
End Sub
