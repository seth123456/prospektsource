Attribute VB_Name = "modRendering"
Option Explicit

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    RHW As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public Type TextureRec
    Texture As Direct3DTexture8
    Width As Long
    Height As Long
    RWidth As Long
    RHeight As Long
    ImageData() As Byte
    path As String
    UnloadTimer As Long
    loaded As Boolean
End Type

Public gTexture() As TextureRec

' ****** PI ******
Public Const DegreeToRadian As Single = 0.0174532919296  'Pi / 180
Public Const RadianToDegree As Single = 57.2958300962816 '180 / Pi

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

Public Type GeomRec
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type

Public CurrentTexture As Long

Public ScreenWidth As Long
Public ScreenHeight As Long


' Texture arrays holders
Public Tex_Anim() As Long
Public Tex_Char() As Long
Public Tex_Face() As Long
Public Tex_Item() As Long
Public Tex_Paperdoll() As Long
Public Tex_Resource() As Long
Public Tex_Spellicon() As Long
Public Tex_Tileset() As Long
Public Tex_Projectile() As Long
Public Tex_Event() As Long
Public Tex_Particle() As Long
Public Tex_Editor() As Long
Public Tex_Fog() As Long

' Single texture holders
Public Tex_Direction As Long
Public Tex_Selection As Long

' Texture count
Public Count_Anim As Long
Public Count_Char As Long
Public Count_Face As Long
Public Count_Item As Long
Public Count_Paperdoll As Long
Public Count_Resource As Long
Public Count_Spellicon As Long
Public Count_Tileset As Long
Public Count_Projectile As Long
Public Count_Event As Long
Public Count_Particle As Long
Public Count_Editor As Long
Public Count_Fog As Long

' Texture paths
Public Const Path_Anim As String = "\data files\graphics\animations\"
Public Const Path_Char As String = "\data files\graphics\characters\"
Public Const Path_Face As String = "\data files\graphics\faces\"
Public Const Path_Item As String = "\data files\graphics\items\"
Public Const Path_Paperdoll As String = "\data files\graphics\paperdolls\"
Public Const Path_Resource As String = "\data files\graphics\resources\"
Public Const Path_Spellicon As String = "\data files\graphics\spellicons\"
Public Const Path_Tileset As String = "\data files\graphics\tilesets\"
Public Const Path_Font As String = "\data files\graphics\fonts\"
Public Const Path_Graphics As String = "\data files\graphics\"
Public Const Path_Projectile As String = "\data files\graphics\projectiles\"
Public Const Path_Event As String = "\data files\graphics\events\"
Public Const Path_Particle As String = "\data files\graphics\particles\"
Public Const Path_Editor As String = "\data files\graphics\editors\"
Public Const Path_Fog As String = "\data files\graphics\editors\"

Public Sub CacheTextures()

    ' Animation Textures
    Count_Anim = 1
    Do While FileExist(App.path & Path_Anim & Count_Anim & GFX_EXT)
        ReDim Preserve Tex_Anim(0 To Count_Anim)
        Tex_Anim(Count_Anim) = Directx8.SetTexturePath(App.path & Path_Anim & Count_Anim & GFX_EXT)
        Count_Anim = Count_Anim + 1
    Loop
    Count_Anim = Count_Anim - 1
    
    ' Character Textures
    Count_Char = 1
    Do While FileExist(App.path & Path_Char & Count_Char & GFX_EXT)
        ReDim Preserve Tex_Char(0 To Count_Char)
        Tex_Char(Count_Char) = Directx8.SetTexturePath(App.path & Path_Char & Count_Char & GFX_EXT)
        Count_Char = Count_Char + 1
    Loop
    Count_Char = Count_Char - 1
    
    ' Face Textures
    Count_Face = 1
    Do While FileExist(App.path & Path_Face & Count_Face & GFX_EXT)
        ReDim Preserve Tex_Face(0 To Count_Face)
        Tex_Face(Count_Face) = Directx8.SetTexturePath(App.path & Path_Face & Count_Face & GFX_EXT)
        Count_Face = Count_Face + 1
    Loop
    Count_Face = Count_Face - 1
    
    ' Item Textures
    Count_Item = 1
    Do While FileExist(App.path & Path_Item & Count_Item & GFX_EXT)
        ReDim Preserve Tex_Item(0 To Count_Item)
        Tex_Item(Count_Item) = Directx8.SetTexturePath(App.path & Path_Item & Count_Item & GFX_EXT)
        Count_Item = Count_Item + 1
    Loop
    Count_Item = Count_Item - 1

    ' Paperdoll Textures
    Count_Paperdoll = 1
    Do While FileExist(App.path & Path_Paperdoll & Count_Paperdoll & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(0 To Count_Paperdoll)
        Tex_Paperdoll(Count_Paperdoll) = Directx8.SetTexturePath(App.path & Path_Paperdoll & Count_Paperdoll & GFX_EXT)
        Count_Paperdoll = Count_Paperdoll + 1
    Loop
    Count_Paperdoll = Count_Paperdoll - 1

    ' Resource Textures
    Count_Resource = 1
    Do While FileExist(App.path & Path_Resource & Count_Resource & GFX_EXT)
        ReDim Preserve Tex_Resource(0 To Count_Resource)
        Tex_Resource(Count_Resource) = Directx8.SetTexturePath(App.path & Path_Resource & Count_Resource & GFX_EXT)
        Count_Resource = Count_Resource + 1
    Loop
    Count_Resource = Count_Resource - 1

    ' SpellIcon Textures
    Count_Spellicon = 1
    Do While FileExist(App.path & Path_Spellicon & Count_Spellicon & GFX_EXT)
        ReDim Preserve Tex_Spellicon(0 To Count_Spellicon)
        Tex_Spellicon(Count_Spellicon) = Directx8.SetTexturePath(App.path & Path_Spellicon & Count_Spellicon & GFX_EXT)
        Count_Spellicon = Count_Spellicon + 1
    Loop
    Count_Spellicon = Count_Spellicon - 1
    
    ' Projectile Textures
    Count_Projectile = 1
    Do While FileExist(App.path & Path_Projectile & Count_Projectile & GFX_EXT)
        ReDim Preserve Tex_Projectile(0 To Count_Projectile)
        Tex_Projectile(Count_Projectile) = Directx8.SetTexturePath(App.path & Path_Projectile & Count_Projectile & GFX_EXT)
        Count_Projectile = Count_Projectile + 1
    Loop
    Count_Projectile = Count_Projectile - 1

    ' Tileset Textures
    Count_Tileset = 1
    Do While FileExist(App.path & Path_Tileset & Count_Tileset & GFX_EXT)
        ReDim Preserve Tex_Tileset(0 To Count_Tileset)
        Tex_Tileset(Count_Tileset) = Directx8.SetTexturePath(App.path & Path_Tileset & Count_Tileset & GFX_EXT)
        Count_Tileset = Count_Tileset + 1
    Loop
    Count_Tileset = Count_Tileset - 1
    
    ' Fog Textures
    Count_Fog = 1
    Do While FileExist(App.path & Path_Fog & Count_Fog & GFX_EXT)
        ReDim Preserve Tex_Fog(0 To Count_Fog)
        Tex_Fog(Count_Fog) = Directx8.SetTexturePath(App.path & Path_Fog & Count_Fog & GFX_EXT)
        Count_Fog = Count_Fog + 1
    Loop
    Count_Fog = Count_Fog - 1
    
    ' event Textures
    Count_Event = 1
    Do While FileExist(App.path & Path_Event & Count_Event & GFX_EXT)
        ReDim Preserve Tex_Event(0 To Count_Event)
        Tex_Event(Count_Event) = Directx8.SetTexturePath(App.path & Path_Event & Count_Event & GFX_EXT)
        Count_Event = Count_Event + 1
    Loop
    Count_Event = Count_Event - 1
    
    ' Particle Textures
    Count_Particle = 1
    Do While FileExist(App.path & Path_Particle & Count_Particle & GFX_EXT)
        ReDim Preserve Tex_Particle(0 To Count_Particle)
        Tex_Particle(Count_Particle) = Directx8.SetTexturePath(App.path & Path_Particle & Count_Particle & GFX_EXT)
        Count_Particle = Count_Particle + 1
    Loop
    Count_Particle = Count_Particle - 1
    
    ' Editor Textures
    Count_Editor = 1
    Do While FileExist(App.path & Path_Editor & Count_Editor & GFX_EXT)
        ReDim Preserve Tex_Editor(0 To Count_Editor)
        Tex_Editor(Count_Editor) = Directx8.SetTexturePath(App.path & Path_Editor & Count_Editor & GFX_EXT)
        Count_Editor = Count_Editor + 1
    Loop
    Count_Editor = Count_Editor - 1
    
    Tex_Direction = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\direction" & GFX_EXT)
    Tex_Selection = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\select" & GFX_EXT)
End Sub

Public Sub Render_Graphics()
Dim X As Long, Y As Long, I As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler
    
    'Check for device lost.
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then Directx8.DeviceLost: Exit Sub
    
    Directx8.UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' render lower tiles
    If Count_Tileset > 0 Then
        If HasMap Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        For I = MapLayer.Ground To MapLayer.Mask2
                            Call DrawMapTile(X, Y, I)
                            Call DrawEvent(X, Y, I)
                        Next I
                        Call DrawItem(X, Y)
                        Call DrawNpc(X, Y)
                    End If
                Next Y
            Next X
        End If
    End If
    
    ' Y-based render. Renders Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        ' Resources
        If Count_Resource > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For I = 1 To Resource_Index
                        If MapResource(I).Y = Y Then
                            Call DrawResource(I)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' render lower tiles
    If Count_Tileset > 0 Then
        If HasMap Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        For I = MapLayer.Fringe To MapLayer.Fringe2
                            Call DrawMapTile(X, Y, I)
                            Call DrawEvent(X, Y, I)
                        Next I
                        If frmMain.chkRoof Then
                            Call DrawMapTile(X, Y, Roof)
                            Call DrawEvent(X, Y, Roof)
                        End If
                        Call DrawGrid(X, Y)
                        If CurEditType = EDIT_ATTRIBUTES Then
                            Call DrawAttribute(X, Y)
                        ElseIf CurEditType = EDIT_DIRBLOCK Then
                            Call DrawDirection(X, Y)
                        End If
                    End If
                Next Y
            Next X
        End If
    End If
    
    If CurEditType = EDIT_MAP Then DrawTileOutLine

    ' End the rendering
    Call D3DDevice8.EndScene
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Directx8.DeviceLost
        Exit Sub
    Else
        Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
        ' GDI Rendering
        DrawGDI
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawGDI()
    If frmMain.Visible Then
        GDIRenderTileset
    ElseIf frmEditor_Item.Visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.Value
        GDIRenderPaperdoll frmEditor_Item.picPaperdoll, frmEditor_Item.scrlPaperdoll.Value
    ElseIf frmEditor_Animation.Visible Then
        GDIRenderAnimation
    ElseIf frmEditor_NPC.Visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.Value
    ElseIf frmEditor_Spell.Visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.Value
    ElseIf frmEditor_Resource.Visible Then
        GDIRenderResource
    ElseIf frmEditor_Effect.Visible Then
        GDIUpdateEffectAll
    ElseIf frmEditor_Events.Visible Then
        GDIRenderEvent
    End If
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim Height As Long, Width As Long, sRect As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Item Then Exit Sub
    
    Height = gTexture(Tex_Item(Sprite)).RHeight
    Width = gTexture(Tex_Item(Sprite)).RWidth
    
    sRect.Top = 0
    sRect.bottom = CELL_SIZE
    sRect.Left = 0
    sRect.Right = CELL_SIZE

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Item(Sprite), 0, 0, 0, 0, CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim Height As Long, Width As Long, sRect As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Spellicon Then Exit Sub
    
    Height = gTexture(Tex_Spellicon(Sprite)).Height
    Width = gTexture(Tex_Spellicon(Sprite)).Width
    
    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If
    
    sRect.Top = 0
    sRect.bottom = Height
    sRect.Left = 0
    sRect.Right = Width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Spellicon(Sprite), 0, 0, 0, 0, CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

' Paperdoll show up in item editor
Public Sub GDIRenderPaperdoll(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim Height As Long, Width As Long, sRect As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Paperdoll Then Exit Sub
    
    Height = CELL_SIZE
    Width = CELL_SIZE
    
    sRect.Top = 0
    sRect.bottom = sRect.Top + Height
    sRect.Left = 0
    sRect.Right = sRect.Left + Width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Paperdoll(Sprite), 0, 0, 0, 0, Width, Height, Width, Height
     
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal Sprite As Long)
Dim Height As Long, Width As Long, sRect As RECT

    ' exit out if doesn't exist
    If Sprite <= 0 Or Sprite > Count_Char Then Exit Sub
    
    Height = CELL_SIZE
    Width = CELL_SIZE
    
    sRect.Top = 0
    sRect.bottom = sRect.Top + Height
    sRect.Left = 0
    sRect.Right = sRect.Left + Width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    Directx8.RenderTexture Tex_Char(Sprite), 0, 0, 0, 0, Width, Height, Width, Height
     
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
Dim Height As Long, Width As Long, Tileset As Byte, sRect As RECT

    ' find tileset number
    Tileset = frmMain.scrlTileset.Value
    
    ' exit out if doesn't exist
    If Tileset <= 0 Or Tileset > Count_Tileset Then Exit Sub
    
    Height = gTexture(Tex_Tileset(Tileset)).Height
    Width = gTexture(Tex_Tileset(Tileset)).Width
    
    If Height = 0 Or Width = 0 Then
        Height = 1
        Width = 1
    End If

    sRect.Top = 0
    sRect.bottom = frmMain.picTileset.Height
    sRect.Left = 0
    sRect.Right = frmMain.picTileset.Width
    
    ' change selected shape for autotiles
    If frmMain.scrlAutotile.Value > 0 Then
        Select Case frmMain.scrlAutotile.Value
            Case 1 ' autotile
                shpSelectedWidth = 2 * CELL_SIZE
                shpSelectedHeight = 3 * CELL_SIZE
            Case 2 ' fake autotile
                shpSelectedWidth = 1 * CELL_SIZE
                shpSelectedHeight = 1 * CELL_SIZE
            Case 3 ' animated
                shpSelectedWidth = 6 * CELL_SIZE
                shpSelectedHeight = 3 * CELL_SIZE
            Case 4 ' cliff
                shpSelectedWidth = 2 * CELL_SIZE
                shpSelectedHeight = 2 * CELL_SIZE
            Case 5 ' waterfall
                shpSelectedWidth = 2 * CELL_SIZE
                shpSelectedHeight = 3 * CELL_SIZE
        End Select
    End If

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If Tex_Tileset(Tileset) <= 0 Then Exit Sub
    Directx8.RenderTexture Tex_Tileset(Tileset), 0, 0, slcTilesetLeft, slcTilesetTop, gTexture(Tex_Tileset(Tileset)).RWidth, gTexture(Tex_Tileset(Tileset)).RHeight, gTexture(Tex_Tileset(Tileset)).RWidth, gTexture(Tex_Tileset(Tileset)).RHeight
    DrawSelectionBox shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRect, ByVal 0, frmMain.picTileset.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderResource()
Dim Sprite As Long
Dim sRect As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > Count_Resource Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRect.Top = 0
        sRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        sRect.Left = 0
        sRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        dRect.Top = 0
        dRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        dRect.Left = 0
        dRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        D3DDevice8.BeginScene
        Directx8.RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        With srcRect
            .X1 = 0
            .X2 = gTexture(Tex_Resource(Sprite)).RWidth
            .Y1 = 0
            .Y2 = gTexture(Tex_Resource(Sprite)).RHeight
        End With
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        D3DDevice8.EndScene
        D3DDevice8.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > Count_Resource Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRect.Top = 0
        sRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        sRect.Left = 0
        sRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        dRect.Top = 0
        dRect.bottom = gTexture(Tex_Resource(Sprite)).RHeight
        dRect.Left = 0
        dRect.Right = gTexture(Tex_Resource(Sprite)).RWidth
        D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        D3DDevice8.BeginScene
        Directx8.RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        
        With destRect
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = gTexture(Tex_Resource(Sprite)).RWidth
            .Y1 = 0
            .Y2 = gTexture(Tex_Resource(Sprite)).RHeight
        End With
                    
        D3DDevice8.EndScene
        D3DDevice8.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub GDIRenderAnimation()
Dim I As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim sX As Long, sY As Long, sRect As RECT
    
    sRect.Top = 0
    sRect.bottom = 192
    sRect.Left = 0
    sRect.Right = 192

    For I = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(I).Value
        
        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(I)
            FrameCount = frmEditor_Animation.scrlFrameCount(I)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(I) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimEditorFrame(I) >= FrameCount Then
                    AnimEditorFrame(I) = 1
                Else
                    AnimEditorFrame(I) = AnimEditorFrame(I) + 1
                End If
                AnimEditorTimer(I) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(I).Value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    Height = 192

                    sY = (Height * ((AnimEditorFrame(I) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(I) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice8.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    Directx8.RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    
                    ' Finish Rendering
                    Call D3DDevice8.EndScene
                    Call D3DDevice8.Present(sRect, ByVal 0, frmEditor_Animation.picSprite(I).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
End Sub

Public Sub GDIRenderEvent()
Dim eventNum As Long
Dim sRect As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler

    eventNum = frmEditor_Events.scrlGraphic.Value

    If eventNum < 1 Or eventNum > Count_Event Then
        frmEditor_Events.picGraphic.Cls
        Exit Sub
    End If


    ' rect for source
    sRect.Top = 0
    sRect.bottom = gTexture(Tex_Event(eventNum)).RHeight
    sRect.Left = 0
    sRect.Right = gTexture(Tex_Event(eventNum)).RWidth
    
    ' same for destination as source
    dRect = sRect
    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice8.BeginScene
    Directx8.RenderTextureByRects Tex_Event(eventNum), sRect, dRect
    
    With destRect
        .X1 = 0
        .X2 = gTexture(Tex_Event(eventNum)).RWidth
        .Y1 = 0
        .Y2 = gTexture(Tex_Event(eventNum)).RHeight
    End With
                    
    D3DDevice8.EndScene
    D3DDevice8.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GDIRenderEvent", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long, ByVal Layer As MapLayer)
    With Map.Tile(X, Y)
        ' skip tile if tileset isn't set
        If Autotile(X, Y).Layer(Layer).renderState = RENDER_STATE_NORMAL Then
            ' Draw normally
            Directx8.RenderTexture Tex_Tileset(.Layer(Layer).Tileset), ConvertMapX(X * CELL_SIZE), ConvertMapY(Y * CELL_SIZE), .Layer(Layer).X * CELL_SIZE, .Layer(Layer).Y * CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE
        ElseIf Autotile(X, Y).Layer(Layer).renderState = RENDER_STATE_AUTOTILE Then
            ' Draw autotiles
            DrawAutoTile Layer, X * CELL_SIZE, Y * CELL_SIZE, 1, X, Y
            DrawAutoTile Layer, (X * CELL_SIZE) + 16, Y * CELL_SIZE, 2, X, Y
            DrawAutoTile Layer, X * CELL_SIZE, (Y * CELL_SIZE) + 16, 3, X, Y
            DrawAutoTile Layer, (X * CELL_SIZE) + 16, (Y * CELL_SIZE) + 16, 4, X, Y
        End If
    End With
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = -CELL_SIZE
        Case AUTOTILE_ANIM
            xOffset = 0
        Case AUTOTILE_CLIFF
            yOffset = -CELL_SIZE
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    Directx8.RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), ConvertMapX(destX), ConvertMapY(destY), Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

Public Sub DrawAttribute(ByVal X As Long, ByVal Y As Long)
Dim I As Long
    With Map.Tile(X, Y)
        Directx8.RenderTexture Tex_Editor(.Type), ConvertMapX(X * CELL_SIZE), ConvertMapY(Y * CELL_SIZE), 0, 0, CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE
    End With
End Sub

Sub DrawSelectionBox(X As Long, Y As Long, Width As Long, Height As Long)
   On Error GoTo ErrorHandler

    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        Directx8.RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        Directx8.RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        Directx8.RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        Directx8.RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        Directx8.RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        Directx8.RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        Directx8.RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        Directx8.RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "DrawSelectionBox", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTileOutLine()
Dim Tileset As Byte

    ' find tileset number
   On Error GoTo ErrorHandler

    ' find tileset number
    Tileset = frmMain.scrlTileset.Value
    
    ' exit out if doesn't exist
    If Tileset <= 0 Or Tileset > Count_Tileset Then Exit Sub
    
    If frmMain.scrlAutotile.Value = 0 Then
        Directx8.RenderTexture Tex_Tileset(Tileset), ConvertMapX(CurX * CELL_SIZE), ConvertMapY(CurY * CELL_SIZE), EditorTileX * CELL_SIZE, EditorTileY * CELL_SIZE, shpSelectedWidth, shpSelectedHeight, shpSelectedWidth, shpSelectedHeight, D3DColorARGB(200, 255, 255, 255)
    Else
        Directx8.RenderTexture Tex_Tileset(Tileset), ConvertMapX(CurX * CELL_SIZE), ConvertMapY(CurY * CELL_SIZE), EditorTileX * CELL_SIZE, EditorTileY * CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE, CELL_SIZE, D3DColorARGB(200, 255, 255, 255)
    End If

   ' Error handler
   Exit Sub
ErrorHandler:
    HandleError "DrawTileOutLine", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub

End Sub

Public Sub DrawResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    X = MapResource(Resource_num).X
    Y = MapResource(Resource_num).Y
    
    If X < 0 Or X > Map.MaxX Then Exit Sub
    If Y < 0 Or Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(X, Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' pre-load texture for calculations
    Directx8.SetTexture Tex_Resource(Resource_sprite)

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Resource(Resource_sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Resource(Resource_sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * CELL_SIZE) - (gTexture(Tex_Resource(Resource_sprite)).RWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * CELL_SIZE) - gTexture(Tex_Resource(Resource_sprite)).RHeight + 32
    
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawEvent(ByVal X As Long, ByVal Y As Long, ByVal Layer As Byte)
Dim Index As Long
Dim Sprite As Long
Dim rec As RECT
Dim Width As Long, Height As Long
    
    If X < 0 Or X > Map.MaxX Then Exit Sub
    If Y < 0 Or Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    If Not Map.Tile(X, Y).Type = TILE_TYPE_EVENT Then Exit Sub
    Index = Map.Tile(X, Y).Data1
    If Index < 1 Or Index > MAX_EVENTS Then Exit Sub
    If Not Events(Index).Layer = Layer Then Exit Sub
    
    Sprite = Events(Index).Graphic(0)
    
    If Sprite <= 0 Or Sprite >= Count_Event Then Exit Sub

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Event(Sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Event(Sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (X * CELL_SIZE) - (gTexture(Tex_Event(Sprite)).RWidth / 2) + 16
    Y = (Y * CELL_SIZE) - gTexture(Tex_Event(Sprite)).RHeight + 32
    
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    Directx8.RenderTexture Tex_Event(Sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawGrid(ByVal X As Long, ByVal Y As Long)
Dim Top As Long, Left As Long
    ' render grid
    Top = 24
    Left = 0
    Directx8.RenderTexture Tex_Direction, ConvertMapX(X * CELL_SIZE), ConvertMapY(Y * CELL_SIZE), Left, Top, 32, 32, 32, 32
End Sub

Public Sub DrawNpc(ByVal X As Long, ByVal Y As Long)
    Dim Anim As Byte, Dir As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As GeomRec
    
    If Not Map.Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then Exit Sub ' no npc set
    If Map.Tile(X, Y).Data1 < 1 Or Map.Tile(X, Y).Data1 > MAX_MAP_NPCS Then Exit Sub
    If Map.Npc(Map.Tile(X, Y).Data1) < 1 Or Map.Npc(Map.Tile(X, Y).Data1) > MAX_NPCS Then Exit Sub
    
    ' pre-load texture for calculations
    Sprite = Npc(Map.Npc(Map.Tile(X, Y).Data1)).Sprite
    'SetTexture Tex_Char(Sprite)

    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub
    Dir = Map.Tile(X, Y).Data2
    
    Anim = 0
    ' Set the left
    Select Case Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .Top = (gTexture(Tex_Char(Sprite)).RHeight / 4) * spritetop
        .Height = gTexture(Tex_Char(Sprite)).RHeight / 4
        .Left = Anim * (gTexture(Tex_Char(Sprite)).RWidth / 4)
        .Width = (gTexture(Tex_Char(Sprite)).RWidth / 4)
    End With

    ' Calculate the X
    X = X * CELL_SIZE - ((gTexture(Tex_Char(Sprite)).RWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (gTexture(Tex_Char(Sprite)).RHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = Y * CELL_SIZE - ((gTexture(Tex_Char(Sprite)).RHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = Y * CELL_SIZE - 4
    End If
    
    Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawItem(ByVal X As Long, ByVal Y As Long)
    Dim Anim As Byte
    Dim Sprite As Long, spritetop As Long
    Dim rec As GeomRec
    
    If Not Map.Tile(X, Y).Type = TILE_TYPE_ITEM Then Exit Sub ' no npc set
    If Map.Tile(X, Y).Data1 < 1 Or Map.Tile(X, Y).Data1 > MAX_ITEMS Then Exit Sub
    ' pre-load texture for calculations
    Sprite = Item(Map.Tile(X, Y).Data1).Pic
    'SetTexture Tex_Char(Sprite)

    If Sprite < 1 Or Sprite > Count_Item Then Exit Sub

    With rec
        .Top = 0
        .Height = 32
        .Left = 0
        .Width = 32
    End With

    ' Calculate the X
    X = X * CELL_SIZE
    Y = Y * CELL_SIZE
    
    Directx8.RenderTexture Tex_Item(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
Dim I As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler
    
    ' render dir blobs
    For I = 1 To 4
        Left = (I - 1) * 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(I)) Then
            Top = 8
        Else
            Top = 16
        End If
        'render!
        Directx8.RenderTexture Tex_Direction, ConvertMapX(X * CELL_SIZE) + DirArrowX(I), ConvertMapY(Y * CELL_SIZE) + DirArrowY(I), Left, Top, 8, 8, 8, 8
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawDirection", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
