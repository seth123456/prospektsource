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
Public Tex_GUI() As Long
Public Tex_Item() As Long
Public Tex_Paperdoll() As Long
Public Tex_Resource() As Long
Public Tex_Spellicon() As Long
Public Tex_Tileset() As Long
Public Tex_Buttons() As Long
Public Tex_Buttons_h() As Long
Public Tex_Buttons_c() As Long
Public Tex_Surface() As Long
Public Tex_Fog() As Long
Public Tex_Projectile() As Long
Public Tex_Panorama() As Long
Public Tex_Class() As Long
Public Tex_Event() As Long
Public Tex_Particle() As Long
Public Tex_Cursor() As Long

' Single texture holders
Public Tex_Bars As Long
Public Tex_Blood As Long
Public Tex_Target As Long
Public Tex_White As Long
Public Tex_Weather As Long

' Texture count
Public Count_Anim As Long
Public Count_Char As Long
Public Count_Face As Long
Public Count_GUI As Long
Public Count_Item As Long
Public Count_Paperdoll As Long
Public Count_Resource As Long
Public Count_Spellicon As Long
Public Count_Tileset As Long
Public Count_Fog As Long
Public Count_Surface As Long
Public Count_Projectile As Long
Public Count_Panorama As Long
Public Count_Class As Long
Public Count_Event As Long
Public Count_Button As Long
Public Count_Particle As Long
Public Count_Cursor As Long

' Texture paths
Public Const Path_Anim As String = "\data files\graphics\animations\"
Public Const Path_Char As String = "\data files\graphics\characters\"
Public Const Path_Face As String = "\data files\graphics\faces\"
Public Const Path_GUI As String = "\data files\graphics\gui\"
Public Const Path_Item As String = "\data files\graphics\items\"
Public Const Path_Paperdoll As String = "\data files\graphics\paperdolls\"
Public Const Path_Resource As String = "\data files\graphics\resources\"
Public Const Path_Spellicon As String = "\data files\graphics\spellicons\"
Public Const Path_Tileset As String = "\data files\graphics\tilesets\"
Public Const Path_Font As String = "\data files\graphics\fonts\"
Public Const Path_Graphics As String = "\data files\graphics\"
Public Const Path_Buttons As String = "\data files\graphics\gui\buttons\"
Public Const Path_Surface As String = "\data files\graphics\surfaces\"
Public Const Path_Fog As String = "\data files\graphics\fog\"
Public Const Path_Projectile As String = "\data files\graphics\projectiles\"
Public Const Path_Panorama As String = "\data files\graphics\panoramas\"
Public Const Path_Class As String = "\data files\graphics\classes\"
Public Const Path_Event As String = "\data files\graphics\events\"
Public Const Path_Particle As String = "\data files\graphics\particles\"
Public Const Path_Cursor As String = "\data files\graphics\cursors\"

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
    
    ' GUI Textures
    Count_GUI = 1
    Do While FileExist(App.path & Path_GUI & Count_GUI & GFX_EXT)
        ReDim Preserve Tex_GUI(0 To Count_GUI)
        Tex_GUI(Count_GUI) = Directx8.SetTexturePath(App.path & Path_GUI & Count_GUI & GFX_EXT)
        Count_GUI = Count_GUI + 1
    Loop
    Count_GUI = Count_GUI - 1
    
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
    
    ' Button Textures
    Count_Button = 1
    Do While FileExist(App.path & Path_Buttons & Count_Button & GFX_EXT)
        ReDim Preserve Tex_Buttons(0 To Count_Button)
        ReDim Preserve Tex_Buttons_h(0 To Count_Button)
        ReDim Preserve Tex_Buttons_c(0 To Count_Button)
        Tex_Buttons(Count_Button) = Directx8.SetTexturePath(App.path & Path_Buttons & Count_Button & GFX_EXT)
        Tex_Buttons_h(Count_Button) = Directx8.SetTexturePath(App.path & Path_Buttons & Count_Button & "_h" & GFX_EXT)
        Tex_Buttons_c(Count_Button) = Directx8.SetTexturePath(App.path & Path_Buttons & Count_Button & "_c" & GFX_EXT)
        Count_Button = Count_Button + 1
    Loop
    Count_Button = Count_Button - 1
    
    ' Fog Textures
    Count_Fog = 1
    Do While FileExist(App.path & Path_Fog & Count_Fog & GFX_EXT)
        ReDim Preserve Tex_Fog(0 To Count_Fog)
        Tex_Fog(Count_Fog) = Directx8.SetTexturePath(App.path & Path_Fog & Count_Fog & GFX_EXT)
        Count_Fog = Count_Fog + 1
    Loop
    Count_Fog = Count_Fog - 1
    
    ' Surfaces
    Count_Surface = 1
    Do While FileExist(App.path & Path_Surface & Count_Surface & GFX_EXT)
        ReDim Preserve Tex_Surface(0 To Count_Surface)
        Tex_Surface(Count_Surface) = Directx8.SetTexturePath(App.path & Path_Surface & Count_Surface & GFX_EXT)
        Count_Surface = Count_Surface + 1
    Loop
    Count_Surface = Count_Surface - 1
    
    ' panoramas
    Count_Panorama = 1
    Do While FileExist(App.path & Path_Panorama & Count_Panorama & GFX_EXT)
        ReDim Preserve Tex_Panorama(0 To Count_Panorama)
        Tex_Panorama(Count_Panorama) = Directx8.SetTexturePath(App.path & Path_Panorama & Count_Panorama & GFX_EXT)
        Count_Panorama = Count_Panorama + 1
    Loop
    Count_Panorama = Count_Panorama - 1
    
    ' Classs
    Count_Class = 1
    Do While FileExist(App.path & Path_Class & Count_Class & GFX_EXT)
        ReDim Preserve Tex_Class(0 To Count_Class)
        Tex_Class(Count_Class) = Directx8.SetTexturePath(App.path & Path_Class & Count_Class & GFX_EXT)
        Count_Class = Count_Class + 1
    Loop
    Count_Class = Count_Class - 1
    
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
    
    ' Cursor Textures
    Count_Cursor = 1
    Do While FileExist(App.path & Path_Cursor & Count_Cursor & GFX_EXT)
        ReDim Preserve Tex_Cursor(0 To Count_Cursor)
        Tex_Cursor(Count_Cursor) = Directx8.SetTexturePath(App.path & Path_Cursor & Count_Cursor & GFX_EXT)
        Count_Cursor = Count_Cursor + 1
    Loop
    Count_Cursor = Count_Cursor - 1
    
    ' Single Textures
    Tex_Bars = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\bars" & GFX_EXT)
    Tex_Blood = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\blood" & GFX_EXT)
    Tex_Target = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\target" & GFX_EXT)
    Tex_White = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\white" & GFX_EXT)
    Tex_Weather = Directx8.SetTexturePath(App.path & Path_Graphics & "misc\weather" & GFX_EXT)
End Sub

'****************************************************
'                  Rendering loops
'****************************************************

Public Sub Render_Graphics()
Dim X As Long, Y As Long, I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'Check for device lost.
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then Directx8.DeviceLost: Exit Sub

    ' update the camera
    UpdateCamera
    
    Directx8.UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' If map is loading, display text and exit out
    If GettingMap Then
        RenderText Font_Default, "Receiving Map...", 350, 280, Blue
        ' End the rendering
        Call D3DDevice8.EndScene
        If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
            Directx8.DeviceLost
            Exit Sub
        Else
            Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
        End If
        Exit Sub
    End If
    
    ' draw panorama
    If Map.Panorama > 0 Then
        If Count_Panorama > 0 Then
            Directx8.RenderTexture Tex_Panorama(Map.Panorama), ParallaxX, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
            Directx8.RenderTexture Tex_Panorama(Map.Panorama), ParallaxX + ScreenWidth, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
        End If
    End If
    
    ' render lower tiles
    For X = TileView.Left To TileView.Right
        For Y = TileView.Top To TileView.bottom
            If IsValidMapPoint(X, Y) Then
                For I = MapLayer.Ground To MapLayer.Mask2
                    If Count_Tileset > 0 Then Call DrawMapTile(X, Y, I)
                    If Count_Event > 0 Then Call DrawEvent(X, Y, I)
                Next I
            End If
        Next Y
    Next X
    
    ' render the decals
    For I = 1 To MAX_BYTE
        Call DrawBlood(I)
    Next I
    
    ' render the items
    If Count_Item > 0 Then
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).Num > 0 Then
                Call DrawItem(I)
            End If
        Next I
    End If
    
    'Updates all of the effects and renders them
    If Count_Particle > 0 Then
        For I = 1 To MAX_BYTE
            UpdateEffectAll I
        Next I
    End If
    
    ' draw animations
    If Count_Anim > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(0) Then
                DrawAnimation I, 0
            End If
        Next I
    End If
    
    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        If Count_Char > 0 Then
            ' Players
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If Player(I).Y = Y Then
                        Call DrawPlayer(I)
                    End If
                End If
            Next
            
            ' Npcs
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Y = Y Then
                    Call DrawNpc(I)
                End If
            Next
        End If
        
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
    
    ' render projectiles
    If Count_Projectile > 0 Then Call DrawProjectile
    
    ' render out upper tiles
    For X = TileView.Left To TileView.Right
        For Y = TileView.Top To TileView.bottom
            If IsValidMapPoint(X, Y) Then
                For I = MapLayer.Fringe To MapLayer.Fringe2
                    If Count_Tileset > 0 Then Call DrawMapTile(X, Y, I)
                    If Count_Event > 0 Then Call DrawEvent(X, Y, I)
                Next I
                If Not Player(MyIndex).Threshold = YES Then
                    If Count_Tileset > 0 Then Call DrawMapTile(X, Y, Roof)
                    If Count_Event > 0 Then Call DrawEvent(X, Y, Roof)
                End If
            End If
        Next Y
    Next X
    
    ' render animations
    If Count_Anim > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(1) Then
                DrawAnimation I, 1
            End If
        Next
    End If
    
    ' blt the hover icon
    DrawTargetHover
    
    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).X * 32) + Player(myTarget).xOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset
        End If
    End If
    
    ' draw the bars
    DrawBars
    
    ' draw player names
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(I)
        End If
    Next
    
    ' draw npc names
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            Call DrawNpcName(I)
        End If
    Next
    
    ' draw action msg
    For I = 1 To MAX_BYTE
        DrawActionMsg I
    Next
    
    ' draw the messages
    For I = 1 To MAX_BYTE
        If chatBubble(I).active Then
            DrawChatBubble I
        End If
    Next
    
    ' render map effects
    DrawWeather
    If Count_Fog > 0 Then DrawFog
    DrawTint
    If FadeAmount > 0 Then Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorRGBA(0, 0, 0, FadeAmount)
    If FlashTimer > timeGetTime Then Directx8.RenderTexture Tex_White, 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32
    If DrawThunder > 0 Then Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
    
    If Not hideGUI Then DrawGUI
    If Count_Cursor > 0 Then DrawCursor
    
    ' Draw fade in
    If canFade Then DrawFader
    
    ' draw loc
    If BLoc Then
        RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 0, 0, Yellow
        RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 0, 16, Yellow
        RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 0, 32, Yellow
    End If
    
    ' End the rendering
    Call D3DDevice8.EndScene
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Directx8.DeviceLost
        Exit Sub
    Else
        Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Menu()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'Check for device lost.
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then Directx8.DeviceLost: Exit Sub
    
    Directx8.UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' fader
    Select Case faderState
        Case 0, 1
            ' render background
            If Not faderAlpha = 255 Then Directx8.RenderTexture Tex_Surface(1), 0, 0, 0, 0, 800, 600, 800, 600
            ' fading in/out to first screen
            DrawFader
        Case 2, 3
            ' render background
            If Not faderAlpha = 255 Then Directx8.RenderTexture Tex_Surface(2), 0, 0, 0, 0, 800, 600, 800, 600
            ' fading in to second screen
            DrawFader
    End Select
    
    ' render menu
    If faderState >= 4 And Not faderAlpha = 255 Then
        ' render background
        Directx8.RenderTexture Tex_Surface(3), 0, 0, 0, 0, 800, 600, 800, 600
        Directx8.RenderTexture Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64
        Directx8.RenderTexture Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64
        
        ' render menu block
        DrawMainMenu
        Directx8.RenderTexture Tex_Cursor(1), GlobalX, GlobalY, 0, 0, 32, 32, 32, 32
    End If
    
    ' render last fader
    If faderState >= 4 Then
        ' fading in to menu
        If Not faderAlpha = 255 Then DrawFader
    End If
    
    If IsConnecting Then
        If faderState >= 4 And Not faderAlpha = 255 Then
            Call Directx8.RenderTexture(Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, D3DColorARGB(150, 0, 0, 0))
            Call RenderText(Font_Default, Trim$(Menu_Alert_Message), 400 - (EngineGetTextWidth(Font_Default, Menu_Alert_Message) \ 2), 580, Menu_Alert_Colour)
        End If
    End If
    
    ' End the rendering
    Call D3DDevice8.EndScene
    If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Or D3DDevice8.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Directx8.DeviceLost
        Exit Sub
    Else
        Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Render_Menu", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, Colour As Long, X As Long, Y As Long, renderState As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > Count_Fog Then Exit Sub
    Colour = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)
    
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            Directx8.RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Colour
        Next
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawFog", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    Directx8.RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long, ByVal Layer As MapLayer)
    With Map.Tile(X, Y)
        ' skip tile if tileset isn't set
        If Autotile(X, Y).Layer(Layer).renderState = RENDER_STATE_NORMAL Then
            ' Draw normally
            Directx8.RenderTexture Tex_Tileset(.Layer(Layer).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(Layer).X * 32, .Layer(Layer).Y * 32, 32, 32, 32, 32
        ElseIf Autotile(X, Y).Layer(Layer).renderState = RENDER_STATE_AUTOTILE Then
            ' Draw autotiles
            DrawAutoTile Layer, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
            DrawAutoTile Layer, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
            DrawAutoTile Layer, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
            DrawAutoTile Layer, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
        End If
    End With
End Sub

Public Sub DrawBars()
Dim Left As Long, Top As Long, Width As Long, Height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, I As Long, npcNum As Long

    ' dynamic bar calculations
    Width = gTexture(Tex_Bars).RWidth
    Height = gTexture(Tex_Bars).RHeight / 6
    
    ' render npc health bars
    For I = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(I).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(I).Vital(Vitals.HP) > 0 And MapNpc(I).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(I).X * PIC_X + MapNpc(I).xOffset + 16 - (Width / 2)
                tmpY = MapNpc(I).Y * PIC_Y + MapNpc(I).yOffset + 35
                
                ' calculate the width to fill
                If Width > 0 Then BarWidth_NpcHP_Max(I) = ((MapNpc(I).Vital(Vitals.HP) / Width) / (Npc(npcNum).HP / Width)) * Width
                
                ' draw bar background
                Top = Height * 1 ' HP bar background
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
                
                ' draw the bar proper
                Top = 0 ' HP bar
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_NpcHP(I), Height, BarWidth_NpcHP(I), Height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (Width / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + (Height * 2) + 1
            
            ' calculate the width to fill
            If Width > 0 Then barWidth = (timeGetTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * Width
            
            ' draw bar background
            Top = Height * 5 ' cooldown bar background
            Left = 0
            Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
             
            ' draw the bar proper
            Top = Height * 4 ' cooldown bar
            Left = 0
            Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, barWidth, Height, barWidth, Height
        End If
    End If
    
    ' draw all players hp and mp bars
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
            If GetPlayerVital(I, Vitals.HP) > 0 And GetPlayerVital(I, Vitals.HP) < GetPlayerMaxVital(I, Vitals.HP) Then
                ' lock to Player
                tmpX = GetPlayerX(I) * PIC_X + Player(I).xOffset + 16 - (Width / 2)
                tmpY = GetPlayerY(I) * PIC_X + Player(I).yOffset + 35
               
                ' calculate the width to fill
                If Width > 0 Then BarWidth_PlayerHP_Max(I) = ((GetPlayerVital(I, Vitals.HP) / Width) / (GetPlayerMaxVital(I, Vitals.HP) / Width)) * Width
               
                ' draw bar background
                Top = Height * 1 ' HP bar background
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
               
                ' draw the bar proper
                Top = 0 ' HP bar
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_PlayerHP(I), Height, BarWidth_PlayerHP(I), Height
            End If
            If GetPlayerVital(I, Vitals.MP) > 0 And GetPlayerVital(I, Vitals.MP) < GetPlayerMaxVital(I, Vitals.MP) Then
                ' lock to Player
                tmpX = GetPlayerX(I) * PIC_X + Player(I).xOffset + 16 - (Width / 2)
                tmpY = GetPlayerY(I) * PIC_X + Player(I).yOffset + 35 + Height + 1
               
                ' calculate the width to fill
                If Width > 0 Then BarWidth_PlayerMP_Max(I) = ((GetPlayerVital(I, Vitals.MP) / Width) / (GetPlayerMaxVital(I, Vitals.MP) / Width)) * Width
               
                ' draw bar background
                Top = Height * 3 ' MP bar background
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, Width, Height, Width, Height
               
                ' draw the bar proper
                Top = Height * 2 ' MP bar
                Left = 0
                Directx8.RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), Left, Top, BarWidth_PlayerMP(I), Height, BarWidth_PlayerMP(I), Height
            End If
        End If
    Next I
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, X As Long, Y As Long, I As Long, MaxWidth As Long, X2 As Long, Y2 As Long, Colour As Long
    
    With chatBubble(Index)
        If .TargetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' change the colour depending on access
                Colour = DarkBrown
                
                ' it's on our map - get co-ords
                X = ConvertMapX((Player(.target).X * 32) + Player(.target).xOffset) + 16
                Y = ConvertMapY((Player(.target).Y * 32) + Player(.target).yOffset) - 32
                
                ' word wrap the text
                WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
                ' find max width
                For I = 1 To UBound(theArray)
                    If EngineGetTextWidth(Font_Default, theArray(I)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(I))
                Next
                
                ' calculate the new position
                X2 = X - (MaxWidth \ 2)
                Y2 = Y - (UBound(theArray) * 12)
                
                ' render bubble - top left
                Directx8.RenderTexture Tex_GUI(23), X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5
                ' top right
                Directx8.RenderTexture Tex_GUI(23), X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5
                ' top
                Directx8.RenderTexture Tex_GUI(23), X2, Y2 - 5, 9, 0, MaxWidth, 5, 5, 5
                ' bottom left
                Directx8.RenderTexture Tex_GUI(23), X2 - 9, Y, 0, 19, 9, 6, 9, 6
                ' bottom right
                Directx8.RenderTexture Tex_GUI(23), X2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6
                ' bottom - left half
                Directx8.RenderTexture Tex_GUI(23), X2, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                ' bottom - right half
                Directx8.RenderTexture Tex_GUI(23), X2 + (MaxWidth \ 2) + 6, Y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                ' left
                Directx8.RenderTexture Tex_GUI(23), X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
                ' right
                Directx8.RenderTexture Tex_GUI(23), X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
                ' center
                Directx8.RenderTexture Tex_GUI(23), X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
                ' little pointy bit
                Directx8.RenderTexture Tex_GUI(23), X - 5, Y, 58, 19, 11, 11, 11, 11
                
                ' render each line centralised
                For I = 1 To UBound(theArray)
                    RenderText Font_Georgia, theArray(I), X - (EngineGetTextWidth(Font_Default, theArray(I)) / 2), Y2, Colour
                    Y2 = Y2 + 12
                Next
            End If
        End If
        ' check if it's timed out - close it if so
        If .timer + 5000 < timeGetTime Then
            .active = False
        End If
    End With
End Sub

Public Function isConstAnimated(ByVal Sprite As Long) As Boolean
    isConstAnimated = False
    Select Case Sprite
        Case 130, 131, 134, 135, 136, 146, 149, 152
            isConstAnimated = True
    End Select
End Function

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    ' pre-load sprite for calculations
    Sprite = GetPlayerSprite(Index)
    'SetTexture Tex_Char(Sprite)

    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    If Not isConstAnimated(GetPlayerSprite(Index)) Then
        ' Reset frame
        Anim = 2
        ' Check for attacking animation
        If Player(Index).AttackTimer + (attackspeed / 2) > timeGetTime Then
            If Player(Index).Attacking = 1 Then
                Anim = 1
            End If
        Else
            ' If not attacking, walk normally
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
                Case DIR_DOWN
                    If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
                Case DIR_LEFT
                    If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
                Case DIR_RIGHT
                    If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
            End Select
        End If
    Else
        If Player(Index).AnimTimer + 100 <= timeGetTime Then
            Player(Index).Anim = Player(Index).Anim + 1
            If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
            Player(Index).AnimTimer = timeGetTime
        End If
        Anim = Player(Index).Anim
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    ' Set the left
    Select Case GetPlayerDir(Index)
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
        .Top = spritetop * (gTexture(Tex_Char(Sprite)).RHeight / 4)
        .Height = (gTexture(Tex_Char(Sprite)).RHeight / 4)
        .Left = Anim * (gTexture(Tex_Char(Sprite)).RWidth / 4)
        .Width = (gTexture(Tex_Char(Sprite)).RWidth / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((gTexture(Tex_Char(Sprite)).RWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (gTexture(Tex_Char(Sprite)).RHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((gTexture(Tex_Char(Sprite)).RHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
    End If
    
    Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
    ' check for paperdolling
    For I = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(I)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll > 0 Then
                Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
End Sub
Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
Dim rec As GeomRec

    If Sprite < 1 Or Sprite > Count_Paperdoll Then Exit Sub
    
        With rec
            .Top = spritetop * (gTexture(Tex_Paperdoll(Sprite)).RHeight / 4)
            .Height = (gTexture(Tex_Paperdoll(Sprite)).RHeight / 4)
            .Left = Anim * (gTexture(Tex_Paperdoll(Sprite)).RWidth / 4)
            .Width = (gTexture(Tex_Paperdoll(Sprite)).RWidth / 4)
        End With

    ' Clip to screen
    If Y2 < 0 Then
        With rec
            .Top = .Top - Y2
        End With
        Y2 = 0
    End If

    If X2 < 0 Then
        With rec
            .Left = .Left - X2
        End With
        X2 = 0
    End If
       
    Directx8.RenderTexture Tex_Paperdoll(Sprite), ConvertMapX(X2), ConvertMapY(Y2), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
        
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub ' no npc set
    
    ' pre-load texture for calculations
    Sprite = Npc(MapNpc(MapNpcNum).Num).Sprite
    'SetTexture Tex_Char(Sprite)

    If Sprite < 1 Or Sprite > Count_Char Then Exit Sub

    attackspeed = 1000

    If Not isConstAnimated(Npc(MapNpc(MapNpcNum).Num).Sprite) Then
        ' Reset frame
        Anim = 2
        ' Check for attacking animation
        If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > timeGetTime Then
            If MapNpc(MapNpcNum).Attacking = 1 Then
                Anim = 1
            End If
        Else
            ' If not attacking, walk normally
            Select Case MapNpc(MapNpcNum).Dir
                Case DIR_UP
                    If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_DOWN
                    If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_LEFT
                    If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_RIGHT
                    If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            End Select
        End If
    Else
        With MapNpc(MapNpcNum)
            If .AnimTimer + 100 <= timeGetTime Then
                .Anim = .Anim + 1
                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = timeGetTime
            End If
            Anim = .Anim
        End With
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
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
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((gTexture(Tex_Char(Sprite)).RWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (gTexture(Tex_Char(Sprite)).RHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((gTexture(Tex_Char(Sprite)).RHeight / 4) - 32) - 4
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If
    
    Directx8.RenderTexture Tex_Char(Sprite), ConvertMapX(X), ConvertMapY(Y), rec.Left, rec.Top, rec.Width, rec.Height, rec.Width, rec.Height
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim Width As Long, Height As Long
    
    ' calculations
    Width = gTexture(Tex_Target).RWidth / 2
    Height = gTexture(Tex_Target).RHeight
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    Directx8.RenderTexture Tex_Target, X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawTargetHover()
Dim I As Long, X As Long, Y As Long, Width As Long, Height As Long

    Width = gTexture(Tex_Target).RWidth / 2
    Height = gTexture(Tex_Target).RHeight
    
    If Width <= 0 Then Width = 1
    If Height <= 0 Then Height = 1
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            X = (Player(I).X * 32) + Player(I).xOffset + 32
            Y = (Player(I).Y * 32) + Player(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    Directx8.RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 32 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If
    Next
    
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            X = (MapNpc(I).X * 32) + MapNpc(I).xOffset + 32
            Y = (MapNpc(I).Y * 32) + MapNpc(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    X = ConvertMapX(X)
                    Y = ConvertMapY(Y)
                    Directx8.RenderTexture Tex_Target, X - 16 - (Width / 2), Y - 32 - (Height / 2), Width, 0, Width, Height, Width, Height
                End If
            End If
        End If
    Next
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
    'SetTexture Tex_Resource(Resource_sprite)

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Resource(Resource_sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Resource(Resource_sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (gTexture(Tex_Resource(Resource_sprite)).RWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - gTexture(Tex_Resource(Resource_sprite)).RHeight + 32
    
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    Directx8.RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawItem(ByVal itemNum As Long)
Dim PicNum As Integer, dontRender As Boolean, I As Long, tmpIndex As Long
    
    PicNum = Item(MapItem(itemNum).Num).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

     ' if it's not us then don't render
    If MapItem(itemNum).playerName <> vbNullString Then
        If Trim$(MapItem(itemNum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For I = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(I)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemNum).playerName) Then
                        If MapItem(itemNum).bound = 0 Then
                            dontRender = False
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    If Not dontRender Then Directx8.RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemNum).X * PIC_X), ConvertMapY(MapItem(itemNum).Y * PIC_Y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemNum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemNum > 0 Then Exit Sub
    
    PicNum = Item(itemNum).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

    Directx8.RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = PlayerSpells(DragSpell)
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = Spell(spellnum).Icon

    If PicNum < 1 Or PicNum > Count_Spellicon Then Exit Sub

    Directx8.RenderTexture Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Integer, sRECT As GeomRec, I As Long, Width As Long, Height As Long, looptime As Long, FrameCount As Long
Dim X As Long, Y As Long, LockIndex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > Count_Anim Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = 192 'gTexture(tex_anim(sprite)).rwidth / frameCount
    Height = 192 'gTexture(tex_anim(sprite)).rheight
    
    With sRECT
        .Top = (Height * ((AnimInstance(Index).frameIndex(Layer) - 1) \ AnimColumns))
        .Height = Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(Layer) - 1) Mod AnimColumns)))
        .Width = Width
    End With
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' check if is ingame
            If IsPlaying(LockIndex) Then
                ' check if on same map
                If GetPlayerMap(LockIndex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(LockIndex) * PIC_X) + 16 - (Width / 2) + Player(LockIndex).xOffset
                    Y = (GetPlayerY(LockIndex) * PIC_Y) + 16 - (Height / 2) + Player(LockIndex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' check if NPC exists
            If MapNpc(LockIndex).Num > 0 Then
                ' check if alive
                If MapNpc(LockIndex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(LockIndex).X * PIC_X) + 16 - (Width / 2) + MapNpc(LockIndex).xOffset
                    Y = (MapNpc(LockIndex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(LockIndex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    Directx8.RenderTexture Tex_Anim(Sprite), X, Y, sRECT.Left, sRECT.Top, sRECT.Width, sRECT.Height, sRECT.Width, sRECT.Height
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_INVENTORY).visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, invSlot) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc GetPlayerInvItemNum(MyIndex, invSlot), GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y, isSB
            ' value
            If InShop > 0 Then
                DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).Y + GUIWindow(GUI_DESCRIPTION).Height + 50
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
    
    If Not GUIWindow(GUI_SHOP).visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).Item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).Item, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y
            DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).X + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).Y + GUIWindow(GUI_DESCRIPTION).Height + 50
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_CHARACTER).visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If GetPlayerEquipment(MyIndex, eqSlot) > 0 Then
            If Item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).Y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal X As Long, ByVal Y As Long)
Dim CostItem As Long, CostValue As Long, itemNum As Long, sString As String, Width As Long, Height As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    ' draw the window
    Width = 190
    Height = 36

    Directx8.RenderTexture Tex_GUI(33), X, Y, 0, 0, Width, Height, Width, Height
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemNum = GetPlayerInvItemNum(MyIndex, slotNum)
        If itemNum = 0 Then Exit Sub
        CostItem = 1
        CostValue = (Item(itemNum).Price / 100) * Shop(InShop).BuyRate
        sString = "The shop will buy for"
    Else
        itemNum = Shop(InShop).TradeItem(slotNum).Item
        If itemNum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue
        sString = "The shop will sell for"
    End If
    
    Directx8.RenderTexture Tex_Item(Item(CostItem).Pic), X + 155, Y + 2, 0, 0, 32, 32, 32, 32
    
    RenderText Font_Default, sString, X + 4, Y + 3, DarkGrey
    
    RenderText Font_Default, CostValue & " " & Trim$(Item(CostItem).name), X + 4, Y + 18, White
End Sub

Public Sub DrawItemDesc(ByVal itemNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal soulBound As Boolean = False)
Dim Colour As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, I As Long, Width As Long, Height As Long
    
    ' get out
    If itemNum = 0 Then Exit Sub

    ' render the window
    Width = GUIWindow(GUI_DESCRIPTION).Width
    If Not LenB(Trim$(Item(itemNum).Desc)) = 0 Then
        Height = 210
    Else
        Height = GUIWindow(GUI_DESCRIPTION).Height
    End If
    Directx8.RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Item(itemNum).Pic > 0 Then
        ' render sprite
        Directx8.RenderTexture Tex_Item(Item(itemNum).Pic), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not LenB(Trim$(Item(itemNum).Desc)) = 0 Then
        RenderText Font_Default, WordWrap(Trim$(Item(itemNum).Desc), Width - 20), X + 10, Y + 128, White
    End If
    
    ' work out name colour
    Select Case Item(itemNum).Rarity
        Case 0 ' white
            Colour = White
        Case 1 ' green
            Colour = Green
        Case 2 ' blue
            Colour = Blue
        Case 3 ' maroon
            Colour = Red
        Case 4 ' purple
            Colour = Pink
        Case 5 ' orange
            Colour = Brown
    End Select
    
    If Not soulBound Then
        theName = Trim$(Item(itemNum).name)
    Else
        theName = "(SB) " & Trim$(Item(itemNum).name)
    End If
    
    ' render name
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, Colour
    
    ' class req
    If Item(itemNum).ClassReq > 0 Then
        className = Trim$(Class(Item(itemNum).ClassReq).name)
        ' do we match it?
        If GetPlayerClass(MyIndex) = Item(itemNum).ClassReq Then
            Colour = Green
        Else
            Colour = BrightRed
        End If
    Else
        className = "No class req."
        Colour = Green
    End If
    RenderText Font_Default, className, X + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), Y + 92, Colour
    
    ' level
    If Item(itemNum).LevelReq > 0 Then
        levelTxt = "Level " & Item(itemNum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(itemNum).LevelReq Then
            Colour = Green
        Else
            Colour = BrightRed
        End If
    Else
        levelTxt = "No level req."
        Colour = Green
    End If
    RenderText Font_Default, levelTxt, X + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), Y + 107, Colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_NONE
            sInfo(I) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(I) = "Weapon"
        Case ITEM_TYPE_ARMOR
            sInfo(I) = "Armour"
        Case ITEM_TYPE_HELMET
            sInfo(I) = "Helmet"
        Case ITEM_TYPE_SHIELD
            sInfo(I) = "Shield"
        Case ITEM_TYPE_CONSUME
            sInfo(I) = "Consume"
        Case ITEM_TYPE_SPELL
            sInfo(I) = "Spell"
    End Select
    ' binding
    If Item(itemNum).BindType = 1 Then
        I = I + 1
        ReDim Preserve sInfo(1 To I) As String
        sInfo(I) = "Bind on Pickup"
    ElseIf Item(itemNum).BindType = 2 Then
        I = I + 1
        ReDim Preserve sInfo(1 To I) As String
        sInfo(I) = "Bind on Equip"
    End If
    ' price
    I = I + 1
    ReDim Preserve sInfo(1 To I) As String
    sInfo(I) = "Value: " & Item(itemNum).Price & "g"
    
    ' more info
    Select Case Item(itemNum).Type
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' damage/defence
            If Item(itemNum).Type = ITEM_TYPE_WEAPON Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Damage: " & Item(itemNum).Data2
                ' speed
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Speed: " & (Item(itemNum).Speed / 1000) & "s"
            Else
                If Item(itemNum).Data2 > 0 Then
                    I = I + 1
                    ReDim Preserve sInfo(1 To I) As String
                    sInfo(I) = "Defence: " & Item(itemNum).Data2
                End If
            End If
            ' stat bonuses
            If Item(itemNum).Add_Stat(Stats.Strength) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(itemNum).Add_Stat(Stats.Endurance) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(itemNum).Add_Stat(Stats.Intelligence) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(itemNum).Add_Stat(Stats.Agility) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(itemNum).Add_Stat(Stats.Willpower) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If Item(itemNum).AddHP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddHP & " HP"
            End If
            If Item(itemNum).AddMP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddMP & " SP"
            End If
            If Item(itemNum).AddEXP > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "+" & Item(itemNum).AddEXP & " EXP"
            End If
        Case ITEM_TYPE_SPELL
            I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Damage: " & Item(itemNum).Data2
            ' price
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            If Item(itemNum).Data1 > 0 Then
                sInfo(I) = "Learn spell: " & Trim$(Spell(Item(itemNum).Data1).name)
            Else
                sInfo(I) = "Learn spell: None"
            End If
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For I = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(I), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), Y, White
    Next
End Sub

Public Sub DrawInventory()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long
Dim Amount As String
Dim Colour As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(5), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(5), GUIWindow(GUI_INVENTORY).X, GUIWindow(GUI_INVENTORY).Y, 0, 0, Width, Height, Width, Height
    
    For I = 1 To MAX_INV
        Top = GUIWindow(GUI_INVENTORY).Y + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
        Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
        
        itemNum = GetPlayerInvItemNum(MyIndex, I)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    If TradeYourOffer(X).Num = I Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = I Then GoTo NextLoop
                
            If ItemPic > 0 And ItemPic <= Count_Item Then
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32

                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                    Y = Top + 21
                    X = Left - 4
                    Amount = CStr(GetPlayerInvItemValue(MyIndex, I))
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
NextLoop:
    Next
    DrawInventoryItemDesc
End Sub

Public Sub DrawHotbarSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_HOTBAR).visible Then Exit Sub
    
    spellSlot = IsHotbarSlot(GlobalX, GlobalY)
    If spellSlot > 0 Then
        Select Case Hotbar(spellSlot).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(spellSlot).Slot).name) > 0 Then
                    DrawItemDesc Hotbar(spellSlot).Slot, GUIWindow(GUI_HOTBAR).X, GUIWindow(GUI_HOTBAR).Y + GUIWindow(GUI_HOTBAR).Height + 10
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(spellSlot).Slot).name) > 0 Then
                    DrawSpellDesc Hotbar(spellSlot).Slot, GUIWindow(GUI_HOTBAR).X, GUIWindow(GUI_HOTBAR).Y + GUIWindow(GUI_HOTBAR).Height + 10
                End If
        End Select
    End If
End Sub

Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_SPELLS).visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot) > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot), GUIWindow(GUI_SPELLS).X - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).Y, spellSlot
        End If
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal spellSlot As Long = 0)
Dim Colour As Long, theName As String, sUse As String, sInfo() As String, I As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, Height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    Width = GUIWindow(GUI_DESCRIPTION).Width
    If Not LenB(Trim$(Spell(spellnum).Desc)) = 0 Then
        Height = 210
    Else
        Height = GUIWindow(GUI_DESCRIPTION).Height
    End If
    'EngineRenderRectangle Tex_GUI(34), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(8), X, Y, 0, 0, Width, Height, Width, Height
    
    ' make sure it has a sprite
    If Spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        Directx8.RenderTexture Tex_Spellicon(Spell(spellnum).Icon), X + 16, Y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    If Not LenB(Trim$(Spell(spellnum).Desc)) = 0 Then
        RenderText Font_Default, WordWrap(Trim$(Spell(spellnum).Desc), Width - 20), X + 10, Y + 128, White
    End If
    
    ' render name
    Colour = White
    theName = Trim$(Spell(spellnum).name)
    RenderText Font_Default, theName, X + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), Y + 6, Colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    I = 1
    ReDim Preserve sInfo(1 To I) As String
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_VITALCHANGE
            sInfo(I) = "Change vitals"
        Case SPELL_TYPE_WARP
            sInfo(I) = "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_VITALCHANGE
            ' damage
            If Spell(spellnum).Vital(Vitals.HP) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    sInfo(I) = "-" & Spell(spellnum).Vital(Vitals.HP) & " HP"
                Else
                    sInfo(I) = "+" & Spell(spellnum).Vital(Vitals.HP) & " HP"
                End If
            End If
            
            If Spell(spellnum).Vital(Vitals.MP) > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                If Spell(spellnum).VitalType(Vitals.HP) = 0 Then
                    sInfo(I) = "-" & Spell(spellnum).Vital(Vitals.MP) & " MP"
                Else
                    sInfo(I) = "+" & Spell(spellnum).Vital(Vitals.MP) & " MP"
                End If
            End If
            
            ' mp cost
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            I = I + 1
            ReDim Preserve sInfo(1 To I) As String
            sInfo(I) = "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                I = I + 1
                ReDim Preserve sInfo(1 To I) As String
                sInfo(I) = "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
    End Select
    
    ' go through and render all this shit
    Y = Y + 12
    For I = 1 To UBound(sInfo)
        Y = Y + 12
        RenderText Font_Default, sInfo(I), X + 141 - (EngineGetTextWidth(Font_Default, sInfo(I)) \ 2), Y, White
    Next
End Sub

Public Sub DrawSkills()
Dim I As Long, X As Long, Y As Long, spellnum As Long, spellpic As Long
Dim Top As Long, Left As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(5), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(5), GUIWindow(GUI_SPELLS).X, GUIWindow(GUI_SPELLS).Y, 0, 0, Width, Height, Width, Height
    
    ' render skills
    For I = 1 To MAX_PLAYER_SPELLS
        Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
        Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
        spellnum = PlayerSpells(I)

        ' make sure not dragging it
        If DragSpell = I Then GoTo NextLoop
        
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= Count_Spellicon Then
                If SpellCD(I) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    Directx8.RenderTexture Tex_Spellicon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    Directx8.RenderTexture Tex_Spellicon(spellpic), Left, Top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
    DrawPlayerSpellDesc
End Sub

Public Sub DrawEquipment()
Dim X As Long, Y As Long, I As Long
Dim itemNum As Long, ItemPic As Long

    For I = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, I)

        ' get the item sprite
        If itemNum > 0 Then
            ItemPic = Tex_Item(Item(itemNum).Pic)
        Else
            ' no item equiped - use blank image
            ItemPic = Tex_GUI(8 + I)
        End If
        
        Y = GUIWindow(GUI_CHARACTER).Y + EqTop
        X = GUIWindow(GUI_CHARACTER).X + EqLeft + ((EqOffsetX + 32) * (((I - 1) Mod EqColumns)))

        'EngineRenderRectangle itempic, x, y, 0, 0, 32, 32, 32, 32, 32, 32
        Directx8.RenderTexture ItemPic, X, Y, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawCharacter()
Dim X As Long, Y As Long, I As Long, dX As Long, dY As Long, tmpString As String, buttonnum As Long
Dim Width As Long, Height As Long
    
    X = GUIWindow(GUI_CHARACTER).X
    Y = GUIWindow(GUI_CHARACTER).Y
    
    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(6), X, Y, 0, 0, Width, Height, Width, Height
    
    ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, X + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), Y + 9, White
    
    ' render stats
    dX = X + 20
    dY = Y + 145
    RenderText Font_Default, "Str: " & GetPlayerStat(MyIndex, Strength), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "End: " & GetPlayerStat(MyIndex, Endurance), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Int: " & GetPlayerStat(MyIndex, Intelligence), dX, dY, White
    dY = Y + 145
    dX = dX + 80
    RenderText Font_Default, "Agi: " & GetPlayerStat(MyIndex, Agility), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Will: " & GetPlayerStat(MyIndex, Willpower), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Pnts: " & GetPlayerPOINTS(MyIndex), dX, dY, White
    
    ' draw the face
    If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= Count_Face Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
        Directx8.RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), X + 49, Y + 38, 0, 0, gTexture(Tex_Face(GetPlayerSprite(MyIndex))).RWidth * 2, gTexture(Tex_Face(GetPlayerSprite(MyIndex))).RHeight * 2, gTexture(Tex_Face(GetPlayerSprite(MyIndex))).RWidth, gTexture(Tex_Face(GetPlayerSprite(MyIndex))).RHeight
    End If
    
    ' draw the equipment
    DrawEquipment
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = Button_AddStats1 To Button_AddStats5
            X = GUIWindow(GUI_CHARACTER).X + Buttons(buttonnum).X
            Y = GUIWindow(GUI_CHARACTER).Y + Buttons(buttonnum).Y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                Height = Buttons(buttonnum).Height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    FMOD.Sound_Play Sound_ButtonHover
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
    
    DrawCharacterItemDesc
End Sub

Public Sub DrawOptions()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(29), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(29), GUIWindow(GUI_OPTIONS).X, GUIWindow(GUI_OPTIONS).Y, 0, 0, Width, Height, Width, Height
    
    ' draw buttons
    For I = Button_MusicOn To Button_FullscreenOff
        ' set co-ordinate
        X = GUIWindow(GUI_OPTIONS).X + Buttons(I).X
        Y = GUIWindow(GUI_OPTIONS).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawParty()
Dim I As Long, X As Long, Y As Long, Width As Long, playerNum As Long, theName As String
Dim Height As Long

    ' render the window
    Width = 195
    Height = 250
    'EngineRenderRectangle Tex_GUI(7), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(7), GUIWindow(GUI_PARTY).X, GUIWindow(GUI_PARTY).Y, 0, 0, Width, Height, Width, Height
    
    ' draw the bars
    If Party.Leader > 0 Then ' make sure we're in a party
        ' draw leader
        playerNum = Party.Leader
        ' name
        theName = Trim$(GetPlayerName(playerNum))
        ' draw name
        Y = GUIWindow(GUI_PARTY).Y + 12
        X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
        RenderText Font_Default, theName, X, Y, White
        ' draw hp
        Y = GUIWindow(GUI_PARTY).Y + 29
        X = GUIWindow(GUI_PARTY).X + 6
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
        End If
        'EngineRenderRectangle Tex_GUI(16), x, y, 0, 0, width, 9, width, 9, width, 9
        Directx8.RenderTexture Tex_GUI(16), X, Y, 0, 0, Width, 9, Width, 9
        ' draw mp
        Y = GUIWindow(GUI_PARTY).Y + 38
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
        End If
        'EngineRenderRectangle Tex_GUI(17), x, y, 0, 0, width, 9, width, 9, width, 9
        Directx8.RenderTexture Tex_GUI(17), X, Y, 0, 0, Width, 9, Width, 9
        
        ' draw members
        For I = 1 To MAX_PARTY_MEMBERS
            If Party.Member(I) > 0 Then
                If Party.Member(I) <> Party.Leader Then
                    ' cache the index
                    playerNum = Party.Member(I)
                    ' name
                    theName = Trim$(GetPlayerName(playerNum))
                    ' draw name
                    Y = GUIWindow(GUI_PARTY).Y + 12 + ((I - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
                    RenderText Font_Default, theName, X, Y, White
                    ' draw hp
                    Y = GUIWindow(GUI_PARTY).Y + 29 + ((I - 1) * 49)
                    X = GUIWindow(GUI_PARTY).X + 6
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(16), x, y, 0, 0, width, 9, width, 9, width, 9
                    Directx8.RenderTexture Tex_GUI(16), X, Y, 0, 0, Width, 9, Width, 9
                    ' draw mp
                    Y = GUIWindow(GUI_PARTY).Y + 38 + ((I - 1) * 49)
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(17), x, y, 0, 0, width, 9, width, 9, width, 9
                    Directx8.RenderTexture Tex_GUI(17), X, Y, 0, 0, Width, 9, Width, 9
                End If
            End If
        Next
    End If
    
    ' draw buttons
    For I = Button_PartyInvite To Button_PartyDisband
        ' set co-ordinate
        X = GUIWindow(GUI_PARTY).X + Buttons(I).X
        Y = GUIWindow(GUI_PARTY).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawHotbar()
Dim I As Long, X As Long, Y As Long, t As Long, sS As String
Dim Width As Long, Height As Long

    For I = 1 To MAX_HOTBAR
        ' draw the box
        X = GUIWindow(GUI_HOTBAR).X + ((I - 1) * (5 + 36))
        Y = GUIWindow(GUI_HOTBAR).Y
        Width = 36
        Height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        Directx8.RenderTexture Tex_GUI(2), X, Y, 0, 0, Width, Height, Width, Height
        ' draw the icon
        Select Case Hotbar(I).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(I).Slot).name) > 0 Then
                    If Item(Hotbar(I).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        Directx8.RenderTexture Tex_Item(Item(Hotbar(I).Slot).Pic), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(I).Slot).name) > 0 Then
                    If Spell(Hotbar(I).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        Directx8.RenderTexture Tex_Spellicon(Spell(Hotbar(I).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t) > 0 Then
                                If PlayerSpells(t) = Hotbar(I).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        Directx8.RenderTexture Tex_Spellicon(Spell(Hotbar(I).Slot).Icon), X + 2, Y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = str(I)
        If I = 10 Then sS = "0"
        If I = 11 Then sS = " -"
        If I = 12 Then sS = " ="
        RenderText Font_Default, sS, X + 4, Y + 20, White
    Next
    DrawHotbarSpellDesc
End Sub

Public Sub DrawGUI()
    ' render shadow
    'EngineRenderRectangle Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
    'EngineRenderRectangle Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
    Directx8.RenderTexture Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64
    Directx8.RenderTexture Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    
    If GUIWindow(GUI_TUTORIAL).visible Then
        DrawTutorial
        Exit Sub
    End If
    
    If GUIWindow(GUI_CHAT).visible Then DrawChat
    
    If GUIWindow(GUI_EVENTCHAT).visible Then DrawEventChat
    If GUIWindow(GUI_CURRENCY).visible Then DrawCurrency
    If GUIWindow(GUI_DIALOGUE).visible Then DrawDialogue
    
    ' render bars
    DrawGUIBars
    
    ' render fps
    If BFPS Then
        If FPS_Lock Then
            RenderText Font_Default, "FPS: " & Round(GameFPS / 1500) & " Ping: " & CStr(Ping), 303, 48, White
        Else
            RenderText Font_Default, "FPS: " & GameFPS & " Ping: " & CStr(Ping), 303, 48, White
        End If
    End If
    
    ' render menu
    DrawMenu
    
    ' render hotbar
    DrawHotbar
    
    ' render menus
    If GUIWindow(GUI_INVENTORY).visible Then DrawInventory
    If GUIWindow(GUI_SPELLS).visible Then DrawSkills
    If GUIWindow(GUI_CHARACTER).visible Then DrawCharacter
    If GUIWindow(GUI_OPTIONS).visible Then DrawOptions
    If GUIWindow(GUI_PARTY).visible Then DrawParty
    If GUIWindow(GUI_SHOP).visible Then DrawShop
    If GUIWindow(GUI_TRADE).visible Then DrawTrade
    If GUIWindow(GUI_BANK).visible Then DrawBank
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
End Sub
Public Sub DrawChat()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long
    If chatOn Then
        ' render chatbox
        Width = 412
        Height = 145
        'EngineRenderRectangle Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height, width, height
        Directx8.RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y, 0, 0, Width, Height, Width, Height
        RenderChatTextBuffer
        ' render the message input
        RenderText Font_Default, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).X + 38, GUIWindow(GUI_CHAT).Y + 126, White
        ' draw buttons
        For I = Button_ChatUp To Button_ChatDown
            ' set co-ordinate
            X = GUIWindow(GUI_CHAT).X + Buttons(I).X
            Y = GUIWindow(GUI_CHAT).Y + Buttons(I).Y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    FMOD.Sound_Play Sound_ButtonHover
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        Next
    Else
        Directx8.RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).X, GUIWindow(GUI_CHAT).Y + 122, 0, 122, 412, 22, 412, 22
    End If
    RenderChatTextBuffer
End Sub

Public Sub DrawTutorial()
Dim X As Long, Y As Long, I As Long, Width As Long
Dim Height As Long

    X = GUIWindow(GUI_TUTORIAL).X
    Y = GUIWindow(GUI_TUTORIAL).Y - 107
    
    ' render chatbox
    Width = GUIWindow(GUI_TUTORIAL).Width
    Height = GUIWindow(GUI_TUTORIAL).Height
    'EngineRenderRectangle Tex_GUI(30), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(30), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(chatText, 260), X + 200, Y + 129, White
    
    ' Draw replies
    For I = 1 To 4
        If Len(Trim$(tutOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(tutOpt(I)) & "]")
            X = GUIWindow(GUI_CHAT).X + 200 + (130 - (Width / 2))
            Y = GUIWindow(GUI_CHAT).Y + 115 - ((I - 1) * 15)
            If tutOptState(I) = 2 Then
                ' clicked
                RenderText Font_Default, "[" & Trim$(tutOpt(I)) & "]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[" & Trim$(tutOpt(I)) & "]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = I Then
                        FMOD.Sound_Play Sound_ButtonHover
                        lastNpcChatsound = I
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[" & Trim$(tutOpt(I)) & "]", X, Y, BrightBlue
                    ' reset sound if needed
                    If lastNpcChatsound = I Then lastNpcChatsound = 0
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawEventChat()
Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_EVENTCHAT).X
    Y = GUIWindow(GUI_EVENTCHAT).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_EVENTCHAT).Width
    Height = GUIWindow(GUI_EVENTCHAT).Height
    Directx8.RenderTexture Tex_GUI(27), X, Y, 0, 0, Width, Height, Width, Height
    
    Select Case CurrentEvent.Type
        Case Evt_Menu
            ' Draw replies
            RenderText Font_Default, WordWrap(Trim$(CurrentEvent.Text(1)), GUIWindow(GUI_EVENTCHAT).Width - 10), X + 10, Y + 10, White
            For I = 1 To UBound(CurrentEvent.Text) - 1
                If Len(Trim$(CurrentEvent.Text(I + 1))) > 0 Then
                    Width = EngineGetTextWidth(Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]")
                    X = GUIWindow(GUI_CHAT).X + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
                    Y = GUIWindow(GUI_CHAT).Y + 115 - ((I - 1) * 15)
                    If chatOptState(I) = 2 Then
                        ' clicked
                        RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]", X, Y, Grey
                    Else
                        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                            ' hover
                            RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]", X, Y, Yellow
                            ' play sound if needed
                            If Not lastNpcChatsound = I Then
                                FMOD.Sound_Play Sound_ButtonHover
                                lastNpcChatsound = I
                            End If
                        Else
                            ' normal
                            RenderText Font_Default, "[" & Trim$(CurrentEvent.Text(I + 1)) & "]", X, Y, BrightBlue
                            ' reset sound if needed
                            If lastNpcChatsound = I Then lastNpcChatsound = 0
                        End If
                    End If
                End If
            Next
        Case Evt_Message
            If CurrentEvent.Data(2) = 3 Then
                RenderText Font_Default, WordWrap(Trim$(CurrentEvent.Text(1)), GUIWindow(GUI_EVENTCHAT).Width - 10), X + 10, Y + 10, White
            Else
                RenderText Font_Default, WordWrap(Trim$(CurrentEvent.Text(1)), GUIWindow(GUI_EVENTCHAT).Width - 75), X + 70, Y + 10, White
            End If
            Select Case CurrentEvent.Data(2)
                Case 0: Sprite = GetPlayerSprite(MyIndex)
                Case 1: If CurrentEvent.Data(1) > 0 Then Sprite = Npc(CurrentEvent.Data(1)).Sprite
                Case 2: Sprite = CurrentEvent.Data(1)
                Case 3: Sprite = 0
            End Select
            If Sprite > 0 And Sprite <= Count_Face Then
                'EngineRenderRectangle Tex_Face(sprite), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
                Directx8.RenderTexture Tex_Face(Sprite), X + 3, Y + 3, 0, 0, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight
            End If
            
            If Sprite > 0 And Sprite <= Count_Char Then
                'EngineRenderRectangle Tex_Face(sprite), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
                Directx8.RenderTexture Tex_Char(Sprite), X + 25, Y + 15, 0, 0, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4
            End If
            
            
            Width = EngineGetTextWidth(Font_Default, "[Continue]")
            X = GUIWindow(GUI_EVENTCHAT).X + GUIWindow(GUI_EVENTCHAT).Width - Width - 10
            Y = GUIWindow(GUI_EVENTCHAT).Y + GUIWindow(GUI_EVENTCHAT).Height - 25
            If chatContinueState = 2 Then
                ' clicked
                RenderText Font_Default, "[Continue]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Continue]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = I Then
                        FMOD.Sound_Play Sound_ButtonHover
                        lastNpcChatsound = I
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Continue]", X, Y, BrightBlue
                    ' reset sound if needed
                    If lastNpcChatsound = I Then lastNpcChatsound = 0
                End If
            End If
    End Select
End Sub

Public Sub DrawShop()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long
Dim Width As Long, Height As Long

    ' render the window
    Width = 252
    Height = 317
    'EngineRenderRectangle Tex_GUI(28), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(28), GUIWindow(GUI_SHOP).X, GUIWindow(GUI_SHOP).Y, 0, 0, Width, Height, Width, Height
    
    ' render the shop items
    For I = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(I).Item
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                
                Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(I).ItemValue > 1 Then
                    Y = Top + 22
                    X = Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(I).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    I = Button_ShopExit
        ' set co-ordinate
        X = GUIWindow(GUI_SHOP).X + Buttons(I).X
        Y = GUIWindow(GUI_SHOP).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    
    ' draw item descriptions
    DrawShopItemDesc
End Sub

Public Sub DrawMenu()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long

    ' draw background
    X = GUIWindow(GUI_MENU).X
    Y = GUIWindow(GUI_MENU).Y
    Width = 232
    Height = 76
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(3), X, Y, 0, 0, Width, Height, Width, Height
    
    ' draw buttons
    For I = Button_Inventory To Button_Party
        ' set co-ordinate
        X = GUIWindow(GUI_MENU).X + Buttons(I).X
        Y = GUIWindow(GUI_MENU).Y + Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawMainMenu()
Dim I As Long, X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' draw logo
    Width = gTexture(Tex_GUI(35)).RWidth
    Height = gTexture(Tex_GUI(35)).RHeight
    'EngineRenderRectangle tex_gui(35), 152, 20, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(35), (ScreenWidth / 2) - (Width / 2), 6, 0, 0, Width, Height, Width, Height

    ' draw background
    X = GUIWindow(GUI_MAINMENU).X
    Y = GUIWindow(GUI_MAINMENU).Y
    Width = GUIWindow(GUI_MAINMENU).Width
    Height = GUIWindow(GUI_MAINMENU).Height
    'EngineRenderRectangle Tex_GUI(18), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(18), X, Y, 0, 0, Width, Height, Width, Height
    
    If SStatus = "Online" Then
        RenderText Font_Default, SStatus, X + Width - 35 - EngineGetTextWidth(Font_Default, SStatus), Y + Height - 84, Green
    Else
        RenderText Font_Default, SStatus, X + Width - 35 - EngineGetTextWidth(Font_Default, SStatus), Y + Height - 84, Red
    End If
    RenderText Font_Default, "Server: ", X + Width - 35 - EngineGetTextWidth(Font_Default, "Server: " & SStatus), Y + Height - 84, White
    RenderText Font_Default, Options.Game_Name & " v" & App.Major & "." & App.Minor & "." & App.Revision, X + 35, Y + Height - 84, White
    
    ' draw buttons
    If Not faderAlpha > 0 Then
        For I = Button_Login To Button_Exit
            ' set co-ordinate
            X = GUIWindow(GUI_MAINMENU).X + Buttons(I).X
            Y = GUIWindow(GUI_MAINMENU).Y + Buttons(I).Y
            Width = Buttons(I).Width
            Height = Buttons(I).Height
            ' check for state
            If Buttons(I).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = I Then
                    FMOD.Sound_Play Sound_ButtonHover
                    lastButtonSound = I
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = I Then lastButtonSound = 0
            End If
        Next
    End If

    ' draw specific menus
    Select Case curMenu
        Case MENU_MAIN
            RenderText Font_Default, "Latest News", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "Latest News") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawNews
        Case MENU_LOGIN
            RenderText Font_Default, "Login", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "Login") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawLogin
        Case MENU_REGISTER
            RenderText Font_Default, "Register", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "Register") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawRegister
         Case MENU_CREDITS
            RenderText Font_Default, "Credits", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "Credits") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawCredits
        Case MENU_CLASS
            RenderText Font_Default, "Class Select", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "Class Select") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawClassSelect
        Case MENU_NEWCHAR
            RenderText Font_Default, "New Character", GUIWindow(GUI_MAINMENU).X + ((GUIWindow(GUI_MAINMENU).Width / 2) - (EngineGetTextWidth(Font_Default, "New Character") / 2)), GUIWindow(GUI_MAINMENU).Y + 7, White
            DrawNewChar
    End Select
End Sub

Public Sub DrawNewChar()
Dim X As Long, Y As Long, buttonnum As Long, Sprite As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X
    Y = GUIWindow(GUI_MAINMENU).Y
    
    ' draw the image
    Width = 291
    Height = 107
    'EngineRenderRectangle Tex_GUI(26), x + 110, y + 92, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(26), X + 110, Y + 92, 0, 0, Width, Height, Width, Height
    
    ' char name
    RenderText Font_Default, sChar & chatShowLine, X + 158, Y + 94, White

    If newCharSex = SEX_MALE Then
        Sprite = Class(newCharClass).MaleSprite(newCharSprite)
    Else
        Sprite = Class(newCharClass).FemaleSprite(newCharSprite)
    End If
    
    
    If Sprite > 0 And Sprite <= Count_Face Then
        Directx8.RenderTexture Tex_Face(Sprite), X + 38, Y + 76, 0, 0, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight
    End If
            
    If Sprite > 0 And Sprite <= Count_Char Then
        Directx8.RenderTexture Tex_Char(Sprite), X + 60, Y + 88, 0, 0, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4
    End If
    If Not faderAlpha > 0 Then
        ' position
        buttonnum = Button_NewCharAccept
        X = GUIWindow(GUI_MAINMENU).X + Buttons(buttonnum).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(buttonnum).Y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
        ' position
        For buttonnum = Button_GenderLeft To Button_GenderRight
        X = GUIWindow(GUI_MAINMENU).X + Buttons(buttonnum).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(buttonnum).Y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
        Next

    End If
End Sub

Public Sub DrawClassSelect()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X
    Y = GUIWindow(GUI_MAINMENU).Y
    
    Select Case newCharClass
        Case 1 ' warrior
            Width = 426
            Height = 209
            'EngineRenderRectangle Tex_GUI(23), x + 30, y + 34, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Class(1), X + 30, Y + 34, 0, 0, Width, Height, Width, Height
        Case 2 ' wizard
            Width = 441
            Height = 213
            'EngineRenderRectangle Tex_GUI(24), x + 30, y + 33, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Class(2), X + 30, Y + 33, 0, 0, Width, Height, Width, Height
        Case 3 ' whisperer
            Width = 455
            Height = 212
            'EngineRenderRectangle Tex_GUI(25), x + 30, y + 38, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Class(3), X + 30, Y + 38, 0, 0, Width, Height, Width, Height
        Case Else ' warrior
            Width = 426
            Height = 209
            'EngineRenderRectangle Tex_GUI(23), x + 30, y + 34, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Class(1), X + 30, Y + 34, 0, 0, Width, Height, Width, Height
        
    End Select
    
    If Not faderAlpha > 0 Then
        For buttonnum = 13 To 14
            X = GUIWindow(GUI_MAINMENU).X + Buttons(buttonnum).X
            Y = GUIWindow(GUI_MAINMENU).Y + Buttons(buttonnum).Y
            Width = Buttons(buttonnum).Width
            Height = Buttons(buttonnum).Height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    FMOD.Sound_Play Sound_ButtonHover
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
End Sub

Public Sub DrawNews()
Dim X As Long, Y As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X + 137
    Y = GUIWindow(GUI_MAINMENU).Y + 80
    Width = 224
    Height = 118
    'EngineRenderRectangle Tex_GUI(22), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(22), X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawLogin()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X + 86
    Y = GUIWindow(GUI_MAINMENU).Y + 102
    buttonnum = 11
    
    ' render block
    Width = 317
    Height = 94
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(21), X, Y, 0, 0, Width, Height, Width, Height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_Default, sUser & chatShowLine, X + 74, Y + 2, White
    Else
        RenderText Font_Default, sUser, X + 74, Y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_Default, CensorWord(sPass) & chatShowLine, X + 74, Y + 26, White
    Else
        RenderText Font_Default, CensorWord(sPass), X + 74, Y + 26, White
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        X = GUIWindow(GUI_MAINMENU).X + Buttons(buttonnum).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(buttonnum).Y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(buttonnum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(buttonnum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawRegister()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X + 86
    Y = GUIWindow(GUI_MAINMENU).Y + 92
    buttonnum = 12
    
    ' render block
    Width = 319
    Height = 107
    'EngineRenderRectangle Tex_GUI(20), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(20), X, Y, 0, 0, Width, Height, Width, Height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_Default, sUser & chatShowLine, X + 74, Y + 2, White
    Else
        RenderText Font_Default, sUser, X + 74, Y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_Default, CensorWord(sPass) & chatShowLine, X + 74, Y + 26, White
    Else
        RenderText Font_Default, CensorWord(sPass), X + 74, Y + 26, White
    End If
    
    ' render password
    If curTextbox = 3 Then ' focuses
        RenderText Font_Default, CensorWord(sPass2) & chatShowLine, X + 74, Y + 50, White
    Else
        RenderText Font_Default, CensorWord(sPass2), X + 74, Y + 50, White
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        X = GUIWindow(GUI_MAINMENU).X + Buttons(buttonnum).X
        Y = GUIWindow(GUI_MAINMENU).Y + Buttons(buttonnum).Y
        Width = Buttons(buttonnum).Width
        Height = Buttons(buttonnum).Height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(buttonnum).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(buttonnum).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawCredits()
Dim X As Long, Y As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_MAINMENU).X + 187
    Y = GUIWindow(GUI_MAINMENU).Y + 86
    Width = 121
    Height = 120
    'engineRenderRectangle Tex_GUI(19), x, y, 0, 0, width, height, width, height, width, height
    Directx8.RenderTexture Tex_GUI(19), X, Y, 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, X As Long, Y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, Height As Long, Sprite As Long

    ' backwindow + empty bars
    X = GUIWindow(GUI_BARS).X
    Y = GUIWindow(GUI_BARS).Y
    Width = GUIWindow(GUI_BARS).Width
    Height = GUIWindow(GUI_BARS).Height
    Directx8.RenderTexture Tex_GUI(4), X, Y, 0, 0, Width, Height, Width, Height
    
    ' draw player sprite and face
    Sprite = GetPlayerSprite(MyIndex)
    If Sprite > 0 And Sprite <= Count_Face Then
        Directx8.RenderTexture Tex_Face(Sprite), X + 7, Y + 7, 0, 0, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight, gTexture(Tex_Face(Sprite)).RWidth, gTexture(Tex_Face(Sprite)).RHeight
    End If
            
    If Sprite > 0 And Sprite <= Count_Char Then
        Directx8.RenderTexture Tex_Char(Sprite), X + 29, Y + 19, 0, 0, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4, gTexture(Tex_Char(Sprite)).RWidth / 4, gTexture(Tex_Char(Sprite)).RHeight / 4
    End If
    
    ' render map name and colorized it based on map moral
    If Map.Moral = MAP_MORAL_NONE Then
        RenderText Font_Default, Trim$(Map.name), X + 5, Y + Height - 19, BrightRed
    ElseIf Map.Moral = MAP_MORAL_SAFE Then
        RenderText Font_Default, Trim$(Map.name), X + 5, Y + Height - 19, White
    ElseIf Map.Moral = MAP_MORAL_BOSS Then
        RenderText Font_Default, Trim$(Map.name), X + 5, Y + Height - 19, Pink
    End If
    
    ' render game time and date
    sString = Right(GameTime.Day, 1)
    If sString = 1 Then
        sString = GameTime.Day & "st"
    ElseIf sString = 2 Then
        sString = GameTime.Day & "nd"
    ElseIf sString = 3 Then
        sString = GameTime.Day & "rd"
    Else
        sString = GameTime.Day & "th"
    End If
    
    Call RenderText(Font_Default, " - " & KeepTwoDigit(GameTime.Hour) & ":" & KeepTwoDigit(GameTime.Minute) & "  " & sString & " " & MonthName(GameTime.Month) & " " & GameTime.Year, X + 5 + EngineGetTextWidth(Font_Default, Trim$(Map.name)), Y + Height - 19, White)
    
    ' hardcoded for POT textures
    barWidth = 186
    
    ' health bar
    BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / barWidth)) * barWidth
    Directx8.RenderTexture Tex_GUI(13), X + 62, Y + 9, 0, 0, BarWidth_GuiHP, gTexture(Tex_GUI(13)).Height, BarWidth_GuiHP, gTexture(Tex_GUI(13)).Height
    ' render health
    sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    dX = X + 62 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 9
    RenderText Font_Default, sString, dX, dY, White
    
    ' spirit bar
    BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / barWidth)) * barWidth
    Directx8.RenderTexture Tex_GUI(14), X + 62, Y + 31, 0, 0, BarWidth_GuiSP, gTexture(Tex_GUI(14)).Height, BarWidth_GuiSP, gTexture(Tex_GUI(14)).Height
    ' render spirit
    sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    dX = X + 62 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 31
    RenderText Font_Default, sString, dX, dY, White
    
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP_Max = barWidth
    End If
    Directx8.RenderTexture Tex_GUI(15), X + 62, Y + 53, 0, 0, BarWidth_GuiEXP, gTexture(Tex_GUI(15)).Height, BarWidth_GuiEXP, gTexture(Tex_GUI(15)).Height
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = GetPlayerExp(MyIndex) & "/" & TNL
    Else
        sString = "Max Level"
    End If
    dX = X + 62 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = Y + 53
    RenderText Font_Default, sString, dX, dY, White
End Sub

Public Sub DrawTrade()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_TRADE).Width
    Height = GUIWindow(GUI_TRADE).Height
    Directx8.RenderTexture Tex_GUI(34), GUIWindow(GUI_TRADE).X, GUIWindow(GUI_TRADE).Y, 0, 0, Width, Height, Width, Height
        For I = 1 To MAX_INV
            ' render your offer
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).Num)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                    Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(I).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                            
                        Amount = CStr(TradeYourOffer(I).Value)
                            
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
            
            ' draw their offer
            itemNum = TradeTheirOffer(I).Num
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
                ItemPic = Item(itemNum).Pic
                If ItemPic > 0 And ItemPic <= Count_Item Then
                
                    Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop - 2 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                    Left = GUIWindow(GUI_TRADE).X + 257 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                    Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(I).Value > 1 Then
                        Y = Top + 21
                        X = Left - 4
                                
                        Amount = CStr(TradeTheirOffer(I).Value)
                                
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If CLng(Amount) < 1000000 Then
                            Colour = White
                        ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                            Colour = Yellow
                        ElseIf CLng(Amount) > 10000000 Then
                            Colour = BrightGreen
                        End If
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, Colour
                    End If
                End If
            End If
        Next
        ' draw buttons
    For I = Button_TradeAccept To Button_TradeDecline
        ' set co-ordinate
        X = Buttons(I).X
        Y = Buttons(I).Y
        Width = Buttons(I).Width
        Height = Buttons(I).Height
        ' check for state
        If Buttons(I).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_c(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
        ElseIf (GlobalX >= X And GlobalX <= X + Buttons(I).Width) And (GlobalY >= Y And GlobalY <= Y + Buttons(I).Height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons_h(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' play sound if needed
            If Not lastButtonSound = I Then
                FMOD.Sound_Play Sound_ButtonHover
                lastButtonSound = I
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            Directx8.RenderTexture Tex_Buttons(Buttons(I).PicNum), X, Y, 0, 0, Width, Height, Width, Height
            ' reset sound if needed
            If lastButtonSound = I Then lastButtonSound = 0
        End If
    Next
    RenderText Font_Default, "Your worth: " & YourWorth, GUIWindow(GUI_TRADE).X + 21, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, "Their worth: " & TheirWorth, GUIWindow(GUI_TRADE).X + 250, GUIWindow(GUI_TRADE).Y + 299, White
    RenderText Font_Default, TradeStatus, (GUIWindow(GUI_TRADE).Width / 2) - (EngineGetTextWidth(Font_Default, TradeStatus) / 2), GUIWindow(GUI_TRADE).Y + 317, Yellow
    DrawTradeItemDesc
End Sub

Public Sub DrawTradeItemDesc()
Dim tradeNum As Long
    If Not GUIWindow(GUI_TRADE).visible Then Exit Sub
        
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum > 0 Then
        If GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num) > 0 Then
            DrawItemDesc GetPlayerInvItemNum(MyIndex, TradeYourOffer(tradeNum).Num), GUIWindow(GUI_TRADE).X + 480 + 10, GUIWindow(GUI_TRADE).Y
        End If
    End If
End Sub

Public Sub DrawFader()
    If faderAlpha < 0 Then faderAlpha = 0
    If faderAlpha > 254 Then faderAlpha = 254
    Directx8.RenderTexture Tex_White, 0, 0, 0, 0, 800, 600, 32, 32, D3DColorARGB(faderAlpha, 0, 0, 0)
End Sub

Public Sub DrawCurrency()
Dim X As Long, Y As Long, buttonnum As Long
Dim Width As Long, Height As Long

    X = GUIWindow(GUI_CURRENCY).X
    Y = GUIWindow(GUI_CURRENCY).Y
    ' render chatbox
    Width = GUIWindow(GUI_CURRENCY).Width
    Height = GUIWindow(GUI_CURRENCY).Height
    Directx8.RenderTexture Tex_GUI(24), X, Y, 0, 0, Width, Height, Width, Height
    Width = EngineGetTextWidth(Font_Default, CurrencyText)
    RenderText Font_Default, CurrencyText, X + 87 + (123 - (Width / 2)), Y + 40, White
    RenderText Font_Default, sDialogue & chatShowLine, X + 90, Y + 65, White
    
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    X = GUIWindow(GUI_CURRENCY).X + 155
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyAcceptState = 2 Then
        ' clicked
        RenderText Font_Default, "[Accept]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, "[Accept]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 1 Then
                FMOD.Sound_Play Sound_ButtonHover
                lastNpcChatsound = 1
            End If
        Else
            ' normal
            RenderText Font_Default, "[Accept]", X, Y, Green
            ' reset sound if needed
            If lastNpcChatsound = 1 Then lastNpcChatsound = 0
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    X = GUIWindow(GUI_CURRENCY).X + 218
    Y = GUIWindow(GUI_CURRENCY).Y + 96
    If CurrencyCloseState = 2 Then
        ' clicked
        RenderText Font_Default, "[Close]", X, Y, Grey
    Else
        If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
            ' hover
            RenderText Font_Default, "[Close]", X, Y, Cyan
            ' play sound if needed
            If Not lastNpcChatsound = 2 Then
                FMOD.Sound_Play Sound_ButtonHover
                lastNpcChatsound = 2
            End If
        Else
            ' normal
            RenderText Font_Default, "[Close]", X, Y, Yellow
            ' reset sound if needed
            If lastNpcChatsound = 2 Then lastNpcChatsound = 0
        End If
    End If
End Sub
Public Sub DrawDialogue()
Dim I As Long, X As Long, Y As Long, Sprite As Long, Width As Long
Dim Height As Long

    ' draw background
    X = GUIWindow(GUI_DIALOGUE).X
    Y = GUIWindow(GUI_DIALOGUE).Y
    
    ' render chatbox
    Width = GUIWindow(GUI_DIALOGUE).Width
    Height = GUIWindow(GUI_DIALOGUE).Height
    Directx8.RenderTexture Tex_GUI(27), X, Y, 0, 0, Width, Height, Width, Height
    
    ' Draw the text
    RenderText Font_Default, Dialogue_TitleCaption, X + (Width / 2) - (EngineGetTextWidth(Font_Default, Dialogue_TitleCaption) / 2), Y + 10, Green
    RenderText Font_Default, WordWrap(Dialogue_TextCaption, Width - 20), X + 10, Y + 30, White
    
    If Dialogue_ButtonVisible(1) Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
            If Dialogue_ButtonState(1) = 2 Then
                ' clicked
                RenderText Font_Default, "[Accept]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Accept]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 1 Then
                        FMOD.Sound_Play Sound_ButtonHover
                        lastNpcChatsound = 1
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Accept]", X, Y, Green
                    ' reset sound if needed
                    If lastNpcChatsound = 1 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(2) Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 105
            If Dialogue_ButtonState(2) = 2 Then
                ' clicked
                RenderText Font_Default, "[Okay]", X, Y, Grey
            Else
                If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                    ' hover
                    RenderText Font_Default, "[Okay]", X, Y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = 2 Then
                        FMOD.Sound_Play Sound_ButtonHover
                        lastNpcChatsound = 2
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[Okay]", X, Y, BrightRed
                    ' reset sound if needed
                    If lastNpcChatsound = 2 Then lastNpcChatsound = 0
                End If
            End If
    End If
    If Dialogue_ButtonVisible(3) Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        X = GUIWindow(GUI_DIALOGUE).X + ((GUIWindow(GUI_DIALOGUE).Width / 2) - (Width / 2))
        Y = GUIWindow(GUI_DIALOGUE).Y + 120
        If Dialogue_ButtonState(3) = 2 Then
            ' clicked
            RenderText Font_Default, "[Close]", X, Y, Grey
        Else
            If (GlobalX >= X And GlobalX <= X + Width) And (GlobalY >= Y And GlobalY <= Y + 14) Then
                ' hover
                RenderText Font_Default, "[Close]", X, Y, Cyan
                ' play sound if needed
                If Not lastNpcChatsound = 3 Then
                    FMOD.Sound_Play Sound_ButtonHover
                    lastNpcChatsound = 3
                End If
            Else
                ' normal
                RenderText Font_Default, "[Close]", X, Y, Yellow
                ' reset sound if needed
                If lastNpcChatsound = 3 Then lastNpcChatsound = 0
            End If
        End If
    End If
End Sub

Public Sub DrawBank()
Dim I As Long, X As Long, Y As Long, itemNum As Long, ItemPic As Long, Left As Long, Top As Long, Amount As Long, Colour As Long, Width As Long
Dim Height As Long

    Width = GUIWindow(GUI_BANK).Width
    Height = GUIWindow(GUI_BANK).Height
    
    Directx8.RenderTexture Tex_GUI(25), GUIWindow(GUI_BANK).X, GUIWindow(GUI_BANK).Y, 0, 0, Width, Height, Width, Height
    
    For I = 1 To MAX_BANK
        itemNum = GetBankItemNum(I)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= Count_Item Then
                Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                Directx8.RenderTexture Tex_Item(ItemPic), Left, Top, 0, 0, 32, 32, 32, 32
                       
                ' If the bank item is in a stack, draw the amount...
                If GetBankItemValue(I) > 1 Then
                    Y = Top + 22
                    X = Left - 4
                    Amount = CStr(GetBankItemValue(I))
                            
                    ' Draw the currency
                    If CLng(Amount) < 1000000 Then
                        Colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        Colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        Colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, Colour
                End If
            End If
        End If
    Next
    DrawBankItemDesc
End Sub
Public Sub DrawBankItemDesc()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum > 0 Then
        If GetBankItemNum(bankNum) > 0 Then
            DrawItemDesc GetBankItemNum(bankNum), GUIWindow(GUI_BANK).X + 480, GUIWindow(GUI_BANK).Y
        End If
    End If
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'load blood then
    BloodCount = gTexture(Tex_Blood).Width / 32
    
    With Blood(Index)
        If .Alpha <= 0 Then Exit Sub
        ' check if we should be seeing it
        If .timer + 20000 < timeGetTime Then
            .Alpha = .Alpha - 1
        End If
        
        rec.Top = 0
        rec.bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        Directx8.RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorARGB(.Alpha, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawBlood", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawProjectile()
Dim Angle As Long, X As Long, Y As Long, I As Long
    If LastProjectile > 0 Then
        
        ' ****** Create Particle ******
        For I = 1 To LastProjectile
            With ProjectileList(I)
                If .Graphic Then
                
                    ' ****** Update Position ******
                    Angle = DegreeToRadian * Engine_GetAngle(.X, .Y, .tx, .ty)
                    .X = .X + (Sin(Angle) * ElapsedTime * 0.3)
                    .Y = .Y - (Cos(Angle) * ElapsedTime * 0.3)
                    X = .X
                    Y = .Y
                    
                    ' ****** Update Rotation ******
                    If .RotateSpeed > 0 Then
                        .Rotate = .Rotate + (.RotateSpeed * ElapsedTime * 0.01)
                        Do While .Rotate > 360
                            .Rotate = .Rotate - 360
                        Loop
                    End If
                    
                    ' ****** Render Projectile ******
                    If .Rotate = 0 Then
                        Call Directx8.RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y)
                    Else
                        Call Directx8.RenderTexture(Tex_Projectile(.Graphic), ConvertMapX(X), ConvertMapY(Y), 0, 0, PIC_X, PIC_Y, PIC_X, PIC_Y, , .Rotate)
                    End If
                    
                End If
            End With
        Next
        
        ' ****** Erase Projectile ******    Seperate Loop For Erasing
        For I = 1 To LastProjectile
            If ProjectileList(I).Graphic Then
                If Abs(ProjectileList(I).X - ProjectileList(I).tx) < 20 Then
                    If Abs(ProjectileList(I).Y - ProjectileList(I).ty) < 20 Then
                        Call ClearProjectile(I)
                    End If
                End If
            End If
        Next
        
    End If
End Sub

Public Sub DrawTint()
Dim color As Long
    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    Directx8.RenderTexture Tex_White, 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, color
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
    
    If Index = 0 Then Exit Sub
    If Not Events(Index).Layer = Layer Then Exit Sub
    If Events(Index).Animated = YES Then
        If eventAnimTimer < timeGetTime Then
            ' animate events
            Select Case eventAnimFrame
                Case 0
                    eventAnimFrame = 1
                Case 1
                    eventAnimFrame = 2
                Case 2
                    eventAnimFrame = 0
            End Select
            eventAnimTimer = timeGetTime + 400
        End If
        Sprite = Events(Index).Graphic(eventAnimFrame)
    Else
        Sprite = Events(Index).Graphic(Player(MyIndex).EventGraphic(Index))
    End If
    If Sprite = 0 Then Exit Sub

    ' src rect
    With rec
        .Top = 0
        .bottom = gTexture(Tex_Event(Sprite)).RHeight
        .Left = 0
        .Right = gTexture(Tex_Event(Sprite)).RWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (X * PIC_X) - (gTexture(Tex_Event(Sprite)).RWidth / 2) + 16
    Y = (Y * PIC_Y) - gTexture(Tex_Event(Sprite)).RHeight + 32
    
    Width = rec.Right - rec.Left
    Height = rec.bottom - rec.Top
    Directx8.RenderTexture Tex_Event(Sprite), ConvertMapX(X), ConvertMapY(Y), 0, 0, Width, Height, Width, Height
End Sub

Public Sub DrawWeather()
Dim color As Long, I As Long, SpriteLeft As Long
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(I).Type - 1
            End If
            Directx8.RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(I).X), ConvertMapY(WeatherParticle(I).Y), SpriteLeft * 32, 0, 32, 32, 32, 32
        End If
    Next
End Sub

Public Sub DrawCursor()
Dim I As Long, X As Long, Y As Long, CursorNum As Long
    CursorNum = 1
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            X = (Player(I).X * 32) + Player(I).xOffset + 32
            Y = (Player(I).Y * 32) + Player(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    CursorNum = 2
                    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                        If Item(GetPlayerEquipment(MyIndex, Weapon)).Projectile > 0 Then CursorNum = 3
                    End If
                End If
            End If
        End If
    Next
    
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            X = (MapNpc(I).X * 32) + MapNpc(I).xOffset + 32
            Y = (MapNpc(I).Y * 32) + MapNpc(I).yOffset + 32
            If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                    CursorNum = 2
                    If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                        If Item(GetPlayerEquipment(MyIndex, Weapon)).Projectile > 0 Then CursorNum = 3
                    End If
                End If
            End If
        End If
    Next
    
    ' Resources
    If Count_Resource > 0 Then
        If Resources_Init Then
            If Resource_Index > 0 Then
                For I = 1 To Resource_Index
                    X = (MapResource(I).X * 32) + 32
                    Y = (MapResource(I).Y * 32) + 32
                    If X >= GlobalX_Map And X <= GlobalX_Map + 32 Then
                        If Y >= GlobalY_Map And Y <= GlobalY_Map + 32 Then
                            CursorNum = 4
                        End If
                    End If
                Next
            End If
        End If
    End If
    Directx8.RenderTexture Tex_Cursor(CursorNum), GlobalX, GlobalY, 0, 0, 32, 32, 32, 32
End Sub
