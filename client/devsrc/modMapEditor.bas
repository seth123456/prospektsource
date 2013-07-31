Attribute VB_Name = "modMapEditor"
Option Explicit

Public Map As MapRec

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_RESOURCE As Byte = 5
Public Const TILE_TYPE_NPCSPAWN As Byte = 6
Public Const TILE_TYPE_SHOP As Byte = 7
Public Const TILE_TYPE_BANK As Byte = 8
Public Const TILE_TYPE_HEAL As Byte = 9
Public Const TILE_TYPE_TRAP As Byte = 10
Public Const TILE_TYPE_SLIDE As Byte = 11
Public Const TILE_TYPE_EVENT As Byte = 12
Public Const TILE_TYPE_SOUND As Byte = 13
Public Const TILE_TYPE_THRESHOLD As Byte = 14

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    Roof
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    BossNpc As Long
    
    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte
    
    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    
    Panorama As Byte
    
    Weather As Long
    WeatherIntensity As Long
End Type

' /////////////////
' // Map Editor //
' /////////////////

Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = CELL_SIZE
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = CELL_SIZE
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = CELL_SIZE
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = CELL_SIZE
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = CELL_SIZE
    autoNE(1).Y = CELL_SIZE
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = CELL_SIZE
    
    ' SW - k
    autoNE(3).X = CELL_SIZE
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = CELL_SIZE
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = CELL_SIZE
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                calculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                cacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub cacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > Count_Tileset Then
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).Layer(layerNum).renderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * CELL_SIZE) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * CELL_SIZE) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub calculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).X <> Map.Tile(X2, Y2).Layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).Y <> Map.Tile(X2, Y2).Layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Sub UpdateCamera()
    If Map.MaxX > MAX_MAPX Then
        frmMain.scrlX.Enabled = True
        frmMain.scrlX.Max = Map.MaxX - MAX_MAPX
    Else
        frmMain.scrlX.Enabled = False
        frmMain.scrlX.Value = 0
    End If
    
    If Map.MaxY > MAX_MAPY Then
        frmMain.scrlY.Enabled = True
        frmMain.scrlY.Max = Map.MaxY - MAX_MAPY
    Else
        frmMain.scrlY.Enabled = False
        frmMain.scrlY.Value = 0
    End If
    
    TileView.Left = frmMain.scrlX.Value
    TileView.Right = frmMain.scrlX.Value + MAX_MAPX
    TileView.Top = frmMain.scrlY.Value
    TileView.bottom = frmMain.scrlY.Value + MAX_MAPY
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = X - (TileView.Left * CELL_SIZE)
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.Top * CELL_SIZE)
End Function

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ' horizontal scrolling
    If gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).RWidth < frmMain.picTileset.Width Then
        frmMain.scrlPictureX.Enabled = False
    Else
        frmMain.scrlPictureX.Enabled = True
        slcTilesetLeft = (frmMain.scrlPictureX.Value * CELL_SIZE)
    End If
    
    ' vertical scrolling
    If gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).RHeight < frmMain.picTileset.Height Then
        frmMain.scrlPictureY.Enabled = False
    Else
        frmMain.scrlPictureY.Enabled = True
        slcTilesetTop = (frmMain.scrlPictureY.Value * CELL_SIZE)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = (slcTilesetLeft \ CELL_SIZE) + (X \ CELL_SIZE)
        EditorTileY = (slcTilesetTop \ CELL_SIZE) + (Y \ CELL_SIZE)
        
        shpSelectedTop = (Y \ CELL_SIZE) * CELL_SIZE
        shpSelectedLeft = (X \ CELL_SIZE) * CELL_SIZE
        
        shpSelectedWidth = CELL_SIZE
        shpSelectedHeight = CELL_SIZE
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (slcTilesetLeft \ CELL_SIZE) + (X \ CELL_SIZE) + 1
        Y = (slcTilesetTop \ CELL_SIZE) + (Y \ CELL_SIZE) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).Width / CELL_SIZE Then X = gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).Width / CELL_SIZE
        If Y < 0 Then Y = 0
        If Y > gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).Height / CELL_SIZE Then Y = gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).Height / CELL_SIZE
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        shpSelectedWidth = EditorTileWidth * CELL_SIZE
        shpSelectedHeight = EditorTileHeight * CELL_SIZE
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If CurEditType = EDIT_MAP Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmMain.scrlAutotile.Value
            Else ' multi tile!
                If frmMain.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmMain.scrlAutotile.Value
                End If
            End If
        ElseIf CurEditType = EDIT_ATTRIBUTES Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmMain.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmMain.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = EditorWarpInstanced
                End If
                ' item spawn
                If frmMain.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc avoid
                If frmMain.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' resource
                If frmMain.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc spawn
                If frmMain.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' shop
                If frmMain.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' bank
                If frmMain.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' heal
                If frmMain.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' trap
                If frmMain.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' slide
                If frmMain.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' Event
                If frmMain.optEvent.Value Then
                    .Type = TILE_TYPE_EVENT
                    .Data1 = MapEditorEventIndex
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' sound
                If frmMain.optSound.Value Then
                    .Type = TILE_TYPE_SOUND
                    .Data1 = MapEditorSound
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' Threshold
                If frmMain.optThreshold.Value Then
                    .Type = TILE_TYPE_THRESHOLD
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
            End With
        ElseIf CurEditType = EDIT_DIRBLOCK Then
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For I = 1 To 4
                If X >= DirArrowX(I) And X <= DirArrowX(I) + 8 Then
                    If Y >= DirArrowY(I) And Y <= DirArrowY(I) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(I), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(I))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If CurEditType = EDIT_MAP Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                cacheRenderState X, Y, CurLayer
            End With
        ElseIf CurEditType = EDIT_ATTRIBUTES Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .Data4 = 0
            End With
        End If
    End If
    
    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If theAutotile > 0 Then
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmMain.scrlTileset.Value
            .Autotile(CurLayer) = theAutotile
            cacheRenderState X, Y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmMain.scrlTileset.Value
            .Autotile(CurLayer) = 0
            cacheRenderState X, Y, CurLayer
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + X2
                            .Layer(CurLayer).Y = EditorTileY + Y2
                            .Layer(CurLayer).Tileset = frmMain.scrlTileset.Value
                            .Autotile(CurLayer) = 0
                            cacheRenderState X, Y, CurLayer
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    frmMain.fraNpcSpawn.Visible = False
    frmMain.fraResource.Visible = False
    frmMain.fraMapItem.Visible = False
    frmMain.fraMapWarp.Visible = False
    frmMain.fraShop.Visible = False
    frmMain.fraEvent.Visible = False
    frmMain.fraSlide.Visible = False
    frmMain.fraTrap.Visible = False
    frmMain.fraHeal.Visible = False
    frmMain.fraSoundEffect.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next
    End If
    
    CacheResources

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim I As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = 0
                Map.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.Tile(X, Y).Layer(CurLayer).Tileset = 0
                cacheRenderState X, Y, CurLayer
            Next
        Next
        
        ' re-cache autos
        initAutotiles
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim I As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).Layer(CurLayer).Tileset = frmMain.scrlTileset.Value
                Map.Tile(X, Y).Autotile(CurLayer) = frmMain.scrlAutotile.Value
                cacheRenderState X, Y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For I = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(I)
    Next
    frmEditor_MapProperties.cmbSound.Clear
    frmEditor_MapProperties.cmbSound.AddItem "None."
    For I = 1 To UBound(soundCache)
        frmEditor_MapProperties.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .scrlBoss.Max = MAX_MAP_NPCS
        .txtName.Text = Trim$(Map.name)
        .ScrlFog.Max = Count_Fog
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For I = 0 To .lstMusic.ListCount
                If .lstMusic.List(I) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = I
                End If
            Next
        End If
        
        If .cmbSound.ListCount >= 0 Then
            .cmbSound.ListIndex = 0
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Map.BGS) Then
                    .cmbSound.ListIndex = I
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(Map.Up)
        .txtDown.Text = CStr(Map.Down)
        .txtLeft.Text = CStr(Map.Left)
        .txtRight.Text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.Text = CStr(Map.BootMap)
        .txtBootX.Text = CStr(Map.BootX)
        .txtBootY.Text = CStr(Map.BootY)
        .scrlBoss = Map.BossNpc
        .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
        
        .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.Alpha
        .cmbPanorama.ListIndex = Map.Panorama
        .CmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.Value = Map.WeatherIntensity

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.Npc(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(Npc(Map.Npc(X)).name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(Npc(X).name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & CurrentMap
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function
