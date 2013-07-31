Attribute VB_Name = "modPathfinding"
Option Explicit

Public mapMatrix(1 To MAX_CACHED_MAPS) As mapMatrixRec
Public Type mapMatrixRec
    created As Boolean
    gaeGrid() As eCell
End Type

Public Enum eCell
    Void = 0&
    Start = 1&
    Obstacle = 2&
    target = 3&
End Enum

Private Type tCell
    x As Long               'Coordinates of the listed cell
    y As Long
    Parent As Long          'Parent Index within the list (-1 for start point)
    Cost As Single          'Cost to get til here
    Heuristic As Single     'Estimated cost til target
    Closed As Boolean       'Not considered anymore
End Type

Private Type tGrid
    ListStat As eListStat   'Status of the list element
    index As Long           'Index into the open list.
End Type

Private Enum eListStat
    Unprocessed = 0&
    IsOpen = 1&
    IsClosed = 2&
End Enum

Public Type tPoint
    x As Long
    y As Long
End Type

Public Function APlus(MapNum As Long, SX As Long, SY As Long, TX As Long, TY As Long, FreeCell As eCell, Path() As tPoint) As Boolean
    'A+ Pathfinding Algorithm:
    'Implementation by Herbert Glarner (herbert.glarner@bluewin.ch)
    'Unlimited use for whatever purpose allowed provided that above credits are given.
    'Suggestions and bug reports welcome.
    Dim lMaxList As Long
    Dim lActList As Long
    Dim sCheapCost As Single, lCheapIndex As Long
    Dim sTotalCost As Single
    Dim lCheapX As Long, lCheapY As Long
    Dim lOffX As Long, lOffY As Long
    Dim lTestX As Long, lTestY As Long
    Dim lMaxX As Long, lMaxY As Long
    Dim sAdditCost As Single
    Dim lPathPtr As Long
    
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure.
    'The two fields of tGrid can also be merged into the source matrix.
    '   Dim abGridCopy() As tGrid
    
    Const cSqr2 As Single = 1.4142135623731
    
    'Define the upper boundaries of the grid.
    lMaxX = UBound(mapMatrix(MapNum).gaeGrid, 1): lMaxY = UBound(mapMatrix(MapNum).gaeGrid, 2)
    
    'For each cell of the grid a bit is defined to hold it's "closed" status
    'and the index to the Open-List.
    'The test program wants to access this grid. For this reason it is defined
    'and initialized globally. Usually one would define and initialize it only
    'in this procedure. (Don't omit here: we need an empty matrix.)
    ReDim abGridCopy(0 To lMaxX, 0 To lMaxY) As tGrid
    
    'The starting point is added to the working list. It has no parent (-1).
    'The cost to get here is 0 (we start here). The direct distance enters
    'the Heuristic.
    ReDim grList(0 To 0) As tCell
    With grList(0)
        .x = SX: .y = SY: .Parent = -1: .Cost = 0
        .Heuristic = Sqr((TX - SX) * (TX - SX) + (TY - SY) * (TY - SY))
    End With
    
    'Start the algorithm
    Do
        'Get the cell with the lowest Cost+Heuristic. Initialize the cheapest cost
        'with an impossible high value (change as needed). The best found index
        'is set to -1 to indicate "none found".
        sCheapCost = 10000000
        lCheapIndex = -1
        'Check all cells of the list. Initially, there is only the start point,
        'but more will be added soon.
        For lActList = 0 To lMaxList
            'Only check if not closed already.
            If Not grList(lActList).Closed Then
                'If this cells total cost (Cost+Heuristic) is lower than the so
                'far lowest cost, then store this total cost and the cell's index
                'as the so far best found.
                sTotalCost = grList(lActList).Cost + grList(lActList).Heuristic
                If sTotalCost < sCheapCost Then
                    'New cheapest cost found.
                    sCheapCost = sTotalCost: lCheapIndex = lActList
                End If
            End If
        Next lActList
        
        'lCheapIndex contains the cell with the lowest total cost now.
        'If no such cell could be found, all cells were already closed and there
        'is no path at all to the target.
        If lCheapIndex = -1 Then
            'There is no path.
            APlus = False: Exit Function
        End If
        
        'Get the cheapest cell's coordinates
        lCheapX = grList(lCheapIndex).x
        lCheapY = grList(lCheapIndex).y
        
        'If the best field is the target field, we have found our path.
        If lCheapX = TX And lCheapY = TY Then
            'Path found.
            Exit Do
        End If
        
        'Check all immediate neighbors
        For lOffY = -1 To 1
            For lOffX = -1 To 1
                'Ignore the actual field, process all others (8 neighbors).
                If lOffX <> 0 Or lOffY <> 0 Then
                    ' ignore all diagonal movement
                    If Not (lOffX <> 0 And lOffY <> 0) Then
                        'Get the neighbor's coordinates.
                        lTestX = lCheapX + lOffX: lTestY = lCheapY + lOffY
                        'Don't test beyond the grid's boundaries.
                        If lTestX >= 0 And lTestX <= lMaxX And lTestY >= 0 And lTestY <= lMaxY Then
                            'The cell is within the grid's boundaries.
                            'Make sure the field is accessible. To be accessible,
                            'the cell must have the value as per the function
                            'argument FreeCell (change as needed). Of course, the
                            'target is allowed as well.
                            If mapMatrix(MapNum).gaeGrid(lTestX, lTestY) = FreeCell Or mapMatrix(MapNum).gaeGrid(lTestX, lTestY) = target Then
                                'The cell is accessible.
                                'For this we created the "bitmatrix" abGridCopy().
                                If abGridCopy(lTestX, lTestY).ListStat = Unprocessed Then
                                    'Register the new cell in the list.
                                    lMaxList = lMaxList + 1
                                    ReDim Preserve grList(0 To lMaxList) As tCell
                                    With grList(lMaxList)
                                        'The parent is where we come from (the cheapest field);
                                        'it's index is registered.
                                        .x = lTestX: .y = lTestY: .Parent = lCheapIndex
                                        'Additional cost is 1 for othogonal movement, cSqr2 for
                                        'diagonal movement (change if diagonal steps should have
                                        'a different cost).
                                        If Abs(lOffX) + Abs(lOffY) = 1 Then sAdditCost = 1# Else sAdditCost = cSqr2
                                        'Store cost to get there by summing the actual cell's cost
                                        'and the additional cost.
                                        .Cost = grList(lCheapIndex).Cost + sAdditCost
                                        'Calculate distance to target as the heuristical part
                                        .Heuristic = Sqr((TX - lTestX) * (TX - lTestX) + (TY - lTestY) * (TY - lTestY))
                                    End With
                                    'Register in the Grid copy as open.
                                    abGridCopy(lTestX, lTestY).ListStat = IsOpen
                                    'Also register the index to quickly find the element in the
                                    '"closed" list.
                                    abGridCopy(lTestX, lTestY).index = lMaxList
                                ElseIf abGridCopy(lTestX, lTestY).ListStat = IsOpen Then
                                    'Is the cost to get to this already open field cheaper when using
                                    'this path via lTestX/lTestY ?
                                    lActList = abGridCopy(lTestX, lTestY).index
                                    sAdditCost = IIf(Abs(lOffX) + Abs(lOffY) = 1, 1#, cSqr2)
                                    If grList(lCheapIndex).Cost + sAdditCost < grList(lActList).Cost Then
                                        'The cost to reach the already open field is lower via the
                                        'actual field.
                                        
                                        'Store new cost
                                        grList(lActList).Cost = grList(lCheapIndex).Cost + sAdditCost
                                        'Store new parent
                                        grList(lActList).Parent = lCheapIndex
                                    End If
                                'ElseIf abGridCopy(lTestX, lTestY) = IsClosed Then
                                '   'This cell can be ignored
                                End If
                            End If
                        End If
                    End If
                End If
            Next lOffX
        Next lOffY
        'Close the just checked cheapest cell.
        grList(lCheapIndex).Closed = True
        abGridCopy(lCheapX, lCheapY).ListStat = IsClosed
    Loop

    'The path can be found by backtracing from the field TX/TY until SX/SY.
    'The path is traversed in backwards order and stored reversely (!) in
    'the "argument" Path().
    ReDim Path(0 To 0) As tPoint
    lPathPtr = -1
    'lCheapIndex (lCheapX/Y) initially contains the target TX/TY
    Do
        'Store the coordinates of the current cell
        lPathPtr = lPathPtr + 1
        ReDim Preserve Path(0 To lPathPtr) As tPoint
        Path(lPathPtr).x = grList(lCheapIndex).x
        Path(lPathPtr).y = grList(lCheapIndex).y
        'Follow the parent
        lCheapIndex = grList(lCheapIndex).Parent
    Loop While lCheapIndex <> -1
    
    APlus = True: Exit Function
End Function

Public Sub CreatePathMatrix(ByVal MapNum As Long)
Dim x As Long, y As Long

    ReDim mapMatrix(MapNum).gaeGrid(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY) As eCell
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE And Map(MapNum).Tile(x, y).Type <> TILE_TYPE_ITEM And Map(MapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
                mapMatrix(MapNum).gaeGrid(x, y) = Obstacle
            Else
                mapMatrix(MapNum).gaeGrid(x, y) = Void
            End If
        Next
    Next
    
    mapMatrix(MapNum).created = True
    Exit Sub
    
errorHandler:
    mapMatrix(MapNum).created = False
End Sub

Public Sub NpcMoveAlongPath(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim x As Long, y As Long
    With MapNpc(MapNum).Npc(MapNpcNum)
        ' make sure we're not at end of path
        If .pathLoc >= 1 Then
            x = .arPath(.pathLoc - 1).x
            y = .arPath(.pathLoc - 1).y
            ' up
            If y < .y Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_UP) Then
                    Call NpcMove(MapNum, MapNpcNum, DIR_UP, MOVING_WALKING)
                    .pathLoc = .pathLoc - 1
                    Exit Sub
                End If
            End If
            ' down
            If y > .y Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_DOWN) Then
                    Call NpcMove(MapNum, MapNpcNum, DIR_DOWN, MOVING_WALKING)
                    .pathLoc = .pathLoc - 1
                    Exit Sub
                End If
            End If
            ' left
            If x < .x Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_LEFT) Then
                    Call NpcMove(MapNum, MapNpcNum, DIR_LEFT, MOVING_WALKING)
                    .pathLoc = .pathLoc - 1
                    Exit Sub
                End If
            End If
            ' right
            If x > .x Then
                If CanNpcMove(MapNum, MapNpcNum, DIR_RIGHT) Then
                    Call NpcMove(MapNum, MapNpcNum, DIR_RIGHT, MOVING_WALKING)
                    .pathLoc = .pathLoc - 1
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

