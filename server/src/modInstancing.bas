Attribute VB_Name = "modInstancing"
Option Explicit

Private Type InstancedMap
    OriginalMap As Long
End Type

Public Const INSTANCED_MAP_SUFFIX As String = " (Instanced)"

Public InstancedMaps(1 To MAX_INSTANCED_MAPS) As InstancedMap

Public Sub ClearInstancedMaps()
    Dim i As Long
    For i = 1 To MAX_INSTANCED_MAPS
        CacheResources i + MAX_MAPS
        InstancedMaps(i).OriginalMap = 0
    Next i
End Sub

Public Function FindFreeInstanceMapSlot() As Long
    Dim i As Long
    For i = 1 To MAX_INSTANCED_MAPS
        If InstancedMaps(i).OriginalMap = 0 Then
            FindFreeInstanceMapSlot = i
            Exit Function
        End If
    Next i
    FindFreeInstanceMapSlot = -1
End Function
Public Function CreateInstance(ByVal MapNum As Long) As Long
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        CreateInstance = -1
        Exit Function
    End If
    
    Dim Slot As Long
    Slot = FindFreeInstanceMapSlot
    If Slot = -1 Then
        CreateInstance = -1
        Exit Function
    End If
    
    InstancedMaps(Slot).OriginalMap = MapNum
    Dim MemSize As Long
    
    'Copy Map Data
    MemSize = LenB(Map(MapNum))
    CopyMemory ByVal VarPtr(Map(Slot + MAX_MAPS)), ByVal VarPtr(Map(MapNum)), MemSize
    
    'Copy Map Item Data
    Dim i As Long
    For i = 1 To MAX_MAP_ITEMS
        MemSize = LenB(MapItem(MapNum, i))
        CopyMemory ByVal VarPtr(MapItem(Slot + MAX_MAPS, i)), ByVal VarPtr(MapItem(MapNum, i)), MemSize
    Next i
    
    'Copy Map NPCs
    MemSize = LenB(MapNpc(MapNum))
    CopyMemory ByVal VarPtr(MapNpc(Slot + MAX_MAPS)), ByVal VarPtr(MapNpc(MapNum)), MemSize
    
    'Re-Cache Resource
    Call CacheResources(Slot + MAX_MAPS)
    
    Call MapCache_Create(Slot + MAX_MAPS)
    
    If Not (Map(Slot + MAX_MAPS).Name = vbNullString) Then Map(Slot + MAX_MAPS).Name = Map(Slot + MAX_MAPS).Name & INSTANCED_MAP_SUFFIX
    CreateInstance = Slot
    Exit Function
End Function

Public Sub DestroyInstancedMap(ByVal Slot As Long)
    Call ClearMap(Slot + MAX_MAPS)
    Dim x As Long
    For x = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(x, Slot + MAX_MAPS)
    Next
    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x, Slot + MAX_MAPS)
    Next
    InstancedMaps(Slot).OriginalMap = 0
End Sub

Public Function IsInstancedMap(ByVal MapNum As Long) As Boolean
    IsInstancedMap = MapNum > MAX_MAPS And MapNum <= MAX_CACHED_MAPS
End Function

Sub InstancedWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    MapNum = CreateInstance(MapNum)
    If MapNum = -1 Then
        'Couldn't create instanced map!
        MapNum = GetPlayerMap(index)
    Else
        MapNum = MapNum + MAX_MAPS
    End If
    
    If TempPlayer(index).inParty Then
        If Party(TempPlayer(index).inParty).Leader Then
            For i = 1 To Party(TempPlayer(index).inParty).MemberCount
                Call PlayerWarp(Party(TempPlayer(index).inParty).Member(i), MapNum, x, y)
            Next
        End If
    Else
        Call PlayerWarp(index, MapNum, x, y)
    End If
End Sub


