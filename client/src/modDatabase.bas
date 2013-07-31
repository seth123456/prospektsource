Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext, ByVal erLine)
Dim FileName As String
    FileName = App.path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "' at line #" & erLine & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    If IgnoreHandler Then Exit Sub
    Select Case MsgBox("The following error occured at '" & procName & "' in '" & contName & "' at line #" & erLine & "." & vbNewLine & "Run-time error '" & erNumber & "': " & erDesc & ".", vbAbortRetryIgnore, Options.Game_Name)
        Case vbAbort: DestroyGame
        Case vbIgnore: IgnoreHandler = True
    End Select
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If LenB(Dir(FileName)) > 0 Then
        FileExist = True
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    FileName = App.path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", str(Options.savePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", str(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", str(Options.sound))
    Call PutVar(FileName, "Options", "Debug", str(Options.Debug))
    Call PutVar(FileName, "Options", "noAuto", str(Options.noAuto))
    Call PutVar(FileName, "Options", "render", str(Options.render))
    Call PutVar(FileName, "Options", "Fullscreen", str(Options.Fullscreen))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    FileName = App.path & "\Data Files\config.ini"
    If Not FileExist(FileName) Then
        Options.Game_Name = "Eclipse Reborn"
        Options.Password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.sound = 1
        Options.Debug = 0
        Options.noAuto = 0
        Options.render = 0
        Options.Fullscreen = 0
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.savePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
        Options.noAuto = GetVar(FileName, "Options", "noAuto")
        Options.render = GetVar(FileName, "Options", "render")
        Options.Fullscreen = GetVar(FileName, "Options", "Fullscreen")
    End If
    
    ' set the button states for options
    setOptionsState
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    FileName = App.path & MAP_PATH & "map" & MapNum & MAP_EXT

    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Map.name
    Put #F, , Map.Music
    Put #F, , Map.Revision
    Put #F, , Map.Moral
    Put #F, , Map.Up
    Put #F, , Map.Down
    Put #F, , Map.Left
    Put #F, , Map.Right
    Put #F, , Map.BootMap
    Put #F, , Map.BootX
    Put #F, , Map.BootY
    Put #F, , Map.MaxX
    Put #F, , Map.MaxY

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #F, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , Map.Npc(X)
    Next
    
    Put #F, , Map.BossNpc
    Put #F, , Map.Fog
    Put #F, , Map.FogSpeed
    Put #F, , Map.FogOpacity
    
    Put #F, , Map.Red
    Put #F, , Map.Green
    Put #F, , Map.Blue
    Put #F, , Map.Alpha
    
    Put #F, , Map.Panorama
    
    Put #F, , Map.Weather
    Put #F, , Map.WeatherIntensity
    Put #F, , Map.BGS
    Close #F
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    FileName = App.path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Map.name
        Get #F, , Map.Music
        Get #F, , Map.Revision
        Get #F, , Map.Moral
        Get #F, , Map.Up
        Get #F, , Map.Down
        Get #F, , Map.Left
        Get #F, , Map.Right
        Get #F, , Map.BootMap
        Get #F, , Map.BootX
        Get #F, , Map.BootY
        Get #F, , Map.MaxX
        Get #F, , Map.MaxY
        ' have to set the tile()
        ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Get #F, , Map.Tile(X, Y)
            Next
        Next
    
        For X = 1 To MAX_MAP_NPCS
            Get #F, , Map.Npc(X)
            MapNpc(X).Num = Map.Npc(X)
        Next
        
        Get #F, , Map.BossNpc
        Get #F, , Map.Fog
        Get #F, , Map.FogSpeed
        Get #F, , Map.FogOpacity
        
        Get #F, , Map.Red
        Get #F, , Map.Green
        Get #F, , Map.Blue
        Get #F, , Map.Alpha
        
        Get #F, , Map.Panorama
        
        Get #F, , Map.Weather
        Get #F, , Map.WeatherIntensity
        Get #F, , Map.BGS
    Close #F
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).name = vbNullString
    Animation(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ANIMATIONS
        Call ClearAnimation(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).name = vbNullString
    Npc(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_NPCS
        Call ClearNPC(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_RESOURCES
        Call ClearResource(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).name = name
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).Num
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Num = itemNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).Value
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Value = ItemValue
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Public Sub ClearEvents()
    Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    For I = 1 To MAX_EVENTS
        Call ClearEvent(I)
    Next I
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(Events(Index)), LenB(Events(Index)))
    Events(Index).name = vbNullString
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
