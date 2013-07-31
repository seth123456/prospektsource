Attribute VB_Name = "modGeneral"
Option Explicit
'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
    Dim i As Long
    Dim f As Long
    Dim time1 As Long
    Dim time2 As Long
    
    'Set the high-resolution timer
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    ' cache packet pointers
    Call InitMessages
    
    ' time the load
    time1 = timeGetTime
    frmServer.Show
    
    ' Initialize the random-number generator
    Randomize ', seed

    Call SetStatus("Checking folders...")
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "events"
    ChkDir App.Path & "\Data\", "effects"
    
    Call SetStatus("Loading options...")
    LoadOptions
    Call SetStatus("Loading time engine...")
    LoadTime
    Call SetStatus("Preparing listening socket...")
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    If Options.HighIndexing = 0 Then
        ' highindexing turned off
        Player_HighIndex = MAX_PLAYERS
    End If
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Loading swear filter...")
    Call LoadSwearFilter
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    ' create the pathfind matrixes
    Call SetStatus("Creating pathfinding matrixes...")
    For i = 1 To MAX_CACHED_MAPS
        CreatePathMatrix i
    Next
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    
    Call SetStatus("Starting listening...")
    ' Start listening
    frmServer.Socket(0).Listen
    
    Call SetStatus("Updating options...")
    Call UpdateCaption
    time2 = timeGetTime
    
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Saving time values...")
    Call SaveTime
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing instanced map data...")
    Call ClearInstancedMaps
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing events...")
    Call ClearEvents
    Call SetStatus("Clearing effects...")
    Call ClearEffects
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading events...")
    Call LoadEvents
    Call SetStatus("Loading effects...")
    Call LoadEffects
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

Public Sub InitTimeGetTime()
'*****************************************************************
'Gets the offset time for the timer so we can start at 0 instead of
'the returned system time, allowing us to not have a time roll-over until
'the program is running for 25 days
'*****************************************************************

    'Get the initial time
    GetSystemTime GetSystemTimeOffset

End Sub

Public Function timeGetTime() As Long
'*****************************************************************
'Grabs the time from the 64-bit system timer and returns it in 32-bit
'after calculating it with the offset - allows us to have the
'"no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
'though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency

    'Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset

End Function
