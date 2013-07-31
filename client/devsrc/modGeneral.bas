Attribute VB_Name = "modGeneral"
Option Explicit

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()
    'Set the high-resolution timer
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    Load frmMain
    frmMain.Show
    
    ' load options as first
    Call SetStatus("Loading options...")
    LoadOptions
    
    frmMain.txtLUser = Trim$(Options.Username)
    If Options.savePass = 1 Then
        frmMain.txtLPass = Trim$(Options.Password)
        frmMain.chkPass.Value = Options.savePass
    End If
    frmMain.txtIP = Trim$(Options.IP)
    frmMain.txtPort = Options.Port
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' load dx8
    Call SetStatus("Initializing DirectX8...")
    Set Directx8 = New clsDirectx8
    Directx8.Init
    
    ' cache sounds and music
    Call SetStatus("Caching sounds and music...")
    PopulateLists
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Loading done")
    InSuite = True
    MainLoop
    
End Sub

Public Sub DestroySuite()
Dim frm As Form
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' break out of MainLoop
    InSuite = False
    
    Call DestroyTCP
    
    ' unload dx8
    Directx8.Destroy
    
    ' unload all forms
    For Each frm In VB.Forms
        Unload frm
    Next
    
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroySuite", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub MainLoop()
Dim FrameTime As Long, Tick As Long
Dim renderspeed As Long
    ' *** Start GameLoop ***
    Do While InSuite
        Tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.
        
        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderspeed < Tick Then
            Call Render_Graphics
            renderspeed = timeGetTime + 15
        End If
        
        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < Tick + 20
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
    Loop
End Sub

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext, ByVal erLine)
Dim FileName As String

    FileName = App.path & "\data files\logs\dev-errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "' at line #" & erLine & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub SetStatus(ByVal Message As String)
    frmMain.stMain.Panels(1).Text = Message
End Sub

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
'Grabs the time from the 64-bit system timer and returns it in CELL_SIZE-bit
'after calculating it with the offset - allows us to have the
'"no roll-over" advantage of 64-bit timers with the RAM usage of CELL_SIZE-bit
'though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency

    'Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a CELL_SIZE-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset

End Function

Public Function FileExist(ByVal FileName As String) As Boolean
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If LenB(Dir(FileName)) > 0 Then
        FileExist = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub PopulateLists()
Dim strLoad As String, I As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = Dir(App.path & MUSIC_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To I) As String
        musicCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Cache sound list
    strLoad = Dir(App.path & SOUND_PATH & "*.*")
    I = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To I) As String
        soundCache(I) = strLoad
        strLoad = Dir
        I = I + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    FileName = App.path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", str(Options.savePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", str(Options.Port))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    FileName = App.path & "\Data Files\config.ini"
    If Not FileExist(FileName) Then
        Options.Password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        SaveOptions
    Else
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.savePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStartData()
Dim I As Long

    frmMain.scrlTileset.Max = Count_Tileset
    CurLayer = MapLayer.Ground
    
    frmMain.cmbSoundEffect.Clear
    For I = 1 To UBound(soundCache)
        frmMain.cmbSoundEffect.AddItem (soundCache(I))
    Next
    frmMain.cmbSoundEffect.ListIndex = 0
    
    frmMain.scrlPictureY.Max = (gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).RHeight - frmMain.picTileset.Height) / 32
    frmMain.scrlPictureX.Max = (gTexture(Tex_Tileset(frmMain.scrlTileset.Value)).RWidth - frmMain.picTileset.Width) / 32
    MapEditorTileScroll
    
    ' save options
    Options.savePass = frmMain.chkPass.Value
    Options.Username = Trim$(frmMain.txtLUser.Text)

    If frmMain.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMain.txtLPass.Text)
    End If
    
    SaveOptions
End Sub
