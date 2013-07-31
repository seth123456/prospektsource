Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim FileName As String
    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim f As Long

    If Options.Logs = 1 Then
        FileName = App.Path & "\data\logs\" & FN

        If Not FileExist(FileName, True) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If

        f = FreeFile
        Open FileName For Append As #f
        Print #f, Time & ": " & Text
        Close #f
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    Dim FileName As String
    FileName = App.Path & "\data\config.ini"
    
    PutVar FileName, "OPTIONS", "Game_Name", Options.Game_Name
    PutVar FileName, "OPTIONS", "Port", CStr(Options.Port)
    PutVar FileName, "OPTIONS", "MOTD", Options.MOTD
    PutVar FileName, "OPTIONS", "Logs", CStr(Options.Logs)
    PutVar FileName, "OPTIONS", "HighIndexing", STR(Options.HighIndexing)
    PutVar FileName, "OPTIONS", "StartMap", CStr(Options.StartMap)
    PutVar FileName, "OPTIONS", "StartX", CStr(Options.StartMap)
    PutVar FileName, "OPTIONS", "StartY", CStr(Options.StartMap)
    
    PutVar FileName, "MAX", "MAX_PLAYERS", CStr(MAX_PLAYERS)
    PutVar FileName, "MAX", "MAX_LEVELS", CStr(MAX_LEVELS)
    PutVar FileName, "MAX", "MAX_PARTYS", CStr(MAX_PARTYS)
    PutVar FileName, "MAX", "MAX_MAPS", CStr(MAX_MAPS)
    
    PutVar FileName, "MAX", "MAX_ITEMS", CStr(MAX_ITEMS)
    PutVar FileName, "MAX", "MAX_NPCS", CStr(MAX_NPCS)
    PutVar FileName, "MAX", "MAX_ANIMATIONS", CStr(MAX_ANIMATIONS)
    PutVar FileName, "MAX", "MAX_SHOPS", CStr(MAX_SHOPS)
    PutVar FileName, "MAX", "MAX_RESOURCES", CStr(MAX_RESOURCES)
    PutVar FileName, "MAX", "MAX_EFFECTS", CStr(MAX_EFFECTS)
End Sub

Public Sub LoadOptions()
    Dim FileName As String
    FileName = App.Path & "\data\config.ini"
    
    ' load options, set if they dont exist
    If Not FileExist(FileName, True) Then
        Options.Game_Name = "Prospekt Source"
        Options.Port = 8000
        Options.MOTD = "Welcome to Prospekt Source."
        Options.Logs = 1
        Options.HighIndexing = 1
        Options.StartMap = 1
        Options.StartX = 1
        Options.StartY = 1
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "OPTIONS", "Game_Name")
        Options.Port = GetVar(FileName, "OPTIONS", "Port")
        Options.MOTD = GetVar(FileName, "OPTIONS", "MOTD")
        Options.Logs = GetVar(FileName, "OPTIONS", "Logs")
        Options.HighIndexing = GetVar(FileName, "OPTIONS", "Highindexing")
        Options.StartMap = GetVar(FileName, "OPTIONS", "StartMap")
        Options.StartX = GetVar(FileName, "OPTIONS", "StartX")
        Options.StartY = GetVar(FileName, "OPTIONS", "StartY")
    End If
End Sub

Public Sub ToggleMute(ByVal index As Long)
    ' exit out for rte9
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub

    ' toggle the player's mute
    If Player(index).isMuted = 1 Then
        Player(index).isMuted = 0
        ' Let them know
        PlayerMsg index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(index) & " has been unmuted."
    Else
        Player(index).isMuted = 1
        ' Let them know
        PlayerMsg index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(index) & " has been muted."
    End If
    
    ' save the player
    SavePlayer index
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
Dim FileName As String, IP As String, f As Long, i As Long

    ' Add banned to the player's index
    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    FileName = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(FileName, True) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    f = FreeFile
    
    Open FileName For Append As #f
        Print #f, IP
    Close #f
    
    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " has been banned.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned.")
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, f As Long
    
    FileName = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(FileName, True) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    f = FreeFile
    Open FileName For Input As #f

    Do While Not EOF(f)
        Input #f, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #f
            Exit Function
        End If
    Loop

    Close #f
End Function

Public Function isBanned_Account(ByVal index As Long) As Boolean
    If Player(index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String
    FileName = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String * PASS_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH + 1, RightPassword
        Close #nFileNum
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer index
    
    Player(index).Login = Name
    Player(index).Password = MD5_string(Password)

    Call SavePlayer(index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean
    If LenB(Trim$(Player(index).Name)) > 0 Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum
        
        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).Dir = DIR_DOWN
        Player(index).Map = Options.StartMap
        Player(index).x = Options.StartX
        Player(index).y = Options.StartY
        Player(index).Dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Spell(n) = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).sType = 2 ' spells
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal index As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\data\accounts\" & Trim$(Player(index).Login) & ".bin"
    
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , Player(index)
    Close #f
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    Dim FileName As String
    Dim f As Long
    Call ClearPlayer(index)
    FileName = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Player(index)
    Close #f
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Login = vbNullString
    Player(index).Password = vbNullString
    Player(index).Name = vbNullString
    Player(index).Class = 1
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim FileName As String
    Dim File As String
    FileName = App.Path & "\data\classes.ini"
    MAX_CLASSES = 2

    If Not FileExist(FileName, True) Then
        File = FreeFile
        Open FileName For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & MAX_CLASSES
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To MAX_CLASSES)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To MAX_CLASSES)
    End If

    Call ClearClasses

    For i = 1 To MAX_CLASSES
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(FileName, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(FileName, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(FileName, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(FileName, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(FileName, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(FileName, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(FileName, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(FileName, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_PLAYER_SPELLS Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(FileName, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long
    Dim x As Long
    
    FileName = App.Path & "\data\classes.ini"

    For i = 1 To MAX_CLASSES
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Maleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(FileName, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(FileName, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(FileName, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(FileName, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(FileName, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(FileName, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemnum As Long)
    Dim FileName As String
    Dim f  As Long
    FileName = App.Path & "\data\items\item" & itemnum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Item(itemnum)
    Close #f
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
        Close #f
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\shops\shop" & shopNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Shop(shopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(i)
        Close #f
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellnum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\spells\spells" & spellnum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Spell(spellnum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NPCNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\npcs\npc" & NPCNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Npc(NPCNum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Npc(i)
        Close #f
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        FileName = App.Path & "\data\resources\resource" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Resource(i)
        Close #f
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
End Sub

Sub LoadAnimations()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        FileName = App.Path & "\data\animations\animation" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Animation(i)
        Close #f
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
    Put #f, , Map(MapNum).Name
    Put #f, , Map(MapNum).Music
    Put #f, , Map(MapNum).Revision
    Put #f, , Map(MapNum).Moral
    Put #f, , Map(MapNum).Up
    Put #f, , Map(MapNum).Down
    Put #f, , Map(MapNum).Left
    Put #f, , Map(MapNum).Right
    Put #f, , Map(MapNum).BootMap
    Put #f, , Map(MapNum).BootX
    Put #f, , Map(MapNum).BootY
    Put #f, , Map(MapNum).MaxX
    Put #f, , Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #f, , Map(MapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map(MapNum).Npc(x)
    Next
    
    Put #f, , Map(MapNum).BossNpc
    Put #f, , Map(MapNum).Fog
    Put #f, , Map(MapNum).FogSpeed
    Put #f, , Map(MapNum).FogOpacity
    
    Put #f, , Map(MapNum).Red
    Put #f, , Map(MapNum).Green
    Put #f, , Map(MapNum).Blue
    Put #f, , Map(MapNum).Alpha
    
    Put #f, , Map(MapNum).Panorama
    
    Put #f, , Map(MapNum).Weather
    Put #f, , Map(MapNum).WeatherIntensity
    Put #f, , Map(MapNum).BGS
    Close #f
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Dim x As Long
    Dim y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\data\maps\map" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Map(i).Name
        Get #f, , Map(i).Music
        Get #f, , Map(i).Revision
        Get #f, , Map(i).Moral
        Get #f, , Map(i).Up
        Get #f, , Map(i).Down
        Get #f, , Map(i).Left
        Get #f, , Map(i).Right
        Get #f, , Map(i).BootMap
        Get #f, , Map(i).BootX
        Get #f, , Map(i).BootY
        Get #f, , Map(i).MaxX
        Get #f, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #f, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #f, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next
        
        Get #f, , Map(i).BossNpc
        Get #f, , Map(i).Fog
        Get #f, , Map(i).FogSpeed
        Get #f, , Map(i).FogOpacity
        
        Get #f, , Map(i).Red
        Get #f, , Map(i).Green
        Get #f, , Map(i).Blue
        Get #f, , Map(i).Alpha
        
        Get #f, , Map(i).Panorama
        
        Get #f, , Map(i).Weather
        Get #f, , Map(i).WeatherIntensity
        Get #f, , Map(i).BGS
        Close #f

        CacheResources i
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, index)), LenB(MapItem(MapNum, index)))
    MapItem(MapNum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_CACHED_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(index)), LenB(MapNpc(MapNum).Npc(index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_CACHED_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_CACHED_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub SaveBank(ByVal index As Long)
    Dim FileName As String
    Dim f As Long
    
    FileName = App.Path & "\data\banks\" & Trim$(Player(index).Login) & ".bin"
    
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Bank(index)
    Close #f
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal Name As String)
    Dim FileName As String
    Dim f As Long

    Call ClearBank(index)

    FileName = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(FileName, True) Then
        Call SaveBank(index)
        Exit Sub
    End If

    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Bank(index)
    Close #f

End Sub

Sub ClearBank(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
End Sub

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub

Public Sub ClearEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call ClearEvent(i)
    Next i
End Sub

Public Sub ClearEvent(ByVal index As Long)
    If index <= 0 Or index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(Events(index)), LenB(Events(index)))
    Events(index).Name = vbNullString
End Sub

Public Sub LoadEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call LoadEvent(i)
    Next i
End Sub

Public Sub LoadEvent(ByVal index As Long)
    On Error GoTo Errorhandle
    
    Dim f As Long, SCount As Long, s As Long, DCount As Long, d As Long
    Dim FileName As String
    FileName = App.Path & "\data\events\event" & index & ".dat"
    If FileExist(FileName, True) Then
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Events(index).Name
            Get #f, , Events(index).chkSwitch
            Get #f, , Events(index).chkVariable
            Get #f, , Events(index).chkHasItem
            Get #f, , Events(index).SwitchIndex
            Get #f, , Events(index).SwitchCompare
            Get #f, , Events(index).VariableIndex
            Get #f, , Events(index).VariableCompare
            Get #f, , Events(index).VariableCondition
            Get #f, , Events(index).HasItemIndex
            Get #f, , SCount
            If SCount <= 0 Then
                Events(index).HasSubEvents = False
                Erase Events(index).SubEvents
            Else
                Events(index).HasSubEvents = True
                ReDim Events(index).SubEvents(1 To SCount)
                For s = 1 To SCount
                    With Events(index).SubEvents(s)
                        Get #f, , .Type
                        Get #f, , DCount
                        If DCount <= 0 Then
                            .HasText = False
                            Erase .Text
                        Else
                            .HasText = True
                            ReDim .Text(1 To DCount)
                            For d = 1 To DCount
                                Get #f, , .Text(d)
                            Next d
                        End If
                        Get #f, , DCount
                        If DCount <= 0 Then
                            .HasData = False
                            Erase .Data
                        Else
                            .HasData = True
                            ReDim .Data(1 To DCount)
                            For d = 1 To DCount
                                Get #f, , .Data(d)
                            Next d
                        End If
                    End With
                Next s
            End If
            Get #f, , Events(index).Trigger
            Get #f, , Events(index).WalkThrought
            Get #f, , Events(index).Animated
            For s = 0 To 2
                Get #f, , Events(index).Graphic(s)
            Next
            Get #f, , Events(index).Layer
        Close #f
    Else
        Call ClearEvent(index)
        Call SaveEvent(index)
    End If
    Exit Sub
Errorhandle:
    HandleError "LoadEvent(Long)", "modEvents", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Call ClearEvent(index)
End Sub

Public Sub SaveEvents()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Call SaveEvent(i)
    Next i
End Sub
Public Sub SaveEvent(ByVal index As Long)
    Dim f As Long, SCount As Long, s As Long, DCount As Long, d As Long
    Dim FileName As String
    FileName = App.Path & "\data\events\event" & index & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Events(index).Name
        Put #f, , Events(index).chkSwitch
        Put #f, , Events(index).chkVariable
        Put #f, , Events(index).chkHasItem
        Put #f, , Events(index).SwitchIndex
        Put #f, , Events(index).SwitchCompare
        Put #f, , Events(index).VariableIndex
        Put #f, , Events(index).VariableCompare
        Put #f, , Events(index).VariableCondition
        Put #f, , Events(index).HasItemIndex
        If Not (Events(index).HasSubEvents) Then
            SCount = 0
            Put #f, , SCount
        Else
            SCount = UBound(Events(index).SubEvents)
            Put #f, , SCount
            For s = 1 To SCount
                With Events(index).SubEvents(s)
                    Put #f, , .Type
                    If Not (.HasText) Then
                        DCount = 0
                        Put #f, , DCount
                    Else
                        DCount = UBound(.Text)
                        Put #f, , DCount
                        For d = 1 To DCount
                            Put #f, , .Text(d)
                        Next d
                    End If
                    If Not (.HasData) Then
                        DCount = 0
                        Put #f, , DCount
                    Else
                        DCount = UBound(.Data)
                        Put #f, , DCount
                        For d = 1 To DCount
                            Put #f, , .Data(d)
                        Next d
                    End If
                End With
            Next s
        End If
        Put #f, , Events(index).Trigger
        Put #f, , Events(index).WalkThrought
        Put #f, , Events(index).Animated
        For s = 0 To 2
            Put #f, , Events(index).Graphic(s)
        Next
        Put #f, , Events(index).Layer
    Close #f
End Sub

Sub SaveSwitches()
Dim i As Long, FileName As String
FileName = App.Path & "\data\switches.ini"

For i = 1 To MAX_SWITCHES
    Call PutVar(FileName, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
Next

End Sub

Sub SaveVariables()
Dim i As Long, FileName As String
FileName = App.Path & "\data\variables.ini"

For i = 1 To MAX_VARIABLES
    Call PutVar(FileName, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
Next

End Sub

Sub LoadSwitches()
Dim i As Long, FileName As String
    FileName = App.Path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(FileName, "Switches", "Switch" & CStr(i) & "Name")
    Next
End Sub

Sub LoadVariables()
Dim i As Long, FileName As String
    FileName = App.Path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(FileName, "Variables", "Variable" & CStr(i) & "Name")
    Next
End Sub

Public Sub SaveTime()
    Dim FileName As String

    FileName = App.Path & "\data\time.ini"
    
    With GameTime
        PutVar FileName, "TIME", "DAY", CStr(.Day)
        PutVar FileName, "TIME", "MONTH", CStr(.Month)
        PutVar FileName, "TIME", "HOUR", CStr(.Hour)
        PutVar FileName, "TIME", "YEAR", CStr(.Year)
        PutVar FileName, "TIME", "MINUTE", CStr(.Minute)
    End With
End Sub

Public Sub LoadTime()

    Dim FileName As String

    FileName = App.Path & "\data\time.ini"
    
    With GameTime
        If FileExist(FileName, True) Then
            .Day = Val(GetVar(FileName, "TIME", "DAY"))
            .Hour = Val(GetVar(FileName, "TIME", "HOUR"))
            .Year = Val(GetVar(FileName, "TIME", "YEAR"))
            .Month = Val(GetVar(FileName, "TIME", "MONTH"))
            .Minute = Val(GetVar(FileName, "TIME", "MINUTE"))
        Else
            .Day = 1
            .Month = 1
            .Year = 1300
            SaveTime
        End If
    End With
    
End Sub

Sub SaveEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS
        Call SaveEffect(i)
    Next

End Sub

Sub SaveEffect(ByVal EffectNum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\data\effects\effect" & EffectNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Effect(EffectNum)
    Close #f
End Sub

Sub LoadEffects()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckEffects

    For i = 1 To MAX_EFFECTS
        FileName = App.Path & "\data\effects\effect" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Effect(i)
        Close #f
    Next

End Sub

Sub CheckEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS

        If Not FileExist("\Data\Effects\Effect" & i & ".dat") Then
            Call SaveEffect(i)
        End If

    Next

End Sub

Sub ClearEffect(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Effect(index)), LenB(Effect(index)))
    Effect(index).Name = vbNullString
    Effect(index).Sound = "None."
End Sub

Sub ClearEffects()
    Dim i As Long

    For i = 1 To MAX_EFFECTS
        Call ClearEffect(i)
    Next
End Sub

Public Sub LoadSwearFilter()
Dim i As Long, FileName As String, Data As String, Parse() As String
    
    On Error GoTo errorHandler
    
    FileName = App.Path & "\data\swearfilter.ini"
    ' Get the maximum amount of possible words.
    MaxSwearWords = GetVar(FileName, "SWEAR_CONFIG", "MaxWords")

    ' Check to make sure there are swear words in memory.
    If MaxSwearWords = 0 Then Exit Sub
    
    ' Redim the type to the maximum amount of words.
    ReDim SwearFilter(1 To MaxSwearWords)
    
    ' Loop through all of the words.
    For i = 1 To MaxSwearWords
        ' Get the bad word from the INI file.
        Data = GetVar(FileName, "SWEAR_FILTER", "Word_" & CStr(i))

        ' If the data isn't blank, then load it.
        If LenB(Data) <> 0 Then
            ' Split the words to be set in the database.
            Parse = Split(Data, ";")

            ' Set the values in the database.
            SwearFilter(i).BadWord = Parse(0)
            SwearFilter(i).NewWord = Parse(1)
        End If
    Next
    
    ' Error handler
   Exit Sub
errorHandler:
    HandleError "LoadSwearFilter", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

