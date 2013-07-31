Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()
    'Loading Messages.ini Custom Messages
    Dim FileName As String
    FileName = App.path & "\data files\messages.ini"
    Dim strLoadingInterface As String
    Dim strLoadingOptions As String
    Dim strDirectX As String
    Dim strTCPIP As String
    Dim strLoadingButtons As String
    
    strLoadingInterface = GetVar(FileName, "MESSAGES", "Loading_Interfaces")
    strLoadingOptions = GetVar(FileName, "MESSAGES", "Loading_Options")
    strDirectX = GetVar(FileName, "MESSAGES", "Initializing_DirectX")
    strTCPIP = GetVar(FileName, "MESSAGES", "Init_TCPIP")
    strLoadingButtons = GetVar(FileName, "MESSAGES", "Loading_Buttons")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'Set the high-resolution timer
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    ' load gui
    Call SetStatus(strLoadingInterface)
    InitialiseGUI
    
    ' load options
    Call SetStatus(strLoadingOptions)
    LoadOptions
    
    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\data files\", "graphics"
    ChkDir App.path & "\data files\graphics\", "animations"
    ChkDir App.path & "\data files\graphics\", "characters"
    ChkDir App.path & "\data files\graphics\", "items"
    ChkDir App.path & "\data files\graphics\", "paperdolls"
    ChkDir App.path & "\data files\graphics\", "resources"
    ChkDir App.path & "\data files\graphics\", "spellicons"
    ChkDir App.path & "\data files\graphics\", "tilesets"
    ChkDir App.path & "\data files\graphics\", "faces"
    ChkDir App.path & "\data files\graphics\", "gui"
    ChkDir App.path & "\data files\graphics\gui\", "buttons"
    ChkDir App.path & "\data files\graphics\", "projectiles"
    ChkDir App.path & "\data files\graphics\", "events"
    ChkDir App.path & "\data files\graphics\", "particles"
    ChkDir App.path & "\data files\graphics\", "cursors"
    ChkDir App.path & "\data files\graphics\", "classes"
    ChkDir App.path & "\data files\graphics\", "fonts"
    ChkDir App.path & "\data files\graphics\", "panoramas"
    ChkDir App.path & "\data files\graphics\", "surfaces"
    ChkDir App.path & "\data files\", "logs"
    ChkDir App.path & "\data files\", "maps"
    ChkDir App.path & "\data files\", "music"
    ChkDir App.path & "\data files\", "sound"
    
    ' Clear game values
    Call SetStatus("Clearing game data...")
    Call ClearGameData
    
    ' load dx8
    Call SetStatus(strDirectX)
    Set Directx8 = New clsDirectX8
    Directx8.Init
    
    ' initialise sound & music engines
    Set FMOD = New clsFMOD
    FMOD.Init
    
    ' populate sound and music cache
    PopulateLists

    ' load the main game
    GettingMap = True
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus(strTCPIP)
    Call TcpInit
    Call InitMessages

    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then FMOD.Music_Play Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    ' cache the buttons then reset & render them
    Call SetStatus(strLoadingButtons)
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' set the main form size
    frmMain.Width = 12090
    frmMain.Height = 9420
    
    ' show the main menu
    frmMain.Show
    HideGame
    ShowMenu

    If ConnectToServer(1) Then
        SStatus = "Online"
    Else
        SStatus = "Offline"
    End If
    
    MenuLoop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub InitialiseGUI()

'Loading Interface.ini data
Dim FileName As String
FileName = App.path & "\data files\interface.ini"
Dim I As Long

    ' re-set chat scroll
    ChatScroll = 8

    ReDim GUIWindow(1 To GUI_Count - 1)
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .X = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .Y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .X = Val(GetVar(FileName, "GUI_HOTBAR", "X"))
        .Y = Val(GetVar(FileName, "GUI_HOTBAR", "Y"))
        .Height = Val(GetVar(FileName, "GUI_HOTBAR", "Height"))
        .Width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .X = Val(GetVar(FileName, "GUI_MENU", "X"))
        .Y = Val(GetVar(FileName, "GUI_MENU", "Y"))
        .Width = Val(GetVar(FileName, "GUI_MENU", "Width"))
        .Height = Val(GetVar(FileName, "GUI_MENU", "Height"))
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .X = Val(GetVar(FileName, "GUI_BARS", "X"))
        .Y = Val(GetVar(FileName, "GUI_BARS", "Y"))
        .Width = Val(GetVar(FileName, "GUI_BARS", "Width"))
        .Height = Val(GetVar(FileName, "GUI_BARS", "Height"))
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .X = Val(GetVar(FileName, "GUI_INVENTORY", "X"))
        .Y = Val(GetVar(FileName, "GUI_INVENTORY", "Y"))
        .Width = Val(GetVar(FileName, "GUI_INVENTORY", "Width"))
        .Height = Val(GetVar(FileName, "GUI_INVENTORY", "Height"))
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .X = Val(GetVar(FileName, "GUI_SPELLS", "X"))
        .Y = Val(GetVar(FileName, "GUI_SPELLS", "Y"))
        .Width = Val(GetVar(FileName, "GUI_SPELLS", "Width"))
        .Height = Val(GetVar(FileName, "GUI_SPELLS", "Height"))
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .X = Val(GetVar(FileName, "GUI_CHARACTER", "X"))
        .Y = Val(GetVar(FileName, "GUI_CHARACTER", "Y"))
        .Width = Val(GetVar(FileName, "GUI_CHARACTER", "Width"))
        .Height = Val(GetVar(FileName, "GUI_CHARACTER", "Height"))
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .X = Val(GetVar(FileName, "GUI_OPTIONS", "X"))
        .Y = Val(GetVar(FileName, "GUI_OPTIONS", "Y"))
        .Width = Val(GetVar(FileName, "GUI_OPTIONS", "Width"))
        .Height = Val(GetVar(FileName, "GUI_OPTIONS", "Height"))
        .visible = False
    End With
    
    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .X = Val(GetVar(FileName, "GUI_PARTY", "X"))
        .Y = Val(GetVar(FileName, "GUI_PARTY", "Y"))
        .Width = Val(GetVar(FileName, "GUI_PARTY", "Width"))
        .Height = Val(GetVar(FileName, "GUI_PARTY", "Height"))
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .X = Val(GetVar(FileName, "GUI_DESCRIPTION", "X"))
        .Y = Val(GetVar(FileName, "GUI_DESCRIPTION", "Y"))
        .Width = Val(GetVar(FileName, "GUI_DESCRIPTION", "Width"))
        .Height = Val(GetVar(FileName, "GUI_DESCRIPTION", "Height"))
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .X = Val(GetVar(FileName, "GUI_MAINMENU", "X"))
        .Y = Val(GetVar(FileName, "GUI_MAINMENU", "Y"))
        .Width = Val(GetVar(FileName, "GUI_MAINMENU", "Width"))
        .Height = Val(GetVar(FileName, "GUI_MAINMENU", "Height"))
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
         .X = Val(GetVar(FileName, "GUI_SHOP", "X"))
        .Y = Val(GetVar(FileName, "GUI_SHOP", "Y"))
        .Width = Val(GetVar(FileName, "GUI_SHOP", "Width"))
        .Height = Val(GetVar(FileName, "GUI_SHOP", "Height"))
        .visible = False
    End With
    
    ' 13 - Bank
    With GUIWindow(GUI_BANK)
        .X = Val(GetVar(FileName, "GUI_BANK", "X"))
        .Y = Val(GetVar(FileName, "GUI_BANK", "Y"))
        .Width = Val(GetVar(FileName, "GUI_BANK", "Width"))
        .Height = Val(GetVar(FileName, "GUI_BANK", "Height"))
        .visible = False
    End With
    
    ' 14 - Trade
    With GUIWindow(GUI_TRADE)
        .X = Val(GetVar(FileName, "GUI_TRADE", "X"))
        .Y = Val(GetVar(FileName, "GUI_TRADE", "Y"))
        .Width = Val(GetVar(FileName, "GUI_TRADE", "Width"))
        .Height = Val(GetVar(FileName, "GUI_TRADE", "Height"))
        .visible = False
    End With
    
    ' 15 - Currency
    With GUIWindow(GUI_CURRENCY)
        .X = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .Y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .X = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .Y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .X = Val(GetVar(FileName, "GUI_CHAT", "X"))
        .Y = Val(GetVar(FileName, "GUI_CHAT", "Y"))
        .Width = Val(GetVar(FileName, "GUI_CHAT", "Width"))
        .Height = Val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 18 - Tutorial
    With GUIWindow(GUI_TUTORIAL)
        .X = Val(GetVar(FileName, "GUI_TUTORIAL", "X"))
        .Y = Val(GetVar(FileName, "GUI_TUTORIAL", "Y"))
        .Width = Val(GetVar(FileName, "GUI_TUTORIAL", "Width"))
        .Height = Val(GetVar(FileName, "GUI_TUTORIAL", "Height"))
        .visible = False
    End With
    
    ' BUTTONS
    ' main - inv
    With Buttons(Button_Inventory)
        .state = 0 ' normal
        .X = 6
        .Y = 6
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(Button_Spells)
        .state = 0 ' normal
        .X = 81
        .Y = 6
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(Button_Character)
        .state = 0 ' normal
        .X = 156
        .Y = 6
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(Button_Options)
        .state = 0 ' normal
        .X = 6
        .Y = 41
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(Button_Trade)
        .state = 0 ' normal
        .X = 81
        .Y = 41
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(Button_Party)
        .state = 0 ' normal
        .X = 156
        .Y = 41
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 6
    End With
    
    ' menu - login
    With Buttons(Button_Login)
        .state = 0 ' normal
        .X = 54
        .Y = 277
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(Button_Register)
        .state = 0 ' normal
        .X = 154
        .Y = 277
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(Button_Credits)
        .state = 0 ' normal
        .X = 254
        .Y = 277
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(Button_Exit)
        .state = 0 ' normal
        .X = 354
        .Y = 277
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(Button_LoginAccept)
        .state = 0 ' normal
        .X = 206
        .Y = 164
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(Button_RegisterAccept)
        .state = 0 ' normal
        .X = 206
        .Y = 169
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(Button_ClassAccept)
        .state = 0 ' normal
        .X = 248
        .Y = 206
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(Button_ClassNext)
        .state = 0 ' normal
        .X = 348
        .Y = 206
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(Button_NewCharAccept)
        .state = 0 ' normal
        .X = 205
        .Y = 169
        .Width = 89
        .Height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - Select Gender Left
        With Buttons(Button_GenderLeft)
            .state = 0 'normal
            .X = 175
            .Y = 114
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 23
        End With
        
    ' main - Select Gender Right
        With Buttons(Button_GenderRight)
            .state = 0 'normal
            .X = 211
            .Y = 114
            .Width = 19
            .Height = 19
            .visible = True
            .PicNum = 24
        End With
    
    ' main - AddStats
    For I = Button_AddStats1 To Button_AddStats5
        With Buttons(I)
            .state = 0 'normal
            .Width = 12
            .Height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For I = Button_AddStats1 To Button_AddStats3 ' first 3
        With Buttons(I)
            .X = 80
            .Y = 147 + ((I - Button_AddStats1) * 15)
        End With
    Next
    For I = Button_AddStats4 To Button_AddStats5
        With Buttons(I)
            .X = 165
            .Y = 147 + ((I - Button_AddStats4) * 15)
        End With
    Next
    
    ' main - shop exit
    With Buttons(Button_ShopExit)
        .state = 0 ' normal
        .X = 90
        .Y = 276
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(Button_PartyInvite)
        .state = 0 ' normal
        .X = 14
        .Y = 209
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(Button_PartyDisband)
        .state = 0 ' normal
        .X = 101
        .Y = 209
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(Button_MusicOn)
        .state = 0 ' normal
        .X = 77
        .Y = 14
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(Button_MusicOff)
        .state = 0 ' normal
        .X = 132
        .Y = 14
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(Button_SoundOn)
        .state = 0 ' normal
        .X = 77
        .Y = 39
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(Button_SoundOff)
        .state = 0 ' normal
        .X = 132
        .Y = 39
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(Button_DebugOn)
        .state = 0 ' normal
        .X = 77
        .Y = 64
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(Button_DebugOff)
        .state = 0 ' normal
        .X = 132
        .Y = 64
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - autotile on
    With Buttons(Button_AutotileOn)
        .state = 0 ' normal
        .X = 77
        .Y = 89
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - autotile off
    With Buttons(Button_AutotileOff)
        .state = 0 ' normal
        .X = 132
        .Y = 89
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - Fullscreen on
    With Buttons(Button_FullscreenOn)
        .state = 0 ' normal
        .X = 77
        .Y = 114
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - Fullscreen off
    With Buttons(Button_FullscreenOff)
        .state = 0 ' normal
        .X = 132
        .Y = 114
        .Width = 49
        .Height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(Button_ChatUp)
        .state = 0 ' normal
        .X = 391
        .Y = 2
        .Width = 19
        .Height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(Button_ChatDown)
        .state = 0 ' normal
        .X = 391
        .Y = 104
        .Width = 19
        .Height = 19
        .visible = True
        .PicNum = 22
    End With
    
    ' main - Accept Trade
    With Buttons(Button_TradeAccept)
        .state = 0 'normal
        .X = GUIWindow(GUI_TRADE).X + 165
        .Y = GUIWindow(GUI_TRADE).Y + 335
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - Decline Trade
    With Buttons(Button_TradeDecline)
        .state = 0 'normal
        .X = GUIWindow(GUI_TRADE).X + 245
        .Y = GUIWindow(GUI_TRADE).Y + 335
        .Width = 69
        .Height = 29
        .visible = True
        .PicNum = 10
    End With

End Sub

Public Sub MenuState(ByVal state As Long)
 
    
    'Variables for loading messages.ini
    Dim FileName As String
    Dim strOfflineMessage As String
    Dim strConnectedAddChar As String
    Dim strConnectedAddAcc As String
    Dim strConnectedLogin As String
    FileName = App.path & "\data files\messages.ini"
    strOfflineMessage = GetVar(FileName, "Messages", "Server_Offline")
    strConnectedAddChar = GetVar(FileName, "Messages", "Connected_AddChar")
    strConnectedAddAcc = GetVar(FileName, "Messages", "Connected_NewAccount")
    strConnectedLogin = GetVar(FileName, "Messages", "Connected_Login")
    
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Select Case state
        Case MENU_STATE_ADDCHAR
            
            If ConnectToServer(1) Then
                Call SetStatus(strConnectedAddChar)
                Call SendAddChar(sChar, newCharSex, newCharClass, newCharSprite)
            End If
            
        Case MENU_STATE_NEWACCOUNT

            If ConnectToServer(1) Then
                Call SetStatus(strConnectedAddAcc)
                Call SendNewAccount(sUser, sPass)
            End If

        Case MENU_STATE_LOGIN
            
            If ConnectToServer(1) Then
                Call SetStatus(strConnectedLogin)
                Call SendLogin(sUser, sPass)
                Exit Sub
            End If
    End Select

    If Not IsConnected Then
        Call SetStatus(strOfflineMessage)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim I As Long

    isLogging = True
    InGame = False
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For I = 1 To MAX_BYTE
        ClearAnimInstance (I)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' destroy the chat
    For I = 1 To ChatTextBufferSize
        ChatTextBuffer(I).Text = vbNullString
    Next
    HideGame
    
    FMOD.Music_Stop
    FMOD.Music_Play Options.MenuMusic
    FMOD.StopAllSounds
    ' set the menu
    curMenu = MENU_MAIN
    
    ' show the GUI
    GUIWindow(GUI_MAINMENU).visible = True
    
    inMenu = True
    MenuLoop
End Sub

Sub GameInit()
Dim MusicFile As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' destroy the chat
    For I = 1 To ChatTextBufferSize
        ChatTextBuffer(I).Text = vbNullString
    Next
    
    ' get ping
    GetPing
    
    ' play music
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        FMOD.Music_Play MusicFile
    Else
        FMOD.Music_Stop
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' break out of GameLoop
    HideGame
    HideMenu
    Call DestroyTCP
    
    ' destroy music & sound engines
    FMOD.Destroy
    
    ' unload dx8
    Directx8.Destroy
    
    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    IsConnecting = True
    Menu_Alert_Colour = White
    Menu_Alert_Message = Caption
    Menu_Alert_Timer = timeGetTime + 3000
    DoEvents
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Prevent high ascii chars
    For I = 1 To Len(sInput)

        If Asc(Mid$(sInput, I, 1)) < vbKeySpace Or Asc(Mid$(sInput, I, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Function
End Function

Public Sub resetClickedButtons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' loop through entire array
    For I = 1 To Button_Count
        Select Case I
            ' option buttons
            Case Button_MusicOn, Button_MusicOff, Button_SoundOn, Button_SoundOff, Button_DebugOn, Button_DebugOff, Button_AutotileOn, Button_AutotileOff, Button_FullscreenOn, Button_FullscreenOff
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(I).state = 0 'normal
        End Select
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "resetClickedButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
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
ErrorHandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowMenu()
    ' set the menu
    curMenu = MENU_MAIN
    
    ' show the GUI
    GUIWindow(GUI_MAINMENU).visible = True
    
    inMenu = True
    
    ' fader
    faderAlpha = 255
    faderState = 0
    faderSpeed = 4
    canFade = True
End Sub

Public Sub HideMenu()
    GUIWindow(GUI_MAINMENU).visible = False
    inMenu = False
End Sub

Public Sub ShowGame()
Dim I As Long

    For I = GUI_CHAT To GUI_BARS
        GUIWindow(I).visible = True
    Next
    
    InGame = True
End Sub

Public Sub HideGame()
Dim I As Long
    
    For I = 1 To GUI_Count - 1
        GUIWindow(I).visible = False
    Next
    
    InGame = False
End Sub

' Converting pixels to twips and vice versa
Public Function TwipsToPixels(ByVal Twips As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        TwipsToPixels = Twips / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = Twips / Screen.TwipsPerPixelY
    End If
End Function

Public Function PixelsToTwips(ByVal Pixels As Long, ByVal XorY As Byte) As Long
    If XorY = 0 Then
        PixelsToTwips = Pixels * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = Pixels * Screen.TwipsPerPixelY
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

Public Function KeepTwoDigit(Num As Byte)
    If (Num < 10) Then
        KeepTwoDigit = "0" & Num
    Else
        KeepTwoDigit = Num
    End If
End Function

Public Sub ChangeGender()

    If newCharSex = SEX_MALE Then
        newCharSex = SEX_FEMALE
    Else
        newCharSex = SEX_MALE
    End If

End Sub

Public Sub ClearGameData()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ClearNpcs
    Call ClearResources
    Call ClearItems
    Call ClearShops
    Call ClearSpells
    Call ClearAnimations
    Call ClearEvents
    Call ClearEffects

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearGameData", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
