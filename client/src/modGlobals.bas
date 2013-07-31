Attribute VB_Name = "modGlobals"
Option Explicit
'******************************************************
'This is base object decalaration for FMOD sound engine
Public FMOD As clsFMOD
'******************************************************

'******************************************************
'This is base object decalaration for DirectX8 graphic engine
Public Directx8 As clsDirectX8
Public D3DDevice8 As Direct3DDevice8
Public Direct3DX8 As D3DX8
'******************************************************

'elastic bars
Public BarWidth_GuiHP As Long
Public BarWidth_GuiSP As Long
Public BarWidth_GuiEXP As Long

Public BarWidth_GuiHP_Max As Long
Public BarWidth_GuiSP_Max As Long
Public BarWidth_GuiEXP_Max As Long

Public BarWidth_NpcHP(1 To MAX_NPCS) As Long
Public BarWidth_NpcHP_Max(1 To MAX_NPCS) As Long
Public BarWidth_PlayerHP(1 To MAX_PLAYERS) As Long
Public BarWidth_PlayerMP(1 To MAX_PLAYERS) As Long
Public BarWidth_PlayerHP_Max(1 To MAX_PLAYERS) As Long
Public BarWidth_PlayerMP_Max(1 To MAX_PLAYERS) As Long

' elastic camera
Public CameraLeft As Long
Public CameraTop As Long

' error handler
Public IgnoreHandler As Boolean

' fog
Public fogOffsetX As Long
Public fogOffsetY As Long

' chat bubble
Public chatBubbleIndex As Long

' map sounds
Public MapSoundCount As Long

' Map animations
Public waterfallFrame As Long
Public autoTileFrame As Long

' tutorial
Public tutorialState As Byte

' NPC Chat
Public chatNpc As Long
Public chatText As String
Public chatOptState() As Byte
Public chatContinueState As Byte
Public CurrentEventIndex As Long
Public tutOpt(1 To 4) As String
Public tutOptState(1 To 4) As Byte

' gui
Public hideGUI As Boolean
Public chatOn As Boolean
Public chatShowLine As String * 1

' fader
Public canFade As Boolean
Public faderAlpha As Long
Public faderState As Long
Public faderSpeed As Long

' menu
Public sUser As String
Public sPass As String
Public sPass2 As String
Public sChar As String
Public savePass As Boolean
Public inMenu As Boolean
Public curMenu As Long
Public curTextbox As Long

' Cursor
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long

' Paperdoll rendering order
Public PaperdollOrder() As Long

' music & sound list cache
Public musicCache() As String
Public soundCache() As String

' Amount of blood decals
Public BloodCount As Long

' main menu unloading
Public EnteringGame As Boolean

' Party GUI
Public Const Party_HPWidth As Long = 182
Public Const Party_SPRWidth As Long = 182

' targetting
Public myTarget As Long
Public myTargetType As Long

' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte

' trading
Public TradeTimer As Long
Public InTrade As Long
Public TradeX As Long
Public TradeY As Long

' Cache the Resources in an array
Public Resource_Index As Long
Public Resources_Init As Boolean

' drag + drop
Public DragInvSlotNum As Long
Public DragBankSlotNum As Long
Public DragSpell As Long

' gui
Public tmpCurrencyItem As Long
Public InShop As Long ' is the player in a shop?
Public InBank As Long
Public CurrencyMenu As Byte

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Long
Public InventoryItemSelected As Long
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long
Public TNL As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Game text buffer
Public MyText As String
Public RenderChatText As String
Public ChatScroll As Long
Public ChatButtonUp As Boolean
Public ChatButtonDown As Boolean
Public totalChatLines As Long

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Game direction vars
Public ShiftDown As Boolean
Public ControlDown As Boolean
Public tabDown As Boolean
Public wDown As Boolean
Public sDown As Boolean
Public aDown As Boolean
Public dDown As Boolean
Public upDown As Boolean
Public downDown As Boolean
Public leftDown As Boolean
Public rightDown As Boolean

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long

' Mouse cursor tile location
Public CurX As Long
Public CurY As Long

' Maximum classes
Public Max_Classes As Long
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte

' fps lock
Public FPS_Lock As Boolean

' New char
Public newCharSprite As Long
Public newCharClass As Long
Public newCharSex As Long

' looping saves
Public Player_HighIndex As Long
Public Npc_HighIndex As Long
Public Action_HighIndex As Long

' TempStrings for rendering
Public CurrencyText As String
Public CurrencyAcceptState As Byte
Public CurrencyCloseState As Byte
Public Dialogue_ButtonVisible(1 To 3) As Boolean
Public Dialogue_ButtonState(1 To 3) As Byte
Public Dialogue_TitleCaption As String
Public Dialogue_TextCaption As String
Public TradeStatus As String
Public YourWorth As String
Public TheirWorth As String

' global dialogue index
Public dialogueIndex As Long
Public dialogueData1 As Long
Public sDialogue As String

Public lastButtonSound As Long
Public lastNpcChatsound As Long

Public SStatus As String
Public Last_Dir As Long
Public LastProjectile As Integer
Public CurrentWeather As Long
Public CurrentWeatherIntensity As Byte
Public CurrentFog As Byte
Public CurrentFogSpeed As Byte
Public CurrentFogOpacity As Byte
Public CurrentTintR As Byte
Public CurrentTintG As Byte
Public CurrentTintB As Byte
Public CurrentTintA As Byte
Public DrawThunder As Byte
Public ParallaxX As Long
Public ParallaxY As Long
Public eventAnimFrame As Byte
Public eventAnimTimer As Long

Public Font_Numbers As CustomFont

Public FadeType As Long
Public FadeAmount As Long
Public FlashTimer As Long

Public IsConnecting As Boolean
Public Menu_Alert_Message As String
Public Menu_Alert_Colour As Long
Public Menu_Alert_Timer As Long

Public MaxSwearWords As Long
