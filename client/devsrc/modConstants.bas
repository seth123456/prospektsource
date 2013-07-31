Attribute VB_Name = "modConstants"
Public Const FPS_Lock As Boolean = False

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

Public Const GFX_EXT As String = ".png"

Public Const CELL_SIZE As Byte = 32

' General constants
Public Const MAX_PLAYERS As Byte = 50
Public Const MAX_LEVELS As Byte = 99
Public Const MAX_PARTYS As Byte = 35
Public Const MAX_MAPS As Byte = 255
Public Const MAX_ITEMS As Byte = 255
Public Const MAX_NPCS As Byte = 255
Public Const MAX_ANIMATIONS As Byte = 255
Public Const MAX_SHOPS As Byte = 255
Public Const MAX_RESOURCES As Byte = 255
Public Const MAX_EFFECTS As Byte = 255
Public Const MAX_SPELLS As Byte = 255
Public Const MAX_EVENTS As Long = 1000
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000

Public Const MAX_MAPX = 22
Public Const MAX_MAPY = 18

' Misc constants
Public Const MAX_NPC_DROPS As Byte = 30
Public Const MAX_NPC_SPELLS As Byte = 10
Public Const MAX_TRADES As Long = 30
Public Const MAX_MULTIPARTICLE As Byte = 5
Public Const MAX_MAP_NPCS As Long = 30
Public MAX_CLASSES As Long

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_EVENT As Byte = 7
Public Const EDITOR_EFFECT As Byte = 8

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Autotiles
Public Const AUTO_INNER As Byte = 1
Public Const AUTO_OUTER As Byte = 2
Public Const AUTO_HORIZONTAL As Byte = 3
Public Const AUTO_VERTICAL As Byte = 4
Public Const AUTO_FILL As Byte = 5

' Autotile types
Public Const AUTOTILE_NONE As Byte = 0
Public Const AUTOTILE_NORMAL As Byte = 1
Public Const AUTOTILE_FAKE As Byte = 2
Public Const AUTOTILE_ANIM As Byte = 3
Public Const AUTOTILE_CLIFF As Byte = 4
Public Const AUTOTILE_WATERFALL As Byte = 5

' Rendering
Public Const RENDER_STATE_NONE As Long = 0
Public Const RENDER_STATE_NORMAL As Long = 1
Public Const RENDER_STATE_AUTOTILE As Long = 2

Public Const EDIT_MAP As Byte = 0
Public Const EDIT_ATTRIBUTES As Byte = 1
Public Const EDIT_DIRBLOCK As Byte = 2
