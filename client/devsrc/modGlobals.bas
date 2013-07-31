Attribute VB_Name = "modGlobals"
Option Explicit

'******************************************************
'This is base object decalaration for DirectX8 graphic engine
Public Directx8 As clsDirectx8
Public D3DDevice8 As Direct3DDevice8
Public Direct3DX8 As D3DX8
'******************************************************

' Used for saving in editors
Public Effect_Changed(1 To MAX_EFFECTS) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Event_Changed(1 To MAX_EVENTS) As Boolean
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public Npc_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_RESOURCES) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean

' Main Loop
Public InSuite As Boolean
Public CurrentMap As Long

' Time-based vars
Public ElapsedTime As Long

' Global position
Public CurX As Long
Public CurY As Long

' Player vals
Public MyIndex As Long

' Game editors
Public Editor As Byte
Public EditorIndex As Long

' music & sound list cache
Public musicCache() As String
Public soundCache() As String

Public HasMap As Boolean

Public CurLayer As Long
Public CurEditType As Byte

' Used storing temporary map editor values
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long
Public MapEditorEventIndex As Long
Public MapEditorSound As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorWarpInstanced As Byte
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long
Public ItemEditorNum As Long
Public ItemEditorValue As Long
Public ResourceEditorNum As Long

' The last offset values stored, used to get the offset difference
Public LastOffsetX As Integer
Public LastOffsetY As Integer
Public ParticleOffsetX  As Long
Public ParticleOffsetY  As Long

' Animation editor
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

' Used for storing currently selected tile in map editor
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public slcTilesetTop As Long
Public slcTilesetLeft As Long
Public shpSelectedTop As Long
Public shpSelectedLeft As Long
Public shpSelectedHeight As Long
Public shpSelectedWidth As Long

' Map resource cache
Public Resource_Index As Long
Public Resources_Init As Boolean

' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte
