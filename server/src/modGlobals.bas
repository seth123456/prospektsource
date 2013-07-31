Attribute VB_Name = "modGlobals"
Option Explicit

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Maximum classes
Public MAX_CLASSES As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long

' lock the CPS?
Public CPSUnlock As Boolean

Public MaxSwearWords As Long
