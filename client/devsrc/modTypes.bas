Attribute VB_Name = "modTypes"
Public Class() As ClassRec
Public Options As OptionsRec
' autotiling
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec
Public Autotile() As AutotileRec
Public TileView As RECT

' Type recs
Private Type ClassRec
    name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type OptionsRec
    savePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
End Type

Private Type PointRec
    X As Long
    Y As Long
End Type

Private Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Private Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type
