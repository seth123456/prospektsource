VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If InGame Then Dialogue "Quit Game", "Are you sure you want to quit game?", DIALOGUE_TYPE_QUIT, True
    Cancel = YES
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_DblClick()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleDoubleClick
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleMouseUp Button
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    DestroyGame
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleMouseDown Button
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' call the procedure
    HandleMouseMove CLng(X), CLng(Y), Button

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InGame Then
        Call HandleKeyPresses(KeyAscii)
    
        ' prevents textbox on error ding sound
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    ElseIf inMenu Then
        HandleMenuKeyPresses KeyAscii
    
        ' prevents textbox on error ding sound
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(keyCode As Integer, Shift As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleKeyUp keyCode

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
