VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   5415
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblCPSLock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unlock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Toggle Mute"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCPSLock.Caption = "Unlock"
    Else
        CPSUnlock = True
        lblCPSLock.Caption = "Lock"
    End If
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
Dim MyText As String, Command() As String

    If KeyAscii = vbKeyReturn Then
        If Not LenB(Trim$(txtChat.Text)) > 0 Then Exit Sub
        
        ' Server message
        If Left$(txtChat.Text, 1) = "-" Then
            MyText = Mid$(txtChat.Text, 2, Len(txtChat.Text) - 1)

            If Len(MyText) > 0 Then
                Call GlobalMsg(MyText, BrightRed)
                Call TextAdd("Server: " & MyText)
            End If
            txtChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Server Commands
        If Left$(Trim$(txtChat.Text), 1) = "/" Then
            Command = Split(Trim$(txtChat.Text), Space(1))
            Select Case Command(0)
                Case "/kick"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If FindPlayer(Command(1)) > 0 Then
                        Call AlertMsg(FindPlayer(Command(1)), "You have been kicked by the server owner!")
                    Else
                        TextAdd "Wrong name."
                    End If
                Case "/disconnect"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If FindPlayer(Command(1)) > 0 Then
                        CloseSocket (FindPlayer(Command(1)))
                    Else
                        TextAdd "Wrong name."
                    End If
                Case "/motd"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    Options.MOTD = Command(1)
                    SaveOptions
                Case "/mute"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If FindPlayer(Command(1)) > 0 Then
                        Call ToggleMute(FindPlayer(Command(1)))
                    Else
                        TextAdd "Wrong name."
                    End If
                Case "/ban"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If FindPlayer(Command(1)) > 0 Then
                        Call BanIndex(FindPlayer(Command(1)))
                    Else
                        TextAdd "Wrong name."
                    End If
                Case "/setaccess"
                    If UBound(Command) < 1 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If UBound(Command) < 2 Then
                        txtChat.Text = vbNullString
                        Exit Sub
                    End If
                    If FindPlayer(Command(1)) > 0 Then
                        Call SetPlayerAccess(FindPlayer(Command(1)), CLng(Command(2)))
                        Call SendPlayerData(FindPlayer(Command(1)))
                    Else
                        TextAdd "Wrong name."
                    End If
                Case "/shutdown"
                    If isShuttingDown Then
                        isShuttingDown = False
                        GlobalMsg "Shutdown canceled.", BrightBlue
                    Else
                        isShuttingDown = True
                    End If
                Case "/cps"
                    Call SetStatus("CPS is " & GameCPS)
            End Select
        End If
        txtChat.Text = vbNullString
        KeyAscii = 0
    End If

End Sub
