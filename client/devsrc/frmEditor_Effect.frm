VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Effect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effect Editor"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Effect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   8760
      Width           =   3135
   End
   Begin VB.PictureBox picEditor 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   49
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EFFECTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   50
         Top             =   160
         Width           =   2775
      End
   End
   Begin VB.PictureBox picEffect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   6360
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   47
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Effect Info"
      Height          =   615
      Left            =   3360
      TabIndex        =   30
      Top             =   720
      Width           =   6735
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Effect Type"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   6735
      Begin VB.OptionButton optEffectType 
         Caption         =   "Multi-Particle"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optEffectType 
         Caption         =   "Effect"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Effect List"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   52
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5910
         ItemData        =   "frmEditor_Effect.frx":0782
         Left            =   120
         List            =   "frmEditor_Effect.frx":0784
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame fraEffect 
      Caption         =   "Effect Properties"
      Height          =   4815
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   6735
      Begin VB.HScrollBar scrlModifier 
         Height          =   255
         Left            =   5280
         Max             =   255
         TabIndex        =   45
         Top             =   960
         Width           =   1335
      End
      Begin VB.HScrollBar scrlDuration 
         Height          =   255
         Left            =   1920
         Max             =   255
         TabIndex        =   36
         Top             =   960
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Effect.frx":0786
         Left            =   4080
         List            =   "frmEditor_Effect.frx":0796
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlParticles 
         Height          =   255
         LargeChange     =   10
         Left            =   1920
         Max             =   5000
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSize 
         Height          =   255
         Left            =   5280
         Max             =   255
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Graphic data"
         Height          =   3495
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
         Begin VB.HScrollBar scrlYAcc 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   43
            Top             =   3120
            Width           =   1335
         End
         Begin VB.HScrollBar scrlXAcc 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   41
            Top             =   2760
            Width           =   1335
         End
         Begin VB.HScrollBar scrlYSpeed 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   39
            Top             =   2400
            Width           =   1335
         End
         Begin VB.HScrollBar scrlXSpeed 
            Height          =   255
            Left            =   1320
            Max             =   100
            Min             =   -100
            TabIndex        =   37
            Top             =   2040
            Width           =   1335
         End
         Begin VB.HScrollBar scrlAlpha 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.HScrollBar scrlDecay 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   10
            Top             =   600
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRed 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin VB.HScrollBar scrlGreen 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   8
            Top             =   1320
            Width           =   1335
         End
         Begin VB.HScrollBar scrlBlue 
            Height          =   255
            Left            =   1320
            Max             =   100
            TabIndex        =   7
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblYAcc 
            AutoSize        =   -1  'True
            Caption         =   "YAcc: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   44
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lblXAcc 
            AutoSize        =   -1  'True
            Caption         =   "XAcc: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   42
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label lblYSpeed 
            AutoSize        =   -1  'True
            Caption         =   "YSpeed: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   2400
            Width           =   780
         End
         Begin VB.Label lblXSpeed 
            AutoSize        =   -1  'True
            Caption         =   "XSpeed: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   38
            Top             =   2040
            Width           =   780
         End
         Begin VB.Label lblAlpha 
            AutoSize        =   -1  'True
            Caption         =   "Alpha: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblDecay 
            AutoSize        =   -1  'True
            Caption         =   "Decay: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   690
         End
         Begin VB.Label lblRed 
            AutoSize        =   -1  'True
            Caption         =   "Red: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblGreen 
            AutoSize        =   -1  'True
            Caption         =   "Green: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   660
         End
         Begin VB.Label lblBlue 
            AutoSize        =   -1  'True
            Caption         =   "Blue: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   540
         End
      End
      Begin VB.Label lblModifier 
         AutoSize        =   -1  'True
         Caption         =   "Modifier: 0"
         Height          =   180
         Left            =   3480
         TabIndex        =   46
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         Caption         =   "Duration: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblParticles 
         AutoSize        =   -1  'True
         Caption         =   "Particles: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size: 0"
         Height          =   180
         Left            =   3480
         TabIndex        =   21
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.Frame fraMultiParticle 
      Caption         =   "Multi-Particle Settings"
      Height          =   975
      Left            =   3360
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   6735
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   480
         Value           =   1
         Width           =   4335
      End
      Begin VB.HScrollBar scrlMultiParticle 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   28
         Top             =   480
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label lblEffect 
         Caption         =   "Effect: XXXXXX"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblMultiParticle 
         Caption         =   "Multi-Particle: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComctlLib.ImageList imglMain 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Effect.frx":07C0
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Effect.frx":08D2
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Effect.frx":09E4
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Effect.frx":0AF6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Effect.frx":0C08
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1058
      ButtonWidth     =   1058
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cut"
            Key             =   "Cut"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Key             =   "Copy"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paste"
            Key             =   "Paste"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditor_Effect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    
    If cmbSound.ListIndex >= 0 Then
        Effect(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Effect(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub

    Effect(EditorIndex).Type = cmbType.ListIndex + 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    
    ClearEffect EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    EffectEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    
    EffectEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    EffectEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    EffectEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optEffectType_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    Select Case Index
        Case 0
            fraMultiParticle.Visible = False
            fraEffect.Visible = True
        Case 1
            fraMultiParticle.Visible = True
            fraEffect.Visible = False
    End Select
    Effect(EditorIndex).isMulti = Index
End Sub

Private Sub scrlAlpha_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblAlpha.Caption = "Alpha: " & scrlAlpha.Value / 100
    Effect(EditorIndex).Alpha = scrlAlpha.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAlpha_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBlue_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblBlue.Caption = "Blue: " & scrlBlue.Value / 100
    Effect(EditorIndex).Blue = scrlBlue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBlue_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDecay_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblDecay.Caption = "Decay: " & scrlDecay.Value / 100
    Effect(EditorIndex).Decay = scrlDecay.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDecay_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblDuration.Caption = "Duration: " & scrlDuration.Value
    Effect(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEffect_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    If scrlEffect.Value > 0 Then
        lblEffect.Caption = "Effect: " & Trim$(Effect(scrlEffect.Value).name)
    Else
        lblEffect.Caption = "Effect: None"
    End If
    
    Effect(EditorIndex).MultiParticle(scrlMultiParticle.Value) = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGreen_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblGreen.Caption = "Green: " & scrlGreen.Value / 100
    Effect(EditorIndex).Green = scrlGreen.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlGreen_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlModifier_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblModifier.Caption = "Modifier: " & scrlModifier.Value
    Effect(EditorIndex).Modifier = scrlModifier.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlModifier_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMultiParticle_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblMultiParticle.Caption = "Multi-Particle: " & scrlMultiParticle.Value
    scrlEffect.Value = Effect(EditorIndex).MultiParticle(scrlMultiParticle)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMultiParticle_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlParticles_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblParticles.Caption = "Particles: " & scrlParticles.Value
    Effect(EditorIndex).Particles = scrlParticles.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlParticles_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRed_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblRed.Caption = "Red: " & scrlRed.Value / 100
    Effect(EditorIndex).Red = scrlRed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSize_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblSize.Caption = "Size: " & scrlSize.Value
    Effect(EditorIndex).Size = scrlSize.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Effect(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    scrlSprite_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlXSpeed_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblXSpeed.Caption = "XSpeed: " & scrlXSpeed.Value
    Effect(EditorIndex).XSpeed = scrlXSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXSpeed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlYSpeed_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblYSpeed.Caption = "YSpeed: " & scrlYSpeed.Value
    Effect(EditorIndex).YSpeed = scrlYSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlYSpeed_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlXAcc_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblXAcc.Caption = "XAcc: " & scrlXAcc.Value
    Effect(EditorIndex).XAcc = scrlXAcc.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXAcc_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlYAcc_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    lblYAcc.Caption = "YAcc: " & scrlYAcc.Value
    Effect(EditorIndex).YAcc = scrlYAcc.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlYAcc_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call EffectEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
            ClearEffect EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            EffectEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
            TempEffect = Effect(EditorIndex)
            ClearEffect EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            EffectEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
            TempEffect = Effect(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
            If Len(Trim$(TempEffect.name)) > 0 Then
                ClearEffect EditorIndex
                Effect(EditorIndex) = TempEffect
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                EffectEditorInit
            End If
    End Select
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EFFECTS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Effect(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Effect(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSearch_Change()
Dim find As String, I As Long, found As Boolean
    find = txtSearch.Text

    For I = 0 To lstIndex.ListCount - 1
        If StrComp(find, Replace(lstIndex.List(I), I + 1 & ": ", ""), vbTextCompare) = 0 Then
            found = True
            lstIndex.SetFocus
            lstIndex.ListIndex = I
            Exit For
        End If
    Next
End Sub
