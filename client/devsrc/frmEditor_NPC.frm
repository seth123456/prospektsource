VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picEditor 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   70
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NPCS"
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
         TabIndex        =   71
         Top             =   160
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   66
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Projectile"
      Height          =   1095
      Left            =   6480
      TabIndex        =   57
      Top             =   6480
      Width           =   3015
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   1440
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   1440
         Max             =   255
         TabIndex        =   59
         Top             =   480
         Width           =   975
      End
      Begin VB.HScrollBar scrlProjectileRotation 
         Height          =   255
         LargeChange     =   10
         Left            =   1440
         Max             =   100
         TabIndex        =   58
         Top             =   720
         Value           =   1
         Width           =   975
      End
      Begin VB.Label lblProjectilePic 
         Caption         =   "Pic: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblProjectileRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblProjectileRotation 
         Caption         =   "Rotation: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell"
      Height          =   1455
      Left            =   3360
      TabIndex        =   49
      Top             =   6120
      Width           =   3015
      Begin VB.HScrollBar scrlSpellNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   52
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   50
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblSpellName 
         Caption         =   "Spell: None"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info"
      Height          =   5295
      Left            =   3360
      TabIndex        =   31
      Top             =   720
      Width           =   3015
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   67
         Top             =   3960
         Width           =   2775
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":0782
         Left            =   1200
         List            =   "frmEditor_NPC.frx":078F
         TabIndex        =   64
         Text            =   "cmbMoral"
         Top             =   4800
         Width           =   1695
      End
      Begin VB.HScrollBar scrlEvent 
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Text            =   "0"
         Top             =   4440
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   2775
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   37
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   36
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   840
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":07A4
         Left            =   1200
         List            =   "frmEditor_NPC.frx":07B7
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   3600
         Width           =   930
      End
      Begin VB.Label Label6 
         Caption         =   "Moral:"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblEvent 
         Caption         =   "Event: None"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate:"
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   4440
         UseMnemonic     =   0   'False
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame Fra7 
      Caption         =   "Vitals"
      Height          =   1815
      Left            =   6480
      TabIndex        =   22
      Top             =   4560
      Width           =   3015
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   26
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   1455
      Left            =   6480
      TabIndex        =   11
      Top             =   720
      Width           =   3015
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         Left            =   1080
         Max             =   255
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         Left            =   120
         Max             =   255
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         Left            =   2040
         Max             =   255
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         Left            =   1080
         Max             =   255
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   255
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   2040
         TabIndex        =   19
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Drop"
      Height          =   2175
      Left            =   6480
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.HScrollBar scrlDrop 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   3
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chance:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   56
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   6180
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2895
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
            Picture         =   "frmEditor_NPC.frx":0809
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_NPC.frx":091B
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_NPC.frx":0A2D
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_NPC.frx":0B3F
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_NPC.frx":0C51
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   69
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
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DropIndex As Long
Private SpellIndex As Long

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call NpcEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = scrlAnimation.Value
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDrop_Change()
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.Text = Npc(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = Npc(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = Npc(EditorIndex).DropItemValue(DropIndex)
End Sub

Private Sub scrlEvent_Change()
    If scrlEvent.Value > 0 Then
        lblEvent.Caption = "Event: " & scrlEvent.Value
    Else
        lblEvent.Caption = "Event: None"
    End If
    Npc(EditorIndex).Event = scrlEvent.Value
End Sub

Private Sub scrlSpell_Change()
    SpellIndex = scrlSpell.Value
    fraSpell.Caption = "Spell - " & SpellIndex
    scrlSpellNum.Value = Npc(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlSpellNum_Change()
    lblSpellNum.Caption = "Num: " & scrlSpellNum.Value
    If scrlSpellNum.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpellNum.Value).name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    Npc(EditorIndex).Spell(SpellIndex) = scrlSpellNum.Value
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Npc(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    Npc(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).name)
    End If
    
    Npc(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    Npc(EditorIndex).Stat(Index) = scrlStat(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    Npc(EditorIndex).DropItemValue(DropIndex) = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call NpcEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
            ClearNPC EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            NpcEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
            TempNpc = Npc(EditorIndex)
            ClearNPC EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            NpcEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
            TempNpc = Npc(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
            If Len(Trim$(TempNpc.name)) > 0 Then
                ClearNPC EditorIndex
                Npc(EditorIndex) = TempNpc
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                NpcEditorInit
            End If
    End Select
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Npc(EditorIndex).AttackSay = txtAttackSay.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Validate(Cancel As Boolean)
    On Error GoTo chanceErr
    
    If DropIndex = 0 Then Exit Sub
    
    If Not IsNumeric(txtChance.Text) And Not Right$(txtChance.Text, 1) = "%" And Not InStr(1, txtChance.Text, "/") > 0 And Not InStr(1, txtChance.Text, ".") Then
        txtChance.Text = "0"
        Npc(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.Text, 1) = "%" Then
        txtChance.Text = Left(txtChance.Text, Len(txtChance.Text) - 1) / 100
    ElseIf InStr(1, txtChance.Text, "/") > 0 Then
        Dim I() As String
        I = Split(txtChance.Text, "/")
        txtChance.Text = Int(I(0) / I(1) * 1000) / 1000
    End If
    
    If txtChance.Text > 1 Or txtChance.Text < 0 Then
        Err.Description = "Value must be between 0 and 1!"
        GoTo chanceErr
    End If
    
    Npc(EditorIndex).DropChance(DropIndex) = txtChance.Text
    Exit Sub
    
chanceErr:
    txtChance.Text = "0"
    Npc(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then Npc(EditorIndex).Damage = Val(txtDamage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not Len(txtEXP.Text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.Text) Then Npc(EditorIndex).EXP = Val(txtEXP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then Npc(EditorIndex).HP = Val(txtHP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
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

Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.Value
    Npc(EditorIndex).Projectile = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub

End Sub
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Npc(EditorIndex).ProjectileRange = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRotation_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_NPCS Then Exit Sub
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.Value / 2
    Npc(EditorIndex).Rotation = scrlProjectileRotation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRotation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbMoral_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Npc(EditorIndex).Moral = cmbMoral.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbMoral_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEffect_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlEffect.Value = 0 Then
        sString = "None"
    Else
        sString = scrlEffect.Value
    End If
    lblEffect.Caption = "Effect: " & sString
    Npc(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
