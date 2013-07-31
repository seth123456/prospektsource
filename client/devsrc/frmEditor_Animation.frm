VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
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
   Icon            =   "frmEditor_Animation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEditor 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   30
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ANIMATIONS"
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
         TabIndex        =   31
         Top             =   160
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animation Properties"
      Height          =   6615
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   6495
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Index           =   0
         Left            =   120
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   196
         TabIndex        =   26
         Top             =   3480
         Width           =   2940
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         Max             =   1000
         TabIndex        =   23
         Top             =   3120
         Width           =   2895
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   1000
         TabIndex        =   21
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         Max             =   100
         TabIndex        =   16
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         Max             =   100
         TabIndex        =   14
         Top             =   1920
         Width           =   2895
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2940
         Index           =   1
         Left            =   3360
         ScaleHeight     =   196
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   196
         TabIndex        =   12
         Top             =   3480
         Width           =   2940
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   100
         TabIndex        =   9
         Top             =   2520
         Width           =   3015
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         Max             =   100
         TabIndex        =   7
         Top             =   1920
         Width           =   3015
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   22
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Layer 1 (Above Player)"
         Height          =   180
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   13
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Layer 0 (Below Player)"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Animation List"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   27
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
         Height          =   6105
         ItemData        =   "frmEditor_Animation.frx":0782
         Left            =   120
         List            =   "frmEditor_Animation.frx":0784
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
            Picture         =   "frmEditor_Animation.frx":0786
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Animation.frx":0898
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Animation.frx":09AA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Animation.frx":0ABC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Animation.frx":0BCE
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   29
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
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Animation(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Animation(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    AnimationEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    AnimationEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    AnimationEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblFrameCount(Index).Caption = "Frame Count: " & scrlFrameCount(Index).Value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFrameCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlFrameCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFrameCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblLoopCount(Index).Caption = "Loop Count: " & scrlLoopCount(Index).Value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopCount_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlLoopCount_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblLoopTime(Index).Caption = "Loop Time: " & scrlLoopTime(Index).Value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLoopTime_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlLoopTime_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLoopTime_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).Value
    Animation(EditorIndex).Sprite(Index) = scrlSprite(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Scroll(Index As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlSprite_Change Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call AnimationEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
            ClearAnimation EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            AnimationEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
            TempAnimation = Animation(EditorIndex)
            ClearAnimation EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            AnimationEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
            TempAnimation = Animation(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
            If Len(Trim$(TempAnimation.name)) > 0 Then
                ClearAnimation EditorIndex
                Animation(EditorIndex) = TempAnimation
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                AnimationEditorInit
            End If
    End Select
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
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

