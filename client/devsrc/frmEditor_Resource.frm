VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
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
   Icon            =   "frmEditor_Resource.frx":0000
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
      TabIndex        =   34
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RESOURCES"
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
         TabIndex        =   35
         Top             =   160
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   8760
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   5055
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   2880
         Max             =   5
         TabIndex        =   31
         Top             =   3120
         Width           =   2055
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2880
         Max             =   6000
         TabIndex        =   29
         Top             =   2760
         Width           =   2055
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   2880
         Max             =   6000
         TabIndex        =   26
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtHealth 
         Height          =   270
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   3975
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2040
         Width           =   3975
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   4200
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   18
         Top             =   4560
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":0782
         Left            =   960
         List            =   "frmEditor_Resource.frx":0792
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   6600
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   6
         Top             =   7200
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   5
         Top             =   4560
         Width           =   2280
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   930
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   2400
         Width           =   2100
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   20
         Top             =   3960
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   6360
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   6960
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   23
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
            Picture         =   "frmEditor_Resource.frx":07B6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Resource.frx":08C8
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Resource.frx":09DA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Resource.frx":0AEC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Resource.frx":0BFE
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   33
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
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = scrlAnimation.Value
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.Value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If scrlReward.Value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.Value).name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTool_Change()
    Dim name As String
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    Select Case scrlTool.Value
        Case 0
            name = "None"
        Case 1
            name = "Hatchet"
        Case 2
            name = "Rod"
        Case 3
            name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool Required: " & name
    
    Resource(EditorIndex).ToolRequired = scrlTool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call ResourceEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_RESOURCES Then Exit Sub
            ClearResource EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ResourceEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_RESOURCES Then Exit Sub
            TempResource = Resource(EditorIndex)
            ClearResource EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ResourceEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_RESOURCES Then Exit Sub
            TempResource = Resource(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_RESOURCES Then Exit Sub
            If Len(Trim$(TempResource.name)) > 0 Then
                ClearResource EditorIndex
                Resource(EditorIndex) = TempResource
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                ResourceEditorInit
            End If
    End Select
End Sub

Private Sub txtHealth_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Resource(EditorIndex).health = Val(txtHealth.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHealth_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
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
    Resource(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEffect_Change", "frmEditor_Effect", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
