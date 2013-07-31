VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
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
   Icon            =   "frmEditor_Shop.frx":0000
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
      TabIndex        =   21
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SHOPS"
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
         TabIndex        =   22
         Top             =   160
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shop Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   5295
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2040
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   16
         Top             =   840
         Value           =   100
         Width           =   5055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   1860
         ItemData        =   "frmEditor_Shop.frx":0782
         Left            =   120
         List            =   "frmEditor_Shop.frx":079E
         TabIndex        =   8
         Top             =   2400
         Width           =   5055
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   10
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   18
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
            Picture         =   "frmEditor_Shop.frx":07C2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Shop.frx":08D4
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Shop.frx":09E6
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Shop.frx":0AF8
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Shop.frx":0C0A
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   20
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
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
 
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearShop EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ShopEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ShopEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long
Dim tmpPos As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.Text)
        .CostItem = cmbCostItem.ListIndex
        .CostValue = Val(txtCostValue.Text)
    End With
    UpdateShopTrade tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdUpdate_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDeleteTrade_Click()
Dim Index As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostValue = 0
    End With
    Call UpdateShopTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDeleteTrade_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ShopEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ShopEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBuy_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblBuy.Caption = "Buy Rate: " & scrlBuy.Value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call ShopEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_SHOPS Then Exit Sub
            ClearShop EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ShopEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_SHOPS Then Exit Sub
            TempShop = Shop(EditorIndex)
            ClearShop EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ShopEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_SHOPS Then Exit Sub
            TempShop = Shop(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_SHOPS Then Exit Sub
            If Len(Trim$(TempShop.name)) > 0 Then
                ClearShop EditorIndex
                Shop(EditorIndex) = TempShop
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                ShopEditorInit
            End If
    End Select
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
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

