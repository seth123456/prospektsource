VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events Editor"
   ClientHeight    =   9300
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmEditor_Event.frx":0000
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
      TabIndex        =   45
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame fraLabeling 
      Caption         =   "Labeling Variables and Switches"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   495
      Begin VB.Frame fraRenaming 
         Caption         =   "Renaming Variable/Switch"
         Height          =   7455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   11535
         Begin VB.Frame fraRandom 
            Caption         =   "Editing Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   3360
            TabIndex        =   15
            Top             =   2640
            Width           =   5055
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   18
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3720
               TabIndex        =   17
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   720
               Width           =   4815
            End
            Begin VB.Label lblEditing 
               Caption         =   "Naming Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   8640
         TabIndex        =   42
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   10200
         TabIndex        =   41
         Top             =   7320
         Width           =   1455
      End
      Begin VB.ListBox lstVariables 
         Height          =   5520
         Left            =   1320
         TabIndex        =   23
         Top             =   840
         Width           =   4335
      End
      Begin VB.ListBox lstSwitches 
         Height          =   5520
         Left            =   5880
         TabIndex        =   22
         Top             =   840
         Width           =   4455
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Rename Variable"
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   6480
         Width           =   4335
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Rename Switch"
         Height          =   375
         Left            =   5880
         TabIndex        =   20
         Top             =   6480
         Width           =   4455
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Variables"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblRandomLabel 
         Alignment       =   2  'Center
         Caption         =   "Player Switches"
         Height          =   255
         Index           =   36
         Left            =   4560
         TabIndex        =   24
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdSwitchesVariables 
      Caption         =   "Switch/Variable"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Event List"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   43
         Text            =   "Search..."
         Top             =   240
         Width           =   2895
      End
      Begin VB.ListBox lstIndex 
         Height          =   6105
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
            Picture         =   "frmEditor_Event.frx":0782
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Event.frx":0894
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Event.frx":09A6
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Event.frx":0AB8
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Event.frx":0BCA
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   44
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
   Begin VB.Frame fraInfo 
      Caption         =   "Info"
      Height          =   7815
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   8535
      Begin VB.Frame fraEditCommand 
         Caption         =   "Edit Command"
         Height          =   5415
         Left            =   120
         TabIndex        =   78
         Top             =   2280
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CommandButton cmdEditOk 
            Caption         =   "Close"
            Height          =   375
            Left            =   6600
            TabIndex        =   79
            Top             =   4920
            Width           =   1575
         End
         Begin VB.Frame fraMapWarp 
            Caption         =   "Map Warp"
            Height          =   3735
            Left            =   120
            TabIndex        =   255
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.CheckBox chkInstanced 
               Caption         =   "Instanced"
               Height          =   255
               Left            =   120
               TabIndex        =   279
               Top             =   2040
               Width           =   2295
            End
            Begin VB.HScrollBar scrlWarpMap 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   258
               Top             =   480
               Value           =   1
               Width           =   5295
            End
            Begin VB.HScrollBar scrlWarpX 
               Height          =   255
               Left            =   120
               Max             =   250
               TabIndex        =   257
               Top             =   1080
               Value           =   1
               Width           =   5295
            End
            Begin VB.HScrollBar scrlWarpY 
               Height          =   255
               Left            =   120
               Max             =   250
               TabIndex        =   256
               Top             =   1680
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label lblWarpMap 
               Caption         =   "Map: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   261
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label lblWarpX 
               Caption         =   "X: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   260
               Top             =   840
               Width           =   5295
            End
            Begin VB.Label lblWarpY 
               Caption         =   "Y: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   259
               Top             =   1440
               Width           =   5295
            End
         End
         Begin VB.Frame fraOpenShop 
            Caption         =   "Open Shop"
            Height          =   3735
            Left            =   120
            TabIndex        =   252
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlOpenShop 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   253
               Top             =   480
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label lblOpenShop 
               Caption         =   "Open Shop: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   254
               Top             =   240
               Width           =   5295
            End
         End
         Begin VB.Frame fraGiveItem 
            Caption         =   "Change items"
            Height          =   3735
            Left            =   120
            TabIndex        =   244
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optItemOperation 
               Caption         =   "Give Item"
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   249
               Top             =   1440
               Width           =   1455
            End
            Begin VB.OptionButton optItemOperation 
               Caption         =   "Change item"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   248
               Top             =   1440
               Width           =   1455
            End
            Begin VB.OptionButton optItemOperation 
               Caption         =   "Take item"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   247
               Top             =   1440
               Width           =   1455
            End
            Begin VB.HScrollBar scrlGiveItemID 
               Height          =   255
               Left            =   120
               TabIndex        =   246
               Top             =   480
               Width           =   5295
            End
            Begin VB.HScrollBar scrlGiveItemAmount 
               Height          =   255
               Left            =   120
               Max             =   250
               Min             =   1
               TabIndex        =   245
               Top             =   1080
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label lblGiveItemID 
               Caption         =   "Item: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   251
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label lblGiveItemAmount 
               Caption         =   "Amount: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   250
               Top             =   840
               Width           =   5295
            End
         End
         Begin VB.Frame fraMenu 
            Caption         =   "Show choices"
            Height          =   3735
            Left            =   120
            TabIndex        =   233
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.TextBox txtMenuQuery 
               Height          =   645
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   240
               Top             =   480
               Width           =   5325
            End
            Begin VB.ListBox lstMenuOptions 
               Height          =   1035
               Left            =   120
               TabIndex        =   239
               Top             =   1200
               Width           =   5295
            End
            Begin VB.CommandButton cmdAddMenuOption 
               Caption         =   "Add"
               Height          =   375
               Left            =   120
               TabIndex        =   238
               Top             =   2280
               Width           =   1335
            End
            Begin VB.CommandButton cmdModifyMenuOption 
               Caption         =   "Modify"
               Height          =   375
               Left            =   2040
               TabIndex        =   237
               Top             =   2280
               Width           =   1335
            End
            Begin VB.CommandButton cmdRemoveMenuOption 
               Caption         =   "Remove"
               Height          =   375
               Left            =   3960
               TabIndex        =   236
               Top             =   2280
               Width           =   1455
            End
            Begin VB.TextBox txtMenuOptText 
               Height          =   285
               Left            =   1440
               TabIndex        =   235
               Top             =   2760
               Width           =   3855
            End
            Begin VB.HScrollBar scrlMenuOptDest 
               Height          =   255
               Left            =   240
               Max             =   10
               Min             =   1
               TabIndex        =   234
               Top             =   3360
               Value           =   1
               Width           =   5175
            End
            Begin VB.Label Label4 
               Caption         =   "Menu Query:"
               Height          =   255
               Left            =   120
               TabIndex        =   243
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "Option Text:"
               Height          =   255
               Left            =   240
               TabIndex        =   242
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label lblMenuOptDest 
               Caption         =   "Destination: 1"
               Height          =   255
               Left            =   240
               TabIndex        =   241
               Top             =   3120
               Width           =   5175
            End
         End
         Begin VB.Frame fraChangeClass 
            Caption         =   "Change Player Class"
            Height          =   3735
            Left            =   120
            TabIndex        =   230
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbChangeClass 
               Height          =   315
               Left            =   1440
               TabIndex        =   231
               Text            =   "cmbChangeClass"
               Top             =   480
               Width           =   3855
            End
            Begin VB.Label Label3 
               Caption         =   "Change class to:"
               Height          =   255
               Left            =   120
               TabIndex        =   232
               Top             =   480
               Width           =   1215
            End
         End
         Begin VB.Frame fraChangeSex 
            Caption         =   "Change Player Sex"
            Height          =   3735
            Left            =   120
            TabIndex        =   227
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optChangeSex 
               Caption         =   "Female"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   229
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton optChangeSex 
               Caption         =   "Male"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   228
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame fraChangeVariable 
            Caption         =   "Change Variable"
            Height          =   3735
            Left            =   120
            TabIndex        =   215
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbVariable 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   223
               Top             =   240
               Width           =   4215
            End
            Begin VB.OptionButton optVariableAction 
               Caption         =   "Set"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   222
               Top             =   720
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optVariableAction 
               Caption         =   "Add"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   221
               Top             =   960
               Width           =   1095
            End
            Begin VB.OptionButton optVariableAction 
               Caption         =   "Subtract"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   220
               Top             =   1200
               Width           =   1095
            End
            Begin VB.OptionButton optVariableAction 
               Caption         =   "Random"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   219
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txtVariableData 
               Height          =   285
               Index           =   0
               Left            =   1920
               TabIndex        =   218
               Text            =   "0"
               Top             =   960
               Width           =   3495
            End
            Begin VB.TextBox txtVariableData 
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2160
               TabIndex        =   217
               Text            =   "0"
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox txtVariableData 
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   3960
               TabIndex        =   216
               Text            =   "0"
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Variable:"
               Height          =   255
               Index           =   12
               Left            =   360
               TabIndex        =   226
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Low:"
               Height          =   255
               Index           =   13
               Left            =   1680
               TabIndex        =   225
               Top             =   1560
               Width           =   495
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "High:"
               Height          =   255
               Index           =   37
               Left            =   3480
               TabIndex        =   224
               Top             =   1560
               Width           =   495
            End
         End
         Begin VB.Frame fraChangeLevel 
            Caption         =   "Change Level"
            Height          =   3735
            Left            =   120
            TabIndex        =   209
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlChangeLevel 
               Height          =   255
               Left            =   120
               TabIndex        =   213
               Top             =   600
               Width           =   3855
            End
            Begin VB.OptionButton optLevelAction 
               Caption         =   "Subtract"
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   212
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton optLevelAction 
               Caption         =   "Add"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   211
               Top             =   960
               Width           =   735
            End
            Begin VB.OptionButton optLevelAction 
               Caption         =   "Set"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   210
               Top             =   960
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.Label lblChangeLevel 
               Caption         =   "Level: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   214
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame fraChangeExp 
            Caption         =   "Change Experience"
            Height          =   3735
            Left            =   120
            TabIndex        =   203
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlChangeExp 
               Height          =   255
               Left            =   120
               Max             =   32000
               TabIndex        =   207
               Top             =   600
               Width           =   3735
            End
            Begin VB.OptionButton optExpAction 
               Caption         =   "Set"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   206
               Top             =   960
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optExpAction 
               Caption         =   "Add"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   205
               Top             =   960
               Width           =   735
            End
            Begin VB.OptionButton optExpAction 
               Caption         =   "Subtract"
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   204
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lblChangeExp 
               Caption         =   "Exp: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   208
               Top             =   360
               Width           =   3735
            End
         End
         Begin VB.Frame fraChangeVitals 
            Caption         =   "Change Vitals"
            Height          =   3735
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optVitalsAction 
               Caption         =   "Subtract"
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   201
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton optVitalsAction 
               Caption         =   "Add"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   200
               Top             =   960
               Width           =   735
            End
            Begin VB.OptionButton optVitalsAction 
               Caption         =   "Set"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   199
               Top             =   960
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.ComboBox cmbChangeVitals 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0CDC
               Left            =   4080
               List            =   "frmEditor_Event.frx":0CE6
               TabIndex        =   198
               Text            =   "cmbChangeVitals"
               Top             =   600
               Width           =   1335
            End
            Begin VB.HScrollBar scrlChangeVitals 
               Height          =   255
               Left            =   120
               Max             =   32000
               TabIndex        =   197
               Top             =   600
               Width           =   3735
            End
            Begin VB.Label lblChangeVitals 
               Caption         =   "HP: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   202
               Top             =   360
               Width           =   3735
            End
         End
         Begin VB.Frame fraSetAccess 
            Caption         =   "Set Access"
            Height          =   3735
            Left            =   120
            TabIndex        =   194
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbSetAccess 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0CF8
               Left            =   240
               List            =   "frmEditor_Event.frx":0D0B
               Style           =   2  'Dropdown List
               TabIndex        =   195
               Top             =   360
               Width           =   5055
            End
         End
         Begin VB.Frame fraCustomScript 
            Caption         =   "Execute Custom Script"
            Height          =   3735
            Left            =   120
            TabIndex        =   191
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlCustomScript 
               Height          =   255
               Left            =   1560
               Max             =   255
               TabIndex        =   192
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label lblCustomScript 
               Caption         =   "Case: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   193
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame fraOpenEvent 
            Caption         =   "Open/Close event"
            Height          =   3735
            Left            =   120
            TabIndex        =   183
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlOpenEventX 
               Height          =   255
               Left            =   1560
               Max             =   255
               TabIndex        =   188
               Top             =   360
               Width           =   3855
            End
            Begin VB.HScrollBar scrlOpenEventY 
               Height          =   255
               Left            =   1560
               Max             =   255
               TabIndex        =   187
               Top             =   720
               Width           =   3855
            End
            Begin VB.OptionButton optOpenEventType 
               Caption         =   "Open"
               Height          =   255
               Index           =   0
               Left            =   1560
               TabIndex        =   186
               Top             =   1200
               Width           =   735
            End
            Begin VB.OptionButton optOpenEventType 
               Caption         =   "Close"
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   185
               Top             =   1200
               Width           =   855
            End
            Begin VB.ComboBox cmbOpenEventType 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0D4E
               Left            =   3120
               List            =   "frmEditor_Event.frx":0D5B
               TabIndex        =   184
               Text            =   "cmbOpenEventType"
               Top             =   1200
               Width           =   2295
            End
            Begin VB.Label lblOpenEventX 
               Caption         =   "X: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   190
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblOpenEventY 
               Caption         =   "Y: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   189
               Top             =   720
               Width           =   1335
            End
         End
         Begin VB.Frame fraChangeGraphic 
            Caption         =   "Change event graphic"
            Height          =   3735
            Left            =   120
            TabIndex        =   175
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlChangeGraphic 
               Height          =   255
               Left            =   1560
               Max             =   2
               TabIndex        =   179
               Top             =   1200
               Width           =   1455
            End
            Begin VB.HScrollBar scrlChangeGraphicX 
               Height          =   255
               Left            =   1560
               Max             =   255
               TabIndex        =   178
               Top             =   360
               Width           =   3855
            End
            Begin VB.HScrollBar scrlChangeGraphicY 
               Height          =   255
               Left            =   1560
               Max             =   255
               TabIndex        =   177
               Top             =   720
               Width           =   3855
            End
            Begin VB.ComboBox cmbChangeGraphicType 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0D92
               Left            =   3120
               List            =   "frmEditor_Event.frx":0D9F
               TabIndex        =   176
               Text            =   "cmbChangeGraphicType"
               Top             =   1200
               Width           =   2295
            End
            Begin VB.Label lblChangeGraphic 
               Caption         =   "Graphic#: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   182
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label lblChangeGraphicX 
               Caption         =   "X: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   181
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblChangeGraphicY 
               Caption         =   "Y: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   180
               Top             =   720
               Width           =   1335
            End
         End
         Begin VB.Frame fraChangeSprite 
            Caption         =   "Change Player Sprite"
            Height          =   3735
            Left            =   120
            TabIndex        =   172
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlChangeSprite 
               Height          =   255
               Left            =   1200
               Max             =   100
               TabIndex        =   173
               Top             =   360
               Width           =   4215
            End
            Begin VB.Label lblChangeSprite 
               Caption         =   "Sprite: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   174
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame fraChangePK 
            Caption         =   "Set Player PK"
            Height          =   3735
            Left            =   120
            TabIndex        =   169
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optChangePK 
               Caption         =   "Yes"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   171
               Top             =   840
               Width           =   1455
            End
            Begin VB.OptionButton optChangePK 
               Caption         =   "No"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   170
               Top             =   1200
               Width           =   1455
            End
         End
         Begin VB.Frame fraChangeSkill 
            Caption         =   "Change Player Skills"
            Height          =   3735
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optChangeSkills 
               Caption         =   "Remove"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   167
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton optChangeSkills 
               Caption         =   "Teach"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   166
               Top             =   720
               Width           =   855
            End
            Begin VB.ComboBox cmbChangeSkills 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0DD6
               Left            =   720
               List            =   "frmEditor_Event.frx":0DD8
               Style           =   2  'Dropdown List
               TabIndex        =   165
               Top             =   360
               Width           =   4695
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Skill:"
               Height          =   255
               Index           =   28
               Left            =   120
               TabIndex        =   168
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame fraBranch 
            Caption         =   "Conditional Branch"
            Height          =   3735
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbBranchSkill 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0DDA
               Left            =   1560
               List            =   "frmEditor_Event.frx":0DDC
               TabIndex        =   159
               Text            =   "cmbBranchSkill"
               Top             =   1680
               Width           =   1695
            End
            Begin VB.ComboBox cmbBranchClass 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0DDE
               Left            =   1560
               List            =   "frmEditor_Event.frx":0DE0
               TabIndex        =   158
               Text            =   "cmbBranchClass"
               Top             =   1320
               Width           =   1695
            End
            Begin VB.ComboBox cmbBranchItem 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0DE2
               Left            =   1560
               List            =   "frmEditor_Event.frx":0DE4
               TabIndex        =   157
               Text            =   "cmbBranchItem"
               Top             =   960
               Width           =   1695
            End
            Begin VB.ComboBox cmbBranchSwitchReq 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0DE6
               Left            =   3480
               List            =   "frmEditor_Event.frx":0DF0
               TabIndex        =   156
               Text            =   "cmbBranchSwitchReq"
               Top             =   600
               Width           =   1815
            End
            Begin VB.ComboBox cmbBranchSwitch 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0E01
               Left            =   1560
               List            =   "frmEditor_Event.frx":0E03
               TabIndex        =   155
               Text            =   "cmbBranchSwitch"
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox txtBranchLevelReq 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3360
               TabIndex        =   154
               Text            =   "0"
               Top             =   2040
               Width           =   855
            End
            Begin VB.ComboBox cmbLevelReqOperator 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0E05
               Left            =   1560
               List            =   "frmEditor_Event.frx":0E07
               TabIndex        =   153
               Text            =   "cmbLevelReqOperator"
               Top             =   2040
               Width           =   1695
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Level is"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   152
               Top             =   2040
               Width           =   975
            End
            Begin VB.ComboBox cmbVarReqOperator 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0E09
               Left            =   3480
               List            =   "frmEditor_Event.frx":0E1F
               Style           =   2  'Dropdown List
               TabIndex        =   151
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtBranchVarReq 
               Height          =   285
               Left            =   4320
               TabIndex        =   150
               Text            =   "0"
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox cmbBranchVar 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0E85
               Left            =   1560
               List            =   "frmEditor_Event.frx":0E87
               TabIndex        =   149
               Text            =   "cmbBranchVar"
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Knows Skill"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   148
               Top             =   1680
               Width           =   1215
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Class Is"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   147
               Top             =   1320
               Width           =   1215
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Has Item"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   146
               Top             =   960
               Width           =   1215
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Player Switch"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   145
               Top             =   600
               Width           =   1335
            End
            Begin VB.OptionButton optCondition_Index 
               Caption         =   "Player Variable"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   1455
            End
            Begin VB.HScrollBar scrlPositive 
               Height          =   255
               Left            =   120
               TabIndex        =   143
               Top             =   2760
               Value           =   1
               Width           =   5295
            End
            Begin VB.HScrollBar scrlNegative 
               Height          =   255
               Left            =   120
               TabIndex        =   142
               Top             =   3360
               Value           =   1
               Width           =   5295
            End
            Begin VB.TextBox txtBranchItemAmount 
               Height          =   285
               Left            =   3480
               TabIndex        =   141
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lblRandomLabel 
               Alignment       =   2  'Center
               Caption         =   "is"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   163
               Top             =   600
               Width           =   255
            End
            Begin VB.Label lblRandomLabel 
               Alignment       =   2  'Center
               Caption         =   "is"
               Height          =   255
               Index           =   0
               Left            =   3240
               TabIndex        =   162
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lblPositive 
               Caption         =   "Positive: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   161
               Top             =   2520
               Width           =   5295
            End
            Begin VB.Label lblNegative 
               Caption         =   "Negative: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   160
               Top             =   3120
               Width           =   5295
            End
         End
         Begin VB.Frame fraChangeSwitch 
            Caption         =   "Change switch"
            Height          =   3735
            Left            =   120
            TabIndex        =   135
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbSwitch 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   360
               Width           =   4215
            End
            Begin VB.ComboBox cmbPlayerSwitchSet 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0E89
               Left            =   1200
               List            =   "frmEditor_Event.frx":0E93
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   795
               Width           =   4215
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Set to:"
               Height          =   255
               Index           =   22
               Left            =   360
               TabIndex        =   139
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Switch:"
               Height          =   255
               Index           =   23
               Left            =   360
               TabIndex        =   138
               Top             =   360
               Width           =   3855
            End
         End
         Begin VB.Frame fraAnimation 
            Caption         =   "Animation"
            Height          =   3735
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlPlayAnimationAnim 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   131
               Top             =   480
               Value           =   1
               Width           =   5295
            End
            Begin VB.HScrollBar scrlPlayAnimationX 
               Height          =   255
               Left            =   120
               Max             =   250
               Min             =   -1
               TabIndex        =   130
               Top             =   1080
               Value           =   1
               Width           =   5295
            End
            Begin VB.HScrollBar scrlPlayAnimationY 
               Height          =   255
               Left            =   120
               Max             =   250
               Min             =   -1
               TabIndex        =   129
               Top             =   1680
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label lblPlayAnimationAnim 
               Caption         =   "Animation: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   134
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label lblPlayAnimationX 
               Caption         =   "X: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   133
               Top             =   840
               Width           =   5295
            End
            Begin VB.Label lblPlayAnimationY 
               Caption         =   "Y: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   132
               Top             =   1440
               Width           =   5295
            End
         End
         Begin VB.Frame fraGoTo 
            Caption         =   "GoTo"
            Height          =   3735
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.HScrollBar scrlGOTO 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   126
               Top             =   480
               Value           =   1
               Width           =   5295
            End
            Begin VB.Label lblGOTO 
               Caption         =   "Goto: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   127
               Top             =   240
               Width           =   5295
            End
         End
         Begin VB.Frame fraAddText 
            Caption         =   "Add Text"
            Height          =   3735
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optChannel 
               Caption         =   "Global"
               Height          =   255
               Index           =   2
               Left            =   2760
               TabIndex        =   122
               Top             =   2760
               Width           =   1095
            End
            Begin VB.OptionButton optChannel 
               Caption         =   "Map"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   121
               Top             =   2760
               Width           =   855
            End
            Begin VB.OptionButton optChannel 
               Caption         =   "Player"
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   120
               Top             =   2760
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.HScrollBar scrlAddText_Colour 
               Height          =   255
               Left            =   120
               Max             =   18
               TabIndex        =   119
               Top             =   2400
               Width           =   5295
            End
            Begin VB.TextBox txtAddText_Text 
               Height          =   1815
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   118
               Top             =   240
               Width           =   5295
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Channel:"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   124
               Top             =   2760
               Width           =   1575
            End
            Begin VB.Label lblAddText_Colour 
               Caption         =   "Colour: Black"
               Height          =   255
               Left            =   120
               TabIndex        =   123
               Top             =   2160
               Width           =   3255
            End
         End
         Begin VB.Frame fraPlaySound 
            Caption         =   "Play Sound"
            Height          =   3735
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbPlaySound 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0EA4
               Left            =   120
               List            =   "frmEditor_Event.frx":0EA6
               Style           =   2  'Dropdown List
               TabIndex        =   116
               Top             =   360
               Width           =   5295
            End
         End
         Begin VB.Frame fraPlayBGM 
            Caption         =   "Play Music"
            Height          =   3735
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.ComboBox cmbPlayBGM 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0EA8
               Left            =   120
               List            =   "frmEditor_Event.frx":0EAA
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   360
               Width           =   5295
            End
         End
         Begin VB.Frame fraSpecialEffect 
            Caption         =   "Special Effect"
            Height          =   3735
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.Frame fraMapOverlay 
               Caption         =   "Map Overlay"
               Height          =   2415
               Left            =   240
               TabIndex        =   103
               Top             =   840
               Visible         =   0   'False
               Width           =   5055
               Begin VB.HScrollBar scrlMapTintData 
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  Max             =   255
                  TabIndex        =   107
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.HScrollBar scrlMapTintData 
                  Height          =   255
                  Index           =   1
                  Left            =   2640
                  Max             =   255
                  TabIndex        =   106
                  Top             =   600
                  Width           =   1935
               End
               Begin VB.HScrollBar scrlMapTintData 
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  Max             =   255
                  TabIndex        =   105
                  Top             =   600
                  Width           =   1935
               End
               Begin VB.HScrollBar scrlMapTintData 
                  Height          =   255
                  Index           =   3
                  Left            =   2640
                  Max             =   255
                  TabIndex        =   104
                  Top             =   1200
                  Width           =   1935
               End
               Begin VB.Label lblMapTintData 
                  Caption         =   "Blue: 0"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   111
                  Top             =   960
                  Width           =   1815
               End
               Begin VB.Label lblMapTintData 
                  Caption         =   "Green: 0"
                  Height          =   255
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   110
                  Top             =   360
                  Width           =   1935
               End
               Begin VB.Label lblMapTintData 
                  Caption         =   "Red: 0"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   109
                  Top             =   360
                  Width           =   1935
               End
               Begin VB.Label lblMapTintData 
                  Caption         =   "Opacity: 0"
                  Height          =   255
                  Index           =   3
                  Left            =   2640
                  TabIndex        =   108
                  Top             =   960
                  Width           =   1935
               End
            End
            Begin VB.Frame fraSetFog 
               Caption         =   "Set Fog"
               Height          =   2415
               Left            =   240
               TabIndex        =   96
               Top             =   840
               Visible         =   0   'False
               Width           =   5055
               Begin VB.HScrollBar ScrlFogData 
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  Max             =   255
                  TabIndex        =   99
                  Top             =   1740
                  Width           =   4815
               End
               Begin VB.HScrollBar ScrlFogData 
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  Max             =   255
                  TabIndex        =   98
                  Top             =   600
                  Width           =   4815
               End
               Begin VB.HScrollBar ScrlFogData 
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  Max             =   255
                  TabIndex        =   97
                  Top             =   1170
                  Width           =   4815
               End
               Begin VB.Label lblFogData 
                  Caption         =   "Fog Opacity: 0"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   102
                  Top             =   1500
                  Width           =   1815
               End
               Begin VB.Label lblFogData 
                  Caption         =   "Fog: None"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   101
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.Label lblFogData 
                  Caption         =   "Fog Speed: 0"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   100
                  Top             =   930
                  Width           =   1815
               End
            End
            Begin VB.Frame fraSetWeather 
               Caption         =   "Set Weather"
               Height          =   2415
               Left            =   240
               TabIndex        =   91
               Top             =   840
               Visible         =   0   'False
               Width           =   5055
               Begin VB.ComboBox CmbWeather 
                  Height          =   315
                  ItemData        =   "frmEditor_Event.frx":0EAC
                  Left            =   120
                  List            =   "frmEditor_Event.frx":0EC2
                  Style           =   2  'Dropdown List
                  TabIndex        =   93
                  Top             =   600
                  Width           =   4695
               End
               Begin VB.HScrollBar scrlWeatherIntensity 
                  Height          =   255
                  Left            =   120
                  Max             =   100
                  TabIndex        =   92
                  Top             =   1320
                  Width           =   4695
               End
               Begin VB.Label lblRandomLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Weather Type:"
                  Height          =   195
                  Index           =   43
                  Left            =   120
                  TabIndex        =   95
                  Top             =   360
                  Width           =   1275
               End
               Begin VB.Label lblWeatherIntensity 
                  Caption         =   "Intensity: 0"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  Top             =   1080
                  Width           =   1455
               End
            End
            Begin VB.ComboBox cmbEffectType 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0EF1
               Left            =   1440
               List            =   "frmEditor_Event.frx":0F07
               TabIndex        =   90
               Text            =   "cmbEffectType"
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label6 
               Caption         =   "Effect type:"
               Height          =   255
               Left            =   240
               TabIndex        =   112
               Top             =   360
               Width           =   2175
            End
         End
         Begin VB.Frame fraPlayerText 
            Caption         =   "Show Message"
            Height          =   3735
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.TextBox txtPlayerText 
               Height          =   2175
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   86
               Top             =   240
               Width           =   5355
            End
            Begin VB.HScrollBar scrlMessageSprite 
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   3240
               Width           =   5055
            End
            Begin VB.OptionButton optMessageType 
               Caption         =   "NPC"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   84
               Top             =   2640
               Width           =   735
            End
            Begin VB.OptionButton optMessageType 
               Caption         =   "Player"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   83
               Top             =   2640
               Width           =   855
            End
            Begin VB.OptionButton optMessageType 
               Caption         =   "Static"
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   82
               Top             =   2640
               Width           =   735
            End
            Begin VB.OptionButton optMessageType 
               Caption         =   "None"
               Height          =   255
               Index           =   3
               Left            =   4320
               TabIndex        =   81
               Top             =   2640
               Width           =   735
            End
            Begin VB.Label lblMessageSprite 
               AutoSize        =   -1  'True
               Caption         =   "Graphic: 0"
               Height          =   195
               Left            =   240
               TabIndex        =   88
               Top             =   3000
               Width           =   735
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Graphic Type:"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   87
               Top             =   2640
               Width           =   1095
            End
         End
         Begin VB.Frame fraChatbubble 
            Caption         =   "Show Chatbubble"
            Height          =   3735
            Left            =   120
            TabIndex        =   262
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
            Begin VB.OptionButton optChatBubbleTarget 
               Caption         =   "Player"
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   266
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox txtChatbubbleText 
               Height          =   1005
               Left            =   1680
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   265
               Top             =   360
               Width           =   3735
            End
            Begin VB.ComboBox cmbChatBubbleTarget 
               Height          =   315
               ItemData        =   "frmEditor_Event.frx":0F48
               Left            =   120
               List            =   "frmEditor_Event.frx":0F4A
               Style           =   2  'Dropdown List
               TabIndex        =   264
               Top             =   1800
               Visible         =   0   'False
               Width           =   5295
            End
            Begin VB.OptionButton optChatBubbleTarget 
               Caption         =   "NPC"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   263
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Target Type:"
               Height          =   255
               Index           =   39
               Left            =   120
               TabIndex        =   268
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lblRandomLabel 
               Caption         =   "Chatbubble Text:"
               Height          =   255
               Index           =   38
               Left            =   120
               TabIndex        =   267
               Top             =   360
               Width           =   1575
            End
         End
      End
      Begin VB.Frame fraCommands 
         Caption         =   "Add command"
         Height          =   5415
         Left            =   120
         TabIndex        =   46
         Top             =   2280
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CommandButton cmdAddOk 
            Caption         =   "Close"
            Height          =   375
            Left            =   6720
            TabIndex        =   77
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change PK"
            Height          =   375
            Index           =   17
            Left            =   1680
            TabIndex        =   76
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change sprite"
            Height          =   375
            Index           =   16
            Left            =   1680
            TabIndex        =   75
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change skill"
            Height          =   375
            Index           =   15
            Left            =   1680
            TabIndex        =   74
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Conditional branch"
            Height          =   375
            Index           =   14
            Left            =   1680
            TabIndex        =   73
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Chatbubble"
            Height          =   375
            Index           =   13
            Left            =   1680
            TabIndex        =   72
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "GoTo"
            Height          =   375
            Index           =   9
            Left            =   1680
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Warp"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   70
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Play animation"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Level"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   68
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change items"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open Bank"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   66
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open Shop"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Exit event"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Show choices"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Show message"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Switch"
            Height          =   375
            Index           =   10
            Left            =   1680
            TabIndex        =   61
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Variable"
            Height          =   375
            Index           =   11
            Left            =   1680
            TabIndex        =   60
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Add Text"
            Height          =   375
            Index           =   12
            Left            =   1680
            TabIndex        =   59
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Special Effect"
            Height          =   375
            Index           =   29
            Left            =   4800
            TabIndex        =   58
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Stop Music"
            Height          =   375
            Index           =   28
            Left            =   4800
            TabIndex        =   57
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Play Music"
            Height          =   375
            Index           =   27
            Left            =   4800
            TabIndex        =   56
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Play Sound"
            Height          =   375
            Index           =   26
            Left            =   3240
            TabIndex        =   55
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Vitals"
            Height          =   375
            Index           =   25
            Left            =   3240
            TabIndex        =   54
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Event Graphic"
            Height          =   375
            Index           =   24
            Left            =   3240
            TabIndex        =   53
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Open/Close ev"
            Height          =   375
            Index           =   23
            Left            =   3240
            TabIndex        =   52
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Custom script"
            Height          =   375
            Index           =   22
            Left            =   3240
            TabIndex        =   51
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Set Access"
            Height          =   375
            Index           =   21
            Left            =   3240
            TabIndex        =   50
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Exp"
            Height          =   375
            Index           =   20
            Left            =   3240
            TabIndex        =   49
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Sex"
            Height          =   375
            Index           =   19
            Left            =   3240
            TabIndex        =   48
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Change Class"
            Height          =   375
            Index           =   18
            Left            =   3240
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraRandom 
         Caption         =   "Conditions"
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtPlayerVariable 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   35
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox cmbPlayerVar 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F4C
            Left            =   1200
            List            =   "frmEditor_Event.frx":0F4E
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkPlayerVar 
            Caption         =   "Variable"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F50
            Left            =   1200
            List            =   "frmEditor_Event.frx":0F52
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkPlayerSwitch 
            Caption         =   "Switch"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cmbHasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F54
            Left            =   1200
            List            =   "frmEditor_Event.frx":0F56
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkHasItem 
            Caption         =   "Has Item"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayerSwitchCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F58
            Left            =   2760
            List            =   "frmEditor_Event.frx":0F62
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbPlayerVarCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F73
            Left            =   2760
            List            =   "frmEditor_Event.frx":0F75
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Alignment       =   2  'Center
            Caption         =   "is"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   36
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Commands"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   6960
         Width           =   8295
         Begin VB.CommandButton cmdSubEventEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   3000
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdSubEventUp 
            Caption         =   "/\"
            Height          =   375
            Left            =   7680
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventDown 
            Caption         =   "\/"
            Height          =   375
            Left            =   7080
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton cmdSubEventRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdSubEventAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Tile only"
         Height          =   1935
         Left            =   4200
         TabIndex        =   38
         Top             =   240
         Width           =   4215
         Begin VB.ComboBox cmbLayer 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0F77
            Left            =   720
            List            =   "frmEditor_Event.frx":0F8D
            TabIndex        =   277
            Text            =   "Ground"
            Top             =   960
            Width           =   1695
         End
         Begin VB.HScrollBar scrlGraphic 
            Height          =   255
            Left            =   1200
            TabIndex        =   272
            Top             =   1560
            Width           =   1335
         End
         Begin VB.HScrollBar scrlCurGraphic 
            Height          =   255
            Left            =   120
            Max             =   2
            TabIndex        =   271
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox chkAnimated 
            Caption         =   "Animated"
            Height          =   255
            Left            =   1440
            TabIndex        =   270
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox picGraphic 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   2640
            ScaleHeight     =   1575
            ScaleWidth      =   1455
            TabIndex        =   269
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkWalkthrought 
            Caption         =   "Walkthrought"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbTrigger 
            Height          =   315
            ItemData        =   "frmEditor_Event.frx":0FBD
            Left            =   720
            List            =   "frmEditor_Event.frx":0FC7
            TabIndex        =   39
            Text            =   "cmbTrigger"
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Layer: "
            Height          =   255
            Left            =   120
            TabIndex        =   278
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Trigger: "
            Height          =   255
            Left            =   120
            TabIndex        =   274
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblGraphic 
            Caption         =   "Sprite(0): 0"
            Height          =   255
            Left            =   120
            TabIndex        =   273
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.ListBox lstSubEvents 
         Height          =   4545
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   8295
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.PictureBox picEditor 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   3075
      TabIndex        =   275
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EVENTS"
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
         TabIndex        =   276
         Top             =   160
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ListIndex As Long
Private RenameType As Long
Private RenameIndex As Long

Private Sub cmbBranchClass_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchClass.ListIndex + 1
End Sub

Private Sub cmbBranchItem_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchItem.ListIndex + 1
End Sub

Private Sub cmbBranchSkill_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchSkill.ListIndex + 1
End Sub

Private Sub cmbBranchSwitch_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbBranchSwitch.ListIndex
End Sub

Private Sub cmbBranchSwitchReq_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbBranchSwitchReq.ListIndex
End Sub

Private Sub cmbBranchVar_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(6) = cmbBranchVar.ListIndex
End Sub

Private Sub cmbEffectType_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    fraSetFog.Visible = False
    fraSetWeather.Visible = False
    fraMapOverlay.Visible = False
    Select Case cmbEffectType.ListIndex
        Case 3: fraSetFog.Visible = True
        Case 4: fraSetWeather.Visible = True
        Case 5: fraMapOverlay.Visible = True
    End Select
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbEffectType.ListIndex
End Sub

Private Sub cmbHasItem_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).HasItemIndex = cmbHasItem.ListIndex + 1
End Sub

Private Sub cmbChangeClass_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbChangeClass.ListIndex + 1
End Sub

Private Sub cmbChangeSkills_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbChangeSkills.ListIndex + 1
End Sub

Private Sub cmbChangeVitals_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbChangeVitals.ListIndex + 1
End Sub

Private Sub cmbChatBubbleTarget_click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbChatBubbleTarget.ListIndex + 1
End Sub

Private Sub cmbLabel_Ok_Click()
    fraLabeling.Visible = False
    SendSwitchesAndVariables
End Sub

Private Sub cmbLevelReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbLevelReqOperator.ListIndex
End Sub

Private Sub cmbPlayBGM_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(musicCache(cmbPlayBGM.ListIndex + 1))
End Sub

Private Sub cmbPlaySound_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(soundCache(cmbPlaySound.ListIndex + 1))
End Sub

Private Sub cmbPlayerSwitch_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchIndex = cmbPlayerSwitch.ListIndex
End Sub

Private Sub cmbPlayerSwitchCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).SwitchCompare = cmbPlayerSwitchCompare.ListIndex
End Sub

Private Sub cmbPlayerSwitchSet_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = cmbPlayerSwitchSet.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableIndex = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPlayerVarCompare_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).VariableCompare = cmbPlayerVarCompare.ListIndex
End Sub

Private Sub cmbSetAccess_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbSetAccess.ListIndex
End Sub

Private Sub cmbSwitch_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbSwitch.ListIndex
End Sub

Private Sub cmbTrigger_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmbVariable_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = cmbVariable.ListIndex
End Sub

Private Sub cmbVarReqOperator_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = cmbVarReqOperator.ListIndex
End Sub

Private Sub CmbWeather_click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = CmbWeather.ListIndex
End Sub

Private Sub cmdAddMenuOption_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Dim optIdx As Long
    With Events(EditorIndex).SubEvents(ListIndex)
        ReDim Preserve .Data(1 To UBound(.Data) + 1)
        ReDim Preserve .Text(1 To UBound(.Data) + 1)
        .Data(UBound(.Data)) = 1
    End With
    lstMenuOptions.AddItem ": " & 1
End Sub

Private Sub cmdAddOk_Click()
    fraCommands.Visible = False
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    Dim Count As Long
    If Not (Events(EditorIndex).HasSubEvents) Then
        ReDim Events(EditorIndex).SubEvents(1 To 1)
        Events(EditorIndex).HasSubEvents = True
    Else
        Count = UBound(Events(EditorIndex).SubEvents) + 1
        ReDim Preserve Events(EditorIndex).SubEvents(1 To Count)
    End If
    Call Events_SetSubEventType(EditorIndex, UBound(Events(EditorIndex).SubEvents), Index)
    Call PopulateSubEventList
    fraCommands.Visible = False
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex <= 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    ListIndex = 0
    ClearEvent EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Events(EditorIndex).name), EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    Event_Changed(EditorIndex) = True
    EventEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditOk_Click()
    Call PopulateSubEventList
    fraEditCommand.Visible = False
End Sub

Private Sub cmdLabel_Cancel_Click()
    fraLabeling.Visible = False
    RequestSwitchesAndVariables
End Sub

Private Sub cmdModifyMenuOption_Click()
    Dim tempIndex As Long, optIdx As Long
    tempIndex = lstSubEvents.ListIndex + 1
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).Data) Then Exit Sub
    
    Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1) = Trim$(txtMenuOptText.Text)
    Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) = scrlMenuOptDest.Value
    lstMenuOptions.List(optIdx - 1) = Trim$(txtMenuOptText.Text) & ": " & scrlMenuOptDest.Value
End Sub

Private Sub cmdRemoveMenuOption_Click()
    Dim Index As Long, I As Long
    
    Index = lstMenuOptions.ListIndex + 1
    If Index > 0 And Index < lstMenuOptions.ListCount And lstMenuOptions.ListCount > 0 Then
        For I = Index + 1 To lstMenuOptions.ListCount
            Events(EditorIndex).SubEvents(ListIndex).Data(I - 1) = Events(EditorIndex).SubEvents(ListIndex).Data(I)
            Events(EditorIndex).SubEvents(ListIndex).Text(I) = Events(EditorIndex).SubEvents(ListIndex).Text(I + 1)
        Next I
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).Data(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).Data) - 1)
        ReDim Preserve Events(EditorIndex).SubEvents(ListIndex).Text(1 To UBound(Events(EditorIndex).SubEvents(ListIndex).Text) - 1)
        Call PopulateSubEventConfig
    End If
End Sub

Private Sub cmdRename_Cancel_Click()
    Dim I As Long
    fraRenaming.Visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
    Dim I As Long
    Select Case RenameType
        Case 1
            'Variable
            If RenameIndex > 0 And RenameIndex <= MAX_VARIABLES + 1 Then
                Variables(RenameIndex) = txtRename.Text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
        Case 2
            'Switch
            If RenameIndex > 0 And RenameIndex <= MAX_SWITCHES + 1 Then
                Switches(RenameIndex) = txtRename.Text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
    End Select
    
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRenameSwitch_Click()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.Text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub cmdRenameVariable_Click()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.Text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call EventEditorOk
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call EventEditorCancel
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub cmdSubEventAdd_Click()
    fraCommands.Visible = True
End Sub

Private Sub cmdSubEventDown_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 0 And Index < lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index + 1)
        Events(EditorIndex).SubEvents(Index + 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventEdit_Click()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.Visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub cmdSubEventRemove_Click()
    Dim Index As Long, I As Long
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        For I = ListIndex + 1 To lstSubEvents.ListCount
            Events(EditorIndex).SubEvents(I - 1) = Events(EditorIndex).SubEvents(I)
        Next I
        If lstSubEvents.ListCount = 1 Then
            Events(EditorIndex).HasSubEvents = False
            Erase Events(EditorIndex).SubEvents
        Else
            ReDim Preserve Events(EditorIndex).SubEvents(1 To lstSubEvents.ListCount - 1)
        End If
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSubEventUp_Click()
    Dim Index As Long
    Index = lstSubEvents.ListIndex + 1
    If Index > 1 And Index <= lstSubEvents.ListCount Then
        Dim temp As SubEventRec
        temp = Events(EditorIndex).SubEvents(Index)
        Events(EditorIndex).SubEvents(Index) = Events(EditorIndex).SubEvents(Index - 1)
        Events(EditorIndex).SubEvents(Index - 1) = temp
        Call PopulateSubEventList
    End If
End Sub

Private Sub cmdSwitchesVariables_Click()
Dim I As Long
    fraLabeling.Visible = True
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub Form_Load()
    Dim I As Long, cap As Long
    fraLabeling.Width = 785
    fraLabeling.Height = 521
    
    ListIndex = 0
    
    cmbLevelReqOperator.Clear
    cmbPlayerVarCompare.Clear
    cmbVarReqOperator.Clear
    For I = 0 To ComparisonOperator_Count - 1
        cmbLevelReqOperator.AddItem GetComparisonOperatorName(I)
        cmbPlayerVarCompare.AddItem GetComparisonOperatorName(I)
        cmbVarReqOperator.AddItem GetComparisonOperatorName(I)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call EventEditorCancel
    frmMain.Visible = True
    ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub chkHasItem_Click()
    If chkHasItem.Value = 0 Then cmbHasItem.Enabled = False Else cmbHasItem.Enabled = True
    Events(EditorIndex).chkHasItem = chkHasItem.Value
End Sub

Private Sub chkInstanced_Click()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(4) = chkInstanced.Value
End Sub

Private Sub chkPlayerSwitch_Click()
    If chkPlayerSwitch.Value = 0 Then
        cmbPlayerSwitch.Enabled = False
        cmbPlayerSwitchCompare.Enabled = False
    Else
        cmbPlayerSwitch.Enabled = True
        cmbPlayerSwitchCompare.Enabled = True
    End If
    Events(EditorIndex).chkSwitch = chkPlayerSwitch.Value
End Sub

Private Sub chkPlayerVar_Click()
    If chkPlayerVar.Value = 0 Then
        cmbPlayerVar.Enabled = False
        txtPlayerVariable.Enabled = False
        cmbPlayerVarCompare.Enabled = False
    Else
        cmbPlayerVar.Enabled = True
        txtPlayerVariable.Enabled = True
        cmbPlayerVarCompare.Enabled = True
    End If
    Events(EditorIndex).chkVariable = chkPlayerVar.Value
End Sub

Private Sub chkWalkthrought_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).WalkThrought = chkWalkthrought.Value
End Sub
Private Sub cmbLayer_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Layer = cmbLayer.ListIndex
End Sub

Private Sub lstIndex_Click()
    EventEditorInit
End Sub

Private Sub lstMenuOptions_Click()
    Dim tempIndex As Long, optIdx As Long
    tempIndex = lstSubEvents.ListIndex + 1
    optIdx = lstMenuOptions.ListIndex + 1
    If optIdx < 1 Or optIdx > UBound(Events(EditorIndex).SubEvents(ListIndex).Data) Then Exit Sub
    
    txtMenuOptText.Text = Trim$(Events(EditorIndex).SubEvents(ListIndex).Text(optIdx + 1))
    If Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) <= 0 Then Events(EditorIndex).SubEvents(ListIndex).Data(optIdx) = 1
    scrlMenuOptDest.Value = Events(EditorIndex).SubEvents(ListIndex).Data(optIdx)
End Sub

Private Sub lstSubEvents_Click()
    ListIndex = lstSubEvents.ListIndex + 1
    If ListIndex > 0 And ListIndex < lstSubEvents.ListCount Then
        cmdSubEventDown.Enabled = True
    Else
        cmdSubEventDown.Enabled = False
    End If
    If ListIndex > 1 And ListIndex <= lstSubEvents.ListCount Then
        cmdSubEventUp.Enabled = True
    Else
        cmdSubEventUp.Enabled = False
    End If
End Sub

Private Sub lstSubEvents_DblClick()
    If ListIndex > 0 And ListIndex <= lstSubEvents.ListCount Then
        fraEditCommand.Visible = True
        PopulateSubEventConfig
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub

    cmbBranchVar.Enabled = False
    cmbVarReqOperator.Enabled = False
    txtBranchVarReq.Enabled = False
    cmbBranchSwitch.Enabled = False
    cmbBranchSwitchReq.Enabled = False
    cmbBranchItem.Enabled = False
    txtBranchItemAmount.Enabled = False
    cmbBranchClass.Enabled = False
    cmbBranchSkill.Enabled = False
    cmbLevelReqOperator.Enabled = False
    txtBranchLevelReq.Enabled = False
    
    Select Case Index
        Case 0
            cmbBranchVar.Enabled = True
            cmbVarReqOperator.Enabled = True
            txtBranchVarReq.Enabled = True
        Case 1
            cmbBranchSwitch.Enabled = True
            cmbBranchSwitchReq.Enabled = True
        Case 2
            cmbBranchItem.Enabled = True
            txtBranchItemAmount.Enabled = True
        Case 3
            cmbBranchClass.Enabled = True
        Case 4
            cmbBranchSkill.Enabled = True
        Case 5
            cmbLevelReqOperator.Enabled = True
            txtBranchLevelReq.Enabled = True
    End Select

    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optExpAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChangePK_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optChangeSex_Click(Index As Integer)
 If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optChangeSkills_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChannel_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optChatBubbleTarget_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If Index = 0 Then
        cmbChatBubbleTarget.Visible = False
    ElseIf Index = 1 Then
        cmbChatBubbleTarget.Visible = True
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = Index
End Sub

Private Sub optItemOperation_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = Index
End Sub

Private Sub optLevelAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optMessageType_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case Index
        Case 0
            scrlMessageSprite.Value = 0
            scrlMessageSprite.Enabled = False
        Case 1
            scrlMessageSprite.Enabled = True
            scrlMessageSprite.Max = MAX_NPCS
        Case 2
            scrlMessageSprite.Enabled = True
            scrlMessageSprite.Max = Count_Char
        Case 3
            scrlMessageSprite.Enabled = False
            scrlMessageSprite.Value = 0
    End Select
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
End Sub

Private Sub optOpenEventType_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = Index
End Sub

Private Sub optVariableAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Index
    Select Case Index
        Case 0, 1, 2
            txtVariableData(0).Enabled = True
            txtVariableData(1).Enabled = False
            txtVariableData(2).Enabled = False
        Case 3
            txtVariableData(0).Enabled = False
            txtVariableData(1).Enabled = True
            txtVariableData(2).Enabled = True
    End Select
End Sub

Private Sub optVitalsAction_Click(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = Index
End Sub

Private Sub scrlAddText_Colour_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlAddText_Colour.Value
End Sub

Private Sub scrlCustomScript_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblCustomScript.Caption = "Case: " & scrlCustomScript.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlCustomScript.Value
End Sub

Private Sub ScrlFogData_Change(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case Index
        Case 0
            lblFogData(Index).Caption = "Fog: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(2) = ScrlFogData(Index).Value
        Case 1
            lblFogData(Index).Caption = "Fog Speed: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(3) = ScrlFogData(Index).Value
        Case 2
            lblFogData(Index).Caption = "Fog Opacity: " & ScrlFogData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(4) = ScrlFogData(Index).Value
    End Select
End Sub

Private Sub scrlChangeExp_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeExp.Caption = "Exp: " & scrlChangeExp.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeExp.Value
End Sub

Private Sub scrlGiveItemAmount_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblGiveItemAmount.Caption = "Amount: " & scrlGiveItemAmount.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlGiveItemAmount.Value
End Sub

Private Sub scrlGiveItemID_Change()
    Dim tempIndex As Long
    tempIndex = lstSubEvents.ListIndex + 1
    lblGiveItemID.Caption = "Item: " & scrlGiveItemID.Value & "-" & Item(scrlGiveItemID.Value).name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlGiveItemID.Value
End Sub

Private Sub scrlGOTO_Change()
    lblGOTO.Caption = "Goto: " & scrlGOTO.Value
    Dim tempIndex As Long
    tempIndex = lstSubEvents.ListIndex + 1
    
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlGOTO.Value
End Sub

Private Sub scrlChangeLevel_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeLevel.Caption = "Level: " & scrlChangeLevel.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeLevel.Value
End Sub

Private Sub scrlChangeSprite_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeSprite.Caption = "Sprite: " & scrlChangeSprite.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeSprite.Value
End Sub

Private Sub scrlChangeVitals_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case cmbChangeVitals.ListIndex + 1
        Case Vitals.HP: lblChangeVitals.Caption = "Health: " & scrlChangeVitals.Value
        Case Vitals.MP: lblChangeVitals.Caption = "Mana: " & scrlChangeVitals.Value
    End Select
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeVitals.Value
End Sub

Private Sub scrlMapTintData_Change(Index As Integer)
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Select Case Index
        Case 0
            lblMapTintData(Index).Caption = "Red: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlMapTintData(Index).Value
        Case 1
            lblMapTintData(Index).Caption = "Green: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlMapTintData(Index).Value
        Case 2
            lblMapTintData(Index).Caption = "Blue: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(4) = scrlMapTintData(Index).Value
        Case 3
            lblMapTintData(Index).Caption = "Opacity: " & scrlMapTintData(Index).Value
            Events(EditorIndex).SubEvents(ListIndex).Data(5) = scrlMapTintData(Index).Value
    End Select
End Sub

Private Sub scrlMenuOptDest_Change()
    lblMenuOptDest.Caption = "Destination: " & scrlMenuOptDest.Value
End Sub

Private Sub scrlMessageSprite_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If optMessageType(1).Value Then
        If scrlMessageSprite.Value > 0 Then
            If Not Trim$(Npc(scrlMessageSprite.Value).name) = vbNullString Then
                lblMessageSprite.Caption = "NPC: " & scrlMessageSprite.Value & " - " & Trim$(Npc(scrlMessageSprite.Value).name)
            Else
                lblMessageSprite.Caption = "NPC: " & scrlMessageSprite.Value
            End If
        Else
            lblMessageSprite.Caption = "NPC: None"
        End If
    ElseIf optMessageType(2).Value Then
        lblMessageSprite.Caption = "Sprite: " & scrlMessageSprite.Value
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlMessageSprite.Value
End Sub

Private Sub scrlOpenEventX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenEventX.Caption = "X: " & scrlOpenEventX.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlOpenEventX.Value
End Sub
Private Sub scrlOpenEventY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenEventY.Caption = "Y: " & scrlOpenEventY.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlOpenEventY.Value
End Sub

Private Sub scrlOpenShop_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblOpenShop.Caption = "Open Shop: " & scrlOpenShop.Value & "-" & Shop(scrlOpenShop.Value).name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlOpenShop.Value
End Sub

Private Sub scrlPlayAnimationAnim_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPlayAnimationAnim.Caption = "Animation: " & scrlPlayAnimationAnim.Value & "-" & Animation(scrlPlayAnimationAnim.Value).name
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlPlayAnimationAnim.Value
End Sub

Private Sub scrlPlayAnimationX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationX.Value >= 0 Then
        lblPlayAnimationX.Caption = "X: " & scrlPlayAnimationX.Value
    Else
        lblPlayAnimationX.Caption = "X: Player's X Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlPlayAnimationX.Value
End Sub

Private Sub scrlPlayAnimationY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    If scrlPlayAnimationY.Value >= 0 Then
        lblPlayAnimationY.Caption = "Y: " & scrlPlayAnimationY.Value
    Else
        lblPlayAnimationY.Caption = "Y: Player's Y Position"
    End If
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlPlayAnimationY.Value
End Sub

Private Sub scrlPositive_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblPositive.Caption = "Positive: " & scrlPositive.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlPositive.Value
End Sub
Private Sub scrlNegative_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblNegative.Caption = "Negative: " & scrlNegative.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(4) = scrlNegative.Value
End Sub

Private Sub scrlWarpMap_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpMap.Caption = "Map: " & scrlWarpMap.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlWarpMap.Value
End Sub

Private Sub scrlWarpX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpX.Caption = "X: " & scrlWarpX.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlWarpX.Value
End Sub

Private Sub scrlWarpY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWarpY.Caption = "Y: " & scrlWarpY.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlWarpY.Value
End Sub

Private Sub scrlWeatherIntensity_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlWeatherIntensity.Value
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call EventEditorOk
        Case 2
            If EditorIndex = 0 Then Exit Sub
            ClearEvent EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Events(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            EventEditorInit
        Case 4
            If EditorIndex = 0 Then Exit Sub
            TempEvent = Events(EditorIndex)
            ClearEvent EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Events(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            EventEditorInit
        Case 5
            If EditorIndex = 0 Then Exit Sub
            TempEvent = Events(EditorIndex)
        Case 6
            If EditorIndex = 0 Then Exit Sub
            If Len(Trim$(TempEvent.name)) > 0 Then
                ClearEvent EditorIndex
                Events(EditorIndex) = TempEvent
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Events(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                EventEditorInit
            End If
    End Select
End Sub

Private Sub txtAddText_Text_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = Trim$(txtAddText_Text.Text)
End Sub

Private Sub txtBranchItemAmount_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(5) = Val(txtBranchItemAmount.Text)
End Sub

Private Sub txtBranchLevelReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Val(txtBranchLevelReq.Text)
End Sub

Private Sub txtBranchVarReq_Change()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = Val(txtBranchVarReq.Text)
End Sub

Private Sub txtChatbubbleText_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtChatbubbleText.Text
End Sub

Private Sub txtMenuQuery_Change()
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtMenuQuery.Text
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_EVENTS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Events(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Events(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateSubEventList()
    Dim tempIndex As Long, I As Long
    tempIndex = lstSubEvents.ListIndex
    
    lstSubEvents.Clear
    If Events(EditorIndex).HasSubEvents Then
        For I = 1 To UBound(Events(EditorIndex).SubEvents)
            lstSubEvents.AddItem I & ": " & GetEventTypeName(EditorIndex, I)
        Next I
    End If
    cmdSubEventRemove.Enabled = Events(EditorIndex).HasSubEvents
    
    If tempIndex >= 0 And tempIndex < lstSubEvents.ListCount - 1 Then lstSubEvents.ListIndex = tempIndex
    Call PopulateSubEventConfig
End Sub

Public Sub PopulateSubEventConfig()
    Dim I As Long, cap As Long
    If Not (fraEditCommand.Visible) Then Exit Sub
    If ListIndex = 0 Then Exit Sub
    HideMenus
    'Ensure Capacity
    Call Events_SetSubEventType(EditorIndex, ListIndex, Events(EditorIndex).SubEvents(ListIndex).Type)
    
    With Events(EditorIndex).SubEvents(ListIndex)
        Select Case .Type
            Case Evt_Message
                txtPlayerText.Text = Trim$(.Text(1))
                optMessageType(.Data(2)).Value = True
                scrlMessageSprite.Value = .Data(1)
                fraPlayerText.Visible = True
            Case Evt_Menu
                txtMenuQuery.Text = Trim$(.Text(1))
                lstMenuOptions.Clear
                For I = 2 To UBound(.Text)
                    lstMenuOptions.AddItem Trim$(.Text(I)) & ": " & .Data(I - 1)
                Next I
                scrlMenuOptDest.Max = UBound(Events(EditorIndex).SubEvents)
                fraMenu.Visible = True
            Case Evt_OpenShop
                If .Data(1) < 1 Or .Data(1) > MAX_SHOPS Then .Data(1) = 1
                
                scrlOpenShop.Value = .Data(1)
                Call scrlOpenShop_Change
                fraOpenShop.Visible = True
            Case Evt_GiveItem
                If .Data(1) < 1 Or .Data(1) > MAX_ITEMS Then .Data(1) = 1
                If .Data(2) < 1 Then .Data(2) = 1
                optItemOperation(.Data(3)).Value = True
                scrlGiveItemID.Value = .Data(1)
                scrlGiveItemAmount.Value = .Data(2)
                Call scrlGiveItemID_Change
                Call scrlGiveItemAmount_Change
                fraGiveItem.Visible = True
            Case Evt_PlayAnimation
                If .Data(1) < 1 Or .Data(1) > MAX_ANIMATIONS Then .Data(1) = 1
                
                scrlPlayAnimationAnim.Value = .Data(1)
                scrlPlayAnimationX.Value = .Data(2)
                scrlPlayAnimationY.Value = .Data(3)
                Call scrlPlayAnimationAnim_Change
                Call scrlPlayAnimationX_Change
                Call scrlPlayAnimationY_Change
                fraAnimation.Visible = True
            Case Evt_Warp
                If .Data(1) < 1 Or .Data(1) > MAX_MAPS Then .Data(1) = 1
                
                scrlWarpMap.Value = .Data(1)
                scrlWarpX.Value = .Data(2)
                scrlWarpY.Value = .Data(3)
                chkInstanced.Value = .Data(4)
                Call scrlWarpMap_Change
                Call scrlWarpX_Change
                Call scrlWarpY_Change
                fraMapWarp.Visible = True
            Case Evt_GOTO
                If .Data(1) < 1 Or .Data(1) > UBound(Events(EditorIndex).SubEvents) Then .Data(1) = 1
                
                scrlGOTO.Max = UBound(Events(EditorIndex).SubEvents)
                scrlGOTO.Value = .Data(1)
                Call scrlGOTO_Change
                fraGoTo.Visible = True
            Case Evt_Switch
                cmbSwitch.ListIndex = .Data(1)
                cmbPlayerSwitchSet.ListIndex = .Data(2)
                fraChangeSwitch.Visible = True
            Case Evt_Variable
                optVariableAction(.Data(1)).Value = True
                If .Data(1) = 3 Then
                    txtVariableData(1) = .Data(2)
                    txtVariableData(2) = .Data(3)
                Else
                    txtVariableData(0) = .Data(2)
                End If
                fraChangeVariable.Visible = True
            Case Evt_AddText
                txtAddText_Text.Text = Trim$(.Text(1))
                scrlAddText_Colour.Value = .Data(1)
                optChannel(.Data(2)).Value = True
                fraAddText.Visible = True
            Case Evt_Chatbubble
                txtChatbubbleText.Text = Trim$(.Text(1))
                optChatBubbleTarget(.Data(1)).Value = True
                cmbChatBubbleTarget.ListIndex = .Data(2) - 1
                fraChatbubble.Visible = True
            Case Evt_Branch
                scrlPositive.Max = UBound(Events(EditorIndex).SubEvents)
                scrlNegative.Max = UBound(Events(EditorIndex).SubEvents)
                scrlPositive.Value = .Data(3)
                scrlNegative.Value = .Data(4)
                optCondition_Index(.Data(1)) = True
                Select Case .Data(1)
                    Case 0
                        cmbBranchVar.ListIndex = .Data(6)
                        txtBranchVarReq.Text = .Data(2)
                        cmbVarReqOperator.ListIndex = .Data(5)
                    Case 1
                        cmbBranchSwitch.ListIndex = .Data(5)
                        cmbBranchSwitchReq.ListIndex = .Data(2)
                    Case 2
                        cmbBranchItem.ListIndex = .Data(2) - 1
                        txtBranchItemAmount.Text = .Data(5)
                    Case 3
                        cmbBranchClass.ListIndex = .Data(2) - 1
                    Case 4
                        cmbBranchSkill.ListIndex = .Data(2) - 1
                    Case 5
                        cmbLevelReqOperator.ListIndex = .Data(5)
                        txtBranchLevelReq.Text = .Data(2)
                End Select
                fraBranch.Visible = True
            Case Evt_ChangeSkill
                cmbChangeSkills.ListIndex = .Data(1) - 1
                optChangeSkills(.Data(2)).Value = True
                fraChangeSkill.Visible = True
            Case Evt_ChangeLevel
                scrlChangeLevel.Value = .Data(1)
                optLevelAction(.Data(2)).Value = True
                fraChangeLevel.Visible = True
            Case Evt_ChangeSprite
                scrlChangeSprite.Value = .Data(1)
                fraChangeSprite.Visible = True
            Case Evt_ChangePK
                optChangePK(.Data(1)).Value = True
                fraChangePK.Visible = True
            Case Evt_ChangeClass
                cmbChangeClass.ListIndex = .Data(1) - 1
                fraChangeClass.Visible = True
            Case Evt_ChangeSex
                optChangeSex(.Data(1)).Value = True
                fraChangeSex.Visible = True
            Case Evt_ChangeExp
                scrlChangeExp.Value = .Data(1)
                optExpAction(.Data(2)).Value = True
                fraChangeExp.Visible = True
            Case Evt_SetAccess
                cmbSetAccess.ListIndex = .Data(1)
                fraSetAccess.Visible = True
            Case Evt_CustomScript
                scrlCustomScript.Value = .Data(1)
                fraCustomScript.Visible = True
            Case Evt_OpenEvent
                scrlOpenEventX.Value = .Data(1)
                scrlOpenEventY.Value = .Data(2)
                optOpenEventType(.Data(3)).Value = True
                cmbOpenEventType.ListIndex = .Data(4)
                fraOpenEvent.Visible = True
            Case Evt_ChangeGraphic
                scrlChangeGraphicX.Value = .Data(1)
                scrlChangeGraphicY.Value = .Data(2)
                scrlChangeGraphic.Value = .Data(3)
                cmbChangeGraphicType.ListIndex = .Data(4)
                fraChangeGraphic.Visible = True
            Case Evt_ChangeVitals
                scrlChangeVitals.Value = .Data(1)
                cmbChangeVitals.ListIndex = .Data(2) - 1
                optVitalsAction(.Data(3)).Value = True
                fraChangeVitals.Visible = True
            Case Evt_PlaySound
                For I = 1 To UBound(soundCache())
                    If soundCache(I) = Trim$(.Text(1)) Then
                        cmbPlaySound.ListIndex = I - 1
                    End If
                Next
                fraPlaySound.Visible = True
            Case Evt_PlayBGM
                For I = 1 To UBound(musicCache())
                    If musicCache(I) = Trim$(.Text(1)) Then
                        cmbPlayBGM.ListIndex = I - 1
                    End If
                Next
                fraPlayBGM.Visible = True
            Case Evt_SpecialEffect
                cmbEffectType.ListIndex = .Data(1)
                Select Case .Data(1)
                    Case 3
                        ScrlFogData(0).Value = .Data(2)
                        ScrlFogData(1).Value = .Data(3)
                        ScrlFogData(2).Value = .Data(4)
                    Case 4
                        CmbWeather.ListIndex = .Data(2)
                        scrlWeatherIntensity.Value = .Data(3)
                    Case 5
                        scrlMapTintData(0).Value = .Data(2)
                        scrlMapTintData(1).Value = .Data(3)
                        scrlMapTintData(2).Value = .Data(4)
                        scrlMapTintData(3).Value = .Data(5)
                End Select
                fraSpecialEffect.Visible = True
        End Select
    End With
End Sub
Private Sub HideMenus()
    fraPlayerText.Visible = False
    fraMenu.Visible = False
    fraOpenShop.Visible = False
    fraGiveItem.Visible = False
    fraAnimation.Visible = False
    fraMapWarp.Visible = False
    fraGoTo.Visible = False
    fraChangeSwitch.Visible = False
    fraChangeVariable.Visible = False
    fraAddText.Visible = False
    fraChatbubble.Visible = False
    fraBranch.Visible = False
    fraChangeLevel.Visible = False
    fraChangeSkill.Visible = False
    fraChangeSprite.Visible = False
    fraChangePK.Visible = False
    fraChangeClass.Visible = False
    fraChangeSex.Visible = False
    fraSetAccess.Visible = False
    fraCustomScript.Visible = False
    fraOpenEvent.Visible = False
    fraChangeExp.Visible = False
    fraPlaySound.Visible = False
    fraPlayBGM.Visible = False
    fraSpecialEffect.Visible = False
End Sub

Private Sub txtPlayerText_Change()
    Dim tempIndex As Long
    tempIndex = lstSubEvents.ListIndex + 1
    Events(EditorIndex).SubEvents(ListIndex).Text(1) = txtPlayerText.Text
End Sub

Private Sub txtPlayerVariable_Change()
    Events(EditorIndex).VariableCondition = Val(txtPlayerVariable.Text)
End Sub

Private Sub txtVariableData_Change(Index As Integer)
    Select Case Index
        Case 0: Events(EditorIndex).SubEvents(ListIndex).Data(3) = Val(txtVariableData(0))
        Case 1: Events(EditorIndex).SubEvents(ListIndex).Data(3) = Val(txtVariableData(1))
        Case 2: Events(EditorIndex).SubEvents(ListIndex).Data(4) = Val(txtVariableData(2))
    End Select
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
Private Sub chkAnimated_Click()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Animated = chkAnimated.Value
End Sub
Private Sub scrlGraphic_Change()
If EditorIndex = 0 Then Exit Sub
    Events(EditorIndex).Graphic(scrlCurGraphic.Value) = scrlGraphic.Value
    lblGraphic.Caption = "Sprite(" & scrlCurGraphic.Value & "): " & scrlGraphic.Value
End Sub
Private Sub scrlCurGraphic_Change()
If EditorIndex = 0 Then Exit Sub
    scrlGraphic.Value = Events(EditorIndex).Graphic(scrlCurGraphic.Value)
    lblGraphic.Caption = "Sprite(" & scrlCurGraphic.Value & "): " & scrlGraphic.Value
End Sub

Private Sub scrlChangeGraphic_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphic.Caption = "Graphic#: " & scrlChangeGraphic.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(3) = scrlChangeGraphic.Value
End Sub

Private Sub scrlChangeGraphicX_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphicX.Caption = "X: " & scrlChangeGraphicX.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(1) = scrlChangeGraphicX.Value
End Sub

Private Sub scrlChangeGraphicY_Change()
If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    lblChangeGraphicY.Caption = "Y: " & scrlChangeGraphicY.Value
    Events(EditorIndex).SubEvents(ListIndex).Data(2) = scrlChangeGraphicY.Value
End Sub

Private Sub cmbChangeGraphicType_Click()
    If EditorIndex = 0 Or ListIndex = 0 Then Exit Sub
    Events(EditorIndex).SubEvents(ListIndex).Data(4) = cmbChangeGraphicType.ListIndex
End Sub
