VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer Suite"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   700
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   617
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   930
      Width           =   15360
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   5760
         ScaleHeight     =   2865
         ScaleWidth      =   4185
         TabIndex        =   7
         Top             =   3240
         Width           =   4215
         Begin VB.TextBox txtLPass 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            IMEMode         =   3  'DISABLE
            Left            =   1320
            MaxLength       =   20
            PasswordChar    =   "•"
            TabIndex        =   12
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtLUser 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   11
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtIP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1320
            TabIndex        =   10
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txtPort 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1320
            TabIndex        =   9
            Top             =   2040
            Width           =   2775
         End
         Begin VB.CheckBox chkPass 
            Appearance      =   0  'Flat
            Caption         =   "Save Password?"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblLAccept 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblBlank 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblBlank 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblBlank 
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblBlank 
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lblLogin 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LOGIN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   3975
         End
      End
   End
   Begin MSComctlLib.Toolbar tlbrSec 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   28
      Top             =   600
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imglSec"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save map"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete map"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut map"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy map"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste map"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ground"
            Object.ToolTipText     =   "Ground"
            ImageIndex      =   6
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mask"
            Object.ToolTipText     =   "Mask"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mask2"
            Object.ToolTipText     =   "Mask2"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fringe"
            Object.ToolTipText     =   "Fringe"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fringe2"
            Object.ToolTipText     =   "Fringe2"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Roof"
            Object.ToolTipText     =   "Roof"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Map"
            Object.ToolTipText     =   "Map"
            ImageIndex      =   14
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Attributes"
            Object.ToolTipText     =   "Attributes"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DirBlock"
            Object.ToolTipText     =   "Directional Block"
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Undo all changes made on this map"
            ImageIndex      =   13
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.VScrollBar scrlY 
      Height          =   9000
      Left            =   15000
      TabIndex        =   6
      Top             =   930
      Width           =   255
   End
   Begin VB.HScrollBar scrlX 
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   9930
      Width           =   11040
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Items"
            Object.ToolTipText     =   "Edit items for your game"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spells"
            Object.ToolTipText     =   "Edit spells for your game"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NPCs"
            Object.ToolTipText     =   "Edit non playable characters (NPCs)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resources"
            Object.ToolTipText     =   "Edit resources (trees, mines, fishing spots)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Events"
            Object.ToolTipText     =   "Edit events here"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Shops"
            Object.ToolTipText     =   "Edit in game shops and barters"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Animations"
            Object.ToolTipText     =   "Make animations for your game"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Effects"
            Object.ToolTipText     =   "Create amazing effects"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.StatusBar stMain 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   10155
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   609
      SimpleText      =   "Loading..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7011
            MinWidth        =   7011
            Text            =   "Loading..."
            TextSave        =   "Loading..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   20082
            MinWidth        =   20082
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglMain 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0782
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1826
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2078
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":311C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":396E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglSec 
      Left            =   0
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A12
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B24
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C36
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D48
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E5A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5610
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5962
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6006
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6358
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   3960
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   4
      Top             =   930
      Width           =   11040
   End
   Begin VB.PictureBox picAttribs 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   9255
      ScaleWidth      =   3975
      TabIndex        =   29
      Top             =   930
      Visible         =   0   'False
      Width           =   3975
      Begin VB.PictureBox picAttributes 
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   0
         ScaleHeight     =   5055
         ScaleWidth      =   3855
         TabIndex        =   46
         Top             =   4200
         Width           =   3855
         Begin VB.Frame fraSoundEffect 
            Caption         =   "Sound Effect"
            Height          =   4335
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.ComboBox cmbSoundEffect 
               Height          =   315
               ItemData        =   "frmMain.frx":6E60
               Left            =   120
               List            =   "frmMain.frx":6E70
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.Frame fraMapItem 
            Caption         =   "Map Item"
            Height          =   4335
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlMapItem 
               Height          =   255
               Left            =   120
               Max             =   10
               Min             =   1
               TabIndex        =   54
               Top             =   480
               Value           =   1
               Width           =   3495
            End
            Begin VB.HScrollBar scrlMapItemValue 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   53
               Top             =   840
               Value           =   1
               Width           =   3495
            End
            Begin VB.Label lblMapItem 
               AutoSize        =   -1  'True
               Caption         =   "Item: None x0"
               Height          =   195
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Width           =   990
            End
         End
         Begin VB.Frame fraSlide 
            Caption         =   "Slide"
            Height          =   4335
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.ComboBox cmbSlide 
               Height          =   315
               ItemData        =   "frmMain.frx":6E8B
               Left            =   120
               List            =   "frmMain.frx":6E9B
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.Frame fraHeal 
            Caption         =   "Heal"
            Height          =   4335
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlHeal 
               Height          =   255
               Left            =   120
               Max             =   10000
               TabIndex        =   62
               Top             =   840
               Width           =   3495
            End
            Begin VB.ComboBox cmbHeal 
               Height          =   315
               ItemData        =   "frmMain.frx":6EB6
               Left            =   120
               List            =   "frmMain.frx":6EC0
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lblHeal 
               AutoSize        =   -1  'True
               Caption         =   "Amount: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   600
               Width           =   720
            End
         End
         Begin VB.Frame fraTrap 
            Caption         =   "Trap"
            Height          =   4335
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlTrap 
               Height          =   255
               Left            =   120
               Max             =   10000
               TabIndex        =   58
               Top             =   600
               Width           =   3495
            End
            Begin VB.Label lblTrap 
               AutoSize        =   -1  'True
               Caption         =   "Amount: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame fraNpcSpawn 
            Caption         =   "Npc Spawn"
            Height          =   4335
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlNpcDir 
               Height          =   255
               Left            =   120
               Max             =   3
               TabIndex        =   66
               Top             =   3840
               Width           =   3495
            End
            Begin VB.ListBox lstNpc 
               Height          =   3180
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   3495
            End
            Begin VB.Label lblNpcDir 
               AutoSize        =   -1  'True
               Caption         =   "Direction: Up"
               Height          =   195
               Left            =   240
               TabIndex        =   67
               Top             =   3600
               Width           =   930
            End
         End
         Begin VB.Frame fraResource 
            Caption         =   "Resource"
            Height          =   4335
            Left            =   120
            TabIndex        =   68
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlResource 
               Height          =   255
               Left            =   120
               Max             =   100
               Min             =   1
               TabIndex        =   69
               Top             =   600
               Value           =   1
               Width           =   3495
            End
            Begin VB.Label lblResource 
               AutoSize        =   -1  'True
               Caption         =   "Resource: 1"
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   360
               Width           =   870
            End
         End
         Begin VB.Frame fraShop 
            Caption         =   "Shop"
            Height          =   4335
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlShop 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   82
               Top             =   600
               Value           =   1
               Width           =   3495
            End
            Begin VB.Label lblShop 
               Caption         =   "Shop: 1"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame fraMapWarp 
            Caption         =   "Map Warp"
            Height          =   4335
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.CheckBox chkInstanced 
               Caption         =   "Instanced"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   2040
               Width           =   1335
            End
            Begin VB.HScrollBar scrlMapWarp 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   74
               Top             =   480
               Value           =   1
               Width           =   3495
            End
            Begin VB.HScrollBar scrlMapWarpX 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   73
               Top             =   1080
               Width           =   3495
            End
            Begin VB.HScrollBar scrlMapWarpY 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   72
               Top             =   1680
               Width           =   3495
            End
            Begin VB.Label lblMapWarp 
               AutoSize        =   -1  'True
               Caption         =   "Map: 1"
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblMapWarpX 
               AutoSize        =   -1  'True
               Caption         =   "X: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   840
               Width           =   285
            End
            Begin VB.Label lblMapWarpY 
               AutoSize        =   -1  'True
               Caption         =   "Y: 0"
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   1440
               Width           =   285
            End
         End
         Begin VB.Frame fraEvent 
            Caption         =   "Event"
            Height          =   4335
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
            Begin VB.HScrollBar scrlEvent 
               Height          =   255
               Left            =   120
               Min             =   1
               TabIndex        =   50
               Top             =   600
               Value           =   1
               Width           =   3495
            End
            Begin VB.Label lblEvent 
               AutoSize        =   -1  'True
               Caption         =   "Event: 1"
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   360
               Width           =   600
            End
         End
      End
      Begin VB.Frame fraAttribs 
         Caption         =   "Attributes"
         Height          =   4215
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   3735
         Begin VB.OptionButton optNpcAvoid 
            Caption         =   "Npc Avoid"
            Height          =   270
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Caption         =   "Item"
            Height          =   270
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear2 
            Caption         =   "Clear"
            Height          =   390
            Left            =   1200
            TabIndex        =   43
            Top             =   3720
            Width           =   1455
         End
         Begin VB.OptionButton optWarp 
            Caption         =   "Warp"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            Caption         =   "Blocked"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optResource 
            Caption         =   "Resource"
            Height          =   240
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optNpcSpawn 
            Caption         =   "Npc Spawn"
            Height          =   270
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton optShop 
            Caption         =   "Shop"
            Height          =   270
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   1215
         End
         Begin VB.OptionButton optBank 
            Caption         =   "Bank"
            Height          =   270
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   1215
         End
         Begin VB.OptionButton optHeal 
            Caption         =   "Heal"
            Height          =   270
            Left            =   120
            TabIndex        =   36
            Top             =   2160
            Width           =   1215
         End
         Begin VB.OptionButton optTrap 
            Caption         =   "Trap"
            Height          =   270
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   1215
         End
         Begin VB.OptionButton optSlide 
            Caption         =   "Slide"
            Height          =   270
            Left            =   120
            TabIndex        =   34
            Top             =   2640
            Width           =   1215
         End
         Begin VB.OptionButton optEvent 
            Caption         =   "Event"
            Height          =   270
            Left            =   120
            TabIndex        =   33
            Top             =   2880
            Width           =   1215
         End
         Begin VB.OptionButton optSound 
            Caption         =   "Sound"
            Height          =   270
            Left            =   120
            TabIndex        =   32
            Top             =   3120
            Width           =   1215
         End
         Begin VB.OptionButton optThreshold 
            Caption         =   "Threshold"
            Height          =   270
            Left            =   120
            TabIndex        =   31
            Top             =   3360
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   617
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   19
      Top             =   930
      Width           =   3975
      Begin VB.CheckBox chkRoof 
         Caption         =   "Roof"
         Height          =   255
         Left            =   3120
         TabIndex        =   84
         Top             =   6000
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   270
         Left            =   2280
         TabIndex        =   81
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   2280
         TabIndex        =   80
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   255
         Left            =   3120
         TabIndex        =   79
         Top             =   6360
         Width           =   735
      End
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   1200
         Max             =   5
         TabIndex        =   26
         Top             =   6360
         Width           =   975
      End
      Begin VB.ListBox lstMaps 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         ItemData        =   "frmMain.frx":6ED2
         Left            =   0
         List            =   "frmMain.frx":6EDF
         TabIndex        =   24
         Top             =   6675
         Width           =   3945
      End
      Begin VB.VScrollBar scrlPictureY 
         Height          =   5760
         Left            =   3840
         TabIndex        =   22
         Top             =   0
         Width           =   135
      End
      Begin VB.HScrollBar scrlPictureX 
         Height          =   135
         Left            =   0
         TabIndex        =   21
         Top             =   5760
         Width           =   3840
      End
      Begin VB.HScrollBar scrlTileset 
         Height          =   255
         Left            =   1200
         Min             =   1
         TabIndex        =   20
         Top             =   6000
         Value           =   1
         Width           =   975
      End
      Begin VB.PictureBox picTileset 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   0
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   23
         Top             =   0
         Width           =   3840
         Begin MSComctlLib.ImageList imglLayers 
            Left            =   0
            Top             =   1200
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":6EF5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":7247
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":7599
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":78EB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":7C3D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMain.frx":7F8F
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label lblTileset 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tileset: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   6000
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbHeal_Click()
    MapEditorHealType = cmbHeal.ListIndex + 1
End Sub

Private Sub cmbSlide_Click()
    MapEditorSlideDir = cmbSlide.ListIndex
End Sub

Private Sub cmbSoundEffect_Click()
    MapEditorSound = cmbSoundEffect.ListIndex + 1
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    MapEditorFillLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call MapEditorClearLayer
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOptions_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroySuite
End Sub

Private Sub chkInstanced_Click()
    EditorWarpInstanced = chkInstanced.Value
End Sub

Private Sub lblLAccept_Click()
    Call SetStatus("Initializing TCP...")
    frmMain.Socket.Close
    Call TcpInit(txtIP.Text, txtPort.Text)
    Call InitMessages
    
    If MyIndex <= 0 Then
        If ConnectToServer(1) Then
            Call SetStatus("Sending login...")
            Call SendDevLogin(Trim$(txtLUser.Text), Trim$(txtLPass.Text))
        End If
    End If
    
    If Not IsConnected Then
        Call SetStatus("Server is offline")
    End If
End Sub

Private Sub lstMaps_DblClick()
    Call WarpTo(lstMaps.ListIndex + 1)
End Sub

Private Sub lstNpc_Click()
    SpawnNpcNum = lstNpc.ListIndex + 1
End Sub

Private Sub optBank_Click()
    ClearAttributeDialogue
End Sub

Private Sub optBlocked_Click()
    ClearAttributeDialogue
End Sub

Private Sub optNpcAvoid_Click()
    ClearAttributeDialogue
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorMouseDown(Button, X, Y)
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = TileView.Left + ((X) \ CELL_SIZE)
    CurY = TileView.Top + ((Y) \ CELL_SIZE)
    stMain.Panels(2).Text = Trim$(Map.name) & ": " & CurX & "-" & CurY
    Call MapEditorMouseDown(Button, X, Y)
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorChooseTile(Button, X, Y)
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorDrag(Button, X, Y)
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Fake"
        Case 3 ' animated
            lblAutotile.Caption = "Animated"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall"
    End Select
End Sub

Private Sub scrlPictureX_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlPictureY_Change()
    Call MapEditorTileScroll
End Sub

Private Sub scrlShop_Change()
    lblShop.Caption = "Shop: " & scrlShop.Value
    EditorShop = scrlShop.Value
End Sub

Private Sub scrlTileset_Change()
    lblTileset.Caption = "Tileset: " & scrlTileset.Value
    scrlPictureX.Value = 0
    scrlPictureY.Value = 0
    scrlPictureY.Max = (gTexture(Tex_Tileset(scrlTileset.Value)).RHeight - picTileset.Height) / 32
    scrlPictureX.Max = (gTexture(Tex_Tileset(scrlTileset.Value)).RWidth - picTileset.Width) / 32
    MapEditorTileScroll
End Sub

Private Sub scrlX_Change()
    UpdateCamera
End Sub

Private Sub scrlY_Change()
    UpdateCamera
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            SendRequestItems
            SendRequestEditItem
        Case 2
            SendRequestSpells
            SendRequestEditSpell
        Case 3
            SendRequestNPCS
            SendRequestEditNpc
        Case 4
            SendRequestResources
            SendRequestEditResource
        Case 5
            Call RequestSwitchesAndVariables
            Call Events_SendRequestEventsData
            Call Events_SendRequestEditEvents
        Case 6
            SendRequestShops
            SendRequestEditShop
        Case 7
            SendRequestAnimations
            SendRequestEditAnimation
        Case 8
            SendRequestEffects
            SendRequestEditEffect
    End Select
End Sub

Private Sub tlbrSec_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            SendMap
            SendMapReport
        Case 8
            CurLayer = MapLayer.Ground
            SetLayer Button.Index
        Case 9
            CurLayer = MapLayer.Mask
            SetLayer Button.Index
        Case 10
            CurLayer = MapLayer.Mask2
            SetLayer Button.Index
        Case 11
            CurLayer = MapLayer.Fringe
            SetLayer Button.Index
        Case 12
            CurLayer = MapLayer.Fringe2
            SetLayer Button.Index
        Case 13
            CurLayer = MapLayer.Roof
            SetLayer Button.Index
        Case 15
            CurEditType = EDIT_MAP
            picAttribs.Visible = False
            SetEditType Button.Index
        Case 16
            CurEditType = EDIT_ATTRIBUTES
            picAttribs.Visible = True
            SetEditType Button.Index
        Case 17
            CurEditType = EDIT_DIRBLOCK
            picAttribs.Visible = False
            SetEditType Button.Index
        Case 19
            Call WarpTo(CurrentMap)
    End Select
End Sub

Private Sub SetLayer(ByVal Layer As Long)
    tlbrSec.Buttons(8).Value = tbrUnpressed
    tlbrSec.Buttons(9).Value = tbrUnpressed
    tlbrSec.Buttons(10).Value = tbrUnpressed
    tlbrSec.Buttons(11).Value = tbrUnpressed
    tlbrSec.Buttons(12).Value = tbrUnpressed
    tlbrSec.Buttons(13).Value = tbrUnpressed
    tlbrSec.Buttons(Layer).Value = tbrPressed
End Sub

Private Sub SetEditType(ByVal EditType As Long)
    tlbrSec.Buttons(15).Value = tbrUnpressed
    tlbrSec.Buttons(16).Value = tbrUnpressed
    tlbrSec.Buttons(17).Value = tbrUnpressed
    tlbrSec.Buttons(EditType).Value = tbrPressed
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraMapWarp.Visible = True
    frmMain.scrlMapWarp.Max = MAX_MAPS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraMapItem.Visible = True
    frmMain.scrlMapItem.Max = MAX_ITEMS
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    
    
    
    fraHeal.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdHeal_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    
    fraHeal.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optHeal_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.Npc(n) > 0 Then
            lstNpc.AddItem n & ": " & Npc(Map.Npc(n)).name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeDialogue
    
    fraNpcSpawn.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optNpcSpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraResource.Visible = True
    frmMain.scrlResource.Max = MAX_RESOURCES
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraShop.Visible = True
    frmMain.scrlShop.Max = MAX_SHOPS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    
    fraSlide.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSlide_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    
    fraTrap.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optTrap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub cmdClear2_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call MapEditorClearAttribs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdClear2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub scrlHeal_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblHeal.Caption = "Amount: " & scrlHeal.Value
    MapEditorHealAmount = scrlHeal.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHeal_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTrap_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblTrap.Caption = "Amount: " & scrlTrap.Value
    MapEditorHealAmount = scrlTrap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTrap_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
        
    If Item(scrlMapItem.Value).Stackable = 1 Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If
        
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    ItemEditorNum = scrlMapItem.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlMapItem_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItem_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    ItemEditorValue = scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlMapItemValue_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapItemValue_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
    EditorWarpMap = scrlMapWarp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlMapWarp_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarp_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    EditorWarpX = scrlMapWarpX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlMapWarpX_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpX_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    EditorWarpY = scrlMapWarpY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlMapWarpY_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMapWarpY_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
    
    SpawnNpcDir = scrlNpcDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlNpcDir_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNpcDir_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblResource.Caption = "Resource: " & Resource(scrlResource.Value).name
    ResourceEditorNum = scrlResource.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    scrlResource_Change
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlResource_Scroll", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub optEvent_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    fraEvent.Visible = True
    frmMain.scrlEvent.Max = MAX_EVENTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optEvent_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlEvent_Change()
    If Trim$(Events(scrlEvent.Value).name) = vbNullString Then
        lblEvent.Caption = "Event: " & scrlEvent.Value
    Else
        lblEvent.Caption = "Event: " & scrlEvent.Value & " - " & Trim$(Events(scrlEvent.Value).name)
    End If
    MapEditorEventIndex = scrlEvent.Value
End Sub

Private Sub optSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ClearAttributeDialogue
    
    fraSoundEffect.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSound_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub optThreshold_Click()
' If debug mode, handle error then exit out
        On Error GoTo errorhandler

        ClearAttributeDialogue
        


        ' Error handler
        Exit Sub
errorhandler:
        HandleError "optThreshold_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
        Err.Clear
        Exit Sub
End Sub

