VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
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
   Icon            =   "frmEditor_Item.frx":0000
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
      TabIndex        =   88
      Top             =   720
      Width           =   3135
      Begin VB.Label lblEditor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ITEMS"
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
         TabIndex        =   89
         Top             =   160
         Width           =   2775
      End
   End
   Begin MSComctlLib.Toolbar tlbrMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   87
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
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   84
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   13
      Top             =   720
      Width           =   8535
      Begin VB.HScrollBar scrlEffect 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   85
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable?"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtPrice 
         Height          =   270
         Left            =   3840
         TabIndex        =   71
         Top             =   240
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   66
         Top             =   3000
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   64
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   3840
         Max             =   5
         TabIndex        =   20
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0782
         Left            =   3840
         List            =   "frmEditor_Item.frx":078F
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":07B8
         Left            =   120
         List            =   "frmEditor_Item.frx":07D4
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblEffect 
         AutoSize        =   -1  'True
         Caption         =   "Effect: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   86
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   67
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   65
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   63
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   26
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   25
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   2880
         TabIndex        =   24
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   23
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   2
      Top             =   4200
      Width           =   8535
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   72
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
            Picture         =   "frmEditor_Item.frx":0814
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Item.frx":0926
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Item.frx":0A38
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Item.frx":0B4A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor_Item.frx":0C5C
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraEvent 
      Caption         =   "Event"
      Height          =   3375
      Left            =   3360
      TabIndex        =   68
      Top             =   5280
      Visible         =   0   'False
      Width           =   8535
      Begin VB.HScrollBar scrlEvent 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   69
         Top             =   1440
         Value           =   1
         Width           =   8175
      End
      Begin VB.Label lblEvent 
         AutoSize        =   -1  'True
         Caption         =   "Event: 1"
         Height          =   180
         Left            =   240
         TabIndex        =   70
         Top             =   1200
         Width           =   645
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   47
      Top             =   5280
      Visible         =   0   'False
      Width           =   8535
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   48
         Top             =   1560
         Value           =   1
         Width           =   6975
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 1"
         Height          =   180
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   555
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   44
      Top             =   5280
      Visible         =   0   'False
      Width           =   8535
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   58
         Top             =   2280
         Width           =   6015
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   56
         Top             =   1560
         Width           =   6015
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   45
         Top             =   840
         Width           =   6015
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   59
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   46
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CheckBox chkTwoHanded 
         Caption         =   "Two Handed?"
         Height          =   255
         Left            =   4800
         TabIndex        =   82
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Projectile"
         Height          =   1935
         Left            =   1920
         TabIndex        =   73
         Top             =   1320
         Width           =   2655
         Begin VB.HScrollBar scrlProjectileAmmo 
            Height          =   255
            Left            =   1440
            TabIndex        =   77
            Top             =   1320
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileRotation 
            Height          =   255
            LargeChange     =   10
            Left            =   1440
            Max             =   100
            TabIndex        =   76
            Top             =   960
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileRange 
            Height          =   255
            Left            =   1440
            Max             =   255
            TabIndex        =   75
            Top             =   600
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectilePic 
            Height          =   255
            Left            =   1440
            TabIndex        =   74
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblProjectileAmmo 
            Caption         =   "Ammo: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblProjectileRotation 
            Caption         =   "Rotation: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblProjectileRange 
            Caption         =   "Range: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblProjectilePic 
            Caption         =   "Pic: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   1560
         Left            =   7200
         ScaleHeight     =   104
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   53
         Top             =   1560
         Width           =   1080
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   2880
         Width           =   2055
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4680
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   35
         Top             =   2280
         Value           =   100
         Width           =   2055
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   34
         Top             =   2640
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   32
         Top             =   3000
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   31
         Top             =   2280
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   30
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0D6E
         Left            =   1320
         List            =   "frmEditor_Item.frx":0D7E
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   3375
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   4680
         TabIndex        =   51
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   4680
         TabIndex        =   43
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   3000
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   4680
         TabIndex        =   38
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_EVENT) Then
        fraEvent.Visible = True
    Else
        fraEvent.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Call ItemEditorCancel
    frmMain.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAnim.Caption = "Anim: " & scrlAnim.Value
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim Text As String

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            Text = "+ Str: "
        Case 2
            Text = "+ End: "
        Case 3
            Text = "+ Int: "
        Case 4
            Text = "+ Agi: "
        Case 5
            Text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = Text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim Text As String
    
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            Text = "Str: "
        Case 2
            Text = "End: "
        Case 3
            Text = "Int: "
        Case 4
            Text = "Agi: "
        Case 5
            Text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = Text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpell.Caption = "Num: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub tlbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpIndex As Long
    Select Case Button.Index
        Case 1
            Call ItemEditorOk
        Case 2
            If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
            ClearItem EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ItemEditorInit
        Case 4
            If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
            TempItem = Item(EditorIndex)
            ClearItem EditorIndex
            tmpIndex = lstIndex.ListIndex
            lstIndex.RemoveItem EditorIndex - 1
            lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
            lstIndex.ListIndex = tmpIndex
            ItemEditorInit
        Case 5
            If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
            TempItem = Item(EditorIndex)
        Case 6
            If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
            If Len(Trim$(TempItem.name)) > 0 Then
                ClearItem EditorIndex
                Item(EditorIndex) = TempItem
                tmpIndex = lstIndex.ListIndex
                lstIndex.RemoveItem EditorIndex - 1
                lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
                lstIndex.ListIndex = tmpIndex
                ItemEditorInit
            End If
    End Select
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPrice_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Price = Val(txtPrice.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
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
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.Value
    Item(EditorIndex).Projectile = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub

End Sub
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRotation_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRotation.Caption = "Rotation: " & scrlProjectileRotation.Value / 2
    Item(EditorIndex).Rotation = scrlProjectileRotation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRotation_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileAmmo_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileAmmo.Caption = "Ammo: " & scrlProjectileAmmo.Value
    Item(EditorIndex).Ammo = scrlProjectileAmmo.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileAmmo_Change", "frmEditor_Item", Err.Ammober, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub chkTwoHanded_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    Item(EditorIndex).isTwoHanded = chkTwoHanded.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkTwoHanded_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub chkStackable_Click()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Stackable = chkStackable.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkStackable_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub


Private Sub scrlEffect_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblEffect.Caption = "Effect: " & scrlEffect.Value
    Item(EditorIndex).Effect = scrlEffect.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlEvent_Change()
    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    lblEvent.Caption = "Event: " & scrlEvent.Value

    Item(EditorIndex).Data1 = scrlEvent.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlEvent_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext, Erl
    Err.Clear
    Exit Sub
End Sub
