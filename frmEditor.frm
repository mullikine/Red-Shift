VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Red Shift - Editor"
   ClientHeight    =   6735
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   8535
   ClipControls    =   0   'False
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab2 
      Height          =   3975
      Left            =   5640
      TabIndex        =   42
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   15663103
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Preview"
      TabPicture(0)   =   "frmEditor.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgPreview"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMisc(13)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tmrRotateShip"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Timer tmrRotateShip 
         Interval        =   200
         Left            =   120
         Top             =   3360
      End
      Begin VB.Label lblMisc 
         BackStyle       =   0  'Transparent
         Caption         =   "A preview of the ship in progress."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   43
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Image imgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2415
         Left            =   120
         Picture         =   "frmEditor.frx":08E6
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   2880
      TabIndex        =   11
      Top             =   2880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   15663103
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1"
      TabPicture(0)   =   "frmEditor.frx":0BF0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblShipProperty(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblShipProperty(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblShipProperty(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblShipProperty(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblShipProperty(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblShipProperty(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtShipProperty(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtShipProperty(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtShipProperty(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtShipProperty(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtShipProperty(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtShipProperty(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "2"
      TabPicture(1)   =   "frmEditor.frx":0C0C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblShipProperty(6)"
      Tab(1).Control(1)=   "lblShipProperty(7)"
      Tab(1).Control(2)=   "lblShipProperty(8)"
      Tab(1).Control(3)=   "lblShipProperty(9)"
      Tab(1).Control(4)=   "lblShipProperty(10)"
      Tab(1).Control(5)=   "lblShipProperty(11)"
      Tab(1).Control(6)=   "txtShipProperty(6)"
      Tab(1).Control(7)=   "txtShipProperty(7)"
      Tab(1).Control(8)=   "txtShipProperty(8)"
      Tab(1).Control(9)=   "txtShipProperty(9)"
      Tab(1).Control(10)=   "txtShipProperty(10)"
      Tab(1).Control(11)=   "txtShipProperty(11)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "3"
      TabPicture(2)   =   "frmEditor.frx":0C28
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblShipProperty(12)"
      Tab(2).Control(1)=   "lblShipProperty(16)"
      Tab(2).Control(2)=   "lblShipProperty(17)"
      Tab(2).Control(3)=   "txtShipProperty(15)"
      Tab(2).Control(4)=   "txtShipProperty(16)"
      Tab(2).Control(5)=   "txtShipProperty(17)"
      Tab(2).ControlCount=   6
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "123"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "123"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "123"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "123"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "123"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "123"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "123"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   24
         Text            =   "123"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   120
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "123"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   120
         MaxLength       =   4
         TabIndex        =   19
         Text            =   "123"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "123"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "123"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "123"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtShipProperty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "ABC"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   -73920
         TabIndex        =   41
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Hyper Speed"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -73920
         TabIndex        =   40
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Shield Incre."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -73920
         TabIndex        =   39
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Friction"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -73920
         TabIndex        =   35
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -73920
         TabIndex        =   34
         Top             =   2325
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Friction"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -73920
         TabIndex        =   32
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Accel."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -73920
         TabIndex        =   30
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Boost Accel."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -73920
         TabIndex        =   28
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Accel."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -73920
         TabIndex        =   26
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Class name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   23
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Max cloak"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   22
         Top             =   2325
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Max fuel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   20
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Max hull"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   18
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Max shield"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   1245
         Width           =   1335
      End
      Begin VB.Label lblShipProperty 
         BackStyle       =   0  'Transparent
         Caption         =   "Image index"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   885
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdRemoveShipType 
      Caption         =   "-"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Saves all values to the data files"
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton cmdAddShipType 
      Caption         =   "+"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Saves all values to the data files"
      Top             =   2520
      Width           =   255
   End
   Begin VB.ListBox lstShipTypes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Keeps values and returns to the main screen"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetDefaults 
      Caption         =   "&Set Defaults"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Saves all values to the data files"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Cancels any changes"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Create your own galaxy"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   225
      Left            =   720
      TabIndex        =   44
      Top             =   480
      Width           =   6405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   368
      X2              =   368
      Y1              =   96
      Y2              =   376
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning, adverse side effects may occur if invalid changes are made to variable values."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   6645
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   552
      X2              =   552
      Y1              =   96
      Y2              =   376
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   192
      X2              =   552
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   184
      X2              =   184
      Y1              =   96
      Y2              =   376
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   8
      X2              =   184
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Ships Editor"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Something else"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      BackColor       =   &H00EEFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Editor"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8565
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      FillColor       =   &H00404040&
      FillStyle       =   4  'Upward Diagonal
      Height          =   6735
      Left            =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReset_Click()

    mSounds.Play sndSelect
    
    LoadValues

End Sub

Private Sub Form_Load()
    
    Me.Caption = "About " & App.Title
    
    imgIcon.Picture = frmMainMenu.Icon
    
    LoadValues

End Sub

Private Sub cmdDone_Click()
    
    mSounds.Play sndProton
    
    SaveValues
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmOptions.Enabled = True

End Sub

Private Sub LoadValues()

   ' code here

End Sub

Private Sub SaveValues()

   ' code here

End Sub
