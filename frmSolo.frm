VERSION 5.00
Begin VB.Form frmSolo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Single Player"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSolo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   746
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeSimSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   4073
      TabIndex        =   3
      Top             =   1440
      Width           =   7095
      Begin VB.ListBox lstShiptype 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmSolo.frx":08CA
         Left            =   3600
         List            =   "frmSolo.frx":08CC
         TabIndex        =   7
         Top             =   960
         Width           =   3135
      End
      Begin VB.ListBox lstSelectSim 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmSolo.frx":08CE
         Left            =   360
         List            =   "frmSolo.frx":08D0
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Header 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Player"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7125
      End
      Begin VB.Label cmdDelete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEL"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   5640
         Width           =   615
      End
      Begin VB.Shape shp 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   0
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label cmdNew 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   5640
         Width           =   615
      End
      Begin VB.Shape shp 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Index           =   1
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shiptype"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F7F7F7&
         BorderColor     =   &H00000080&
         Height          =   6135
         Left            =   0
         Top             =   0
         Width           =   7095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000040&
         Height          =   5775
         Left            =   15
         Top             =   360
         Width           =   7065
      End
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   2
      Left            =   8580
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   1
      Left            =   6660
      TabIndex        =   1
      ToolTipText     =   "Starts game with a randomly generated character"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   0
      Left            =   4740
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnMenu_Click(Index As Integer)

   mSounds.Play sndSelect
   
   Select Case Index
   Case 0
      PlayerSim = PlayerSims(PlayerSimByName(lstSelectSim.List(lstSelectSim.ListIndex)))
      mGame.RunGame
      
   Case 1
      With PlayerSim
         .Name = "Random dude"
         .System = Int(Rnd * (UBound(Systems) + 1))
         .ShipType = 2 'Int(Rnd * (UBound(ShipTypes) + 1))
      End With
      mGame.RunGame
      
   Case 2
      frmMainMenu.Show
      frmMainMenu.Enabled = True
      Me.Hide

   End Select

End Sub


Private Sub btnMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   For i = 0 To 2
      If Not i = Index Then
         btnMenu(i).ForeColor = &HFFC0C0
      End If
   Next i
   
   btnMenu(Index).ForeColor = &HFFF0F0

End Sub


Private Sub cmdDelete_Click()
   
   If lstSelectSim.ListIndex <> -1 Then ' if a pilot is selected
      DeletePlayerSim lstSelectSim.List(lstSelectSim.ListIndex)
   End If
   DisplayPlayerSimsToPlayerSimList
   
   btnMenu(0).Enabled = lstSelectSim.ListCount

End Sub

Private Sub cmdDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmdDelete
      .ForeColor = vbBlack
      shp(0).BackColor = vbWhite
      .Font.Bold = True
   End With

End Sub

Private Sub cmdDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmdDelete
      .ForeColor = vbWhite
      shp(0).BackColor = vbBlack
      .Font.Bold = False
   End With

End Sub

Private Sub cmdNew_Click()
   
   AddPlayerSim InputBox("Callsign:"), 1
   lstSelectSim.AddItem PlayerSims(nPlayerSims).Name
   DisplayPlayerSimsToPlayerSimList

End Sub

Private Sub cmdNew_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmdNew
      .ForeColor = vbBlack
      shp(1).BackColor = vbWhite
      .Font.Bold = True
   End With

End Sub

Private Sub cmdNew_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmdNew
      .ForeColor = vbWhite
      shp(1).BackColor = vbBlack
      .Font.Bold = False
   End With

End Sub

Private Sub Form_Load()

   SetTranslucent Me.hWnd, &H0, 230, LWA_ALPHA
   
   DisplayPlayerSimsToPlayerSimList

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   For i = 0 To 2
      btnMenu(i).ForeColor = &HFFC0C0
   Next i

End Sub


Sub DisplayPlayerSimsToPlayerSimList()

   lstSelectSim.Clear
   lstShiptype.Clear
   
   ' Add sims to the list
   If nPlayerSims <> -1 Then
      For i = 0 To nPlayerSims
         lstSelectSim.AddItem PlayerSims(i).Name
         lstShiptype.AddItem ShipTypes(PlayerSims(i).ShipType).ClassName
      Next i
   End If

End Sub


Private Sub lstSelectSim_Click()

   btnMenu(0).Enabled = (lstSelectSim.ListIndex <> -1 And Len(lstSelectSim.List(lstSelectSim.ListIndex)))
   lstShiptype.ListIndex = lstSelectSim.ListIndex

End Sub


Private Sub lstShiptype_Click()

   lstSelectSim.ListIndex = lstShiptype.ListIndex

End Sub
