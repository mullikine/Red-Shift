VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Swflash.ocx"
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Menu"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   746
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDebug 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SF1 
      Height          =   10935
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   15375
      _cx             =   4221424
      _cy             =   4213592
      Movie           =   "g"
      Src             =   "g"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "FFFFFF"
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00606060&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
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
      Index           =   4
      Left            =   10560
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00606060&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
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
      Index           =   3
      Left            =   8640
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00606060&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00606060&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label btnMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00606060&
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblNoFlash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Flash Here"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   15255
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NoFlash As Boolean


Private Sub btnMenu_Click(Index As Integer)

   mSounds.Play sndSelect
   SF1.Stop
   
   Select Case Index
   Case 0
      frmSolo.Show
      
   Case 1
      
      AskBox Me, "Not Implemented.", App.ProductName, vbOKOnly
      
   Case 2
      frmMainMenu.Enabled = False
      frmHelp.Show

   Case 3
      frmMainMenu.Enabled = False
      frmAbout.Show

   Case 4
      If AskBox(Me, "Are you sure you want to quit?", App.ProductName, vbYesNo) = vbYes Then mApplication.CloseApp

   End Select

End Sub


Private Sub btnMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   For i = 0 To 4
      If Not i = Index Then
         btnMenu(i).ForeColor = &HFFC0C0
      End If
   Next i
   
   btnMenu(Index).ForeColor = &HFFF0F0

End Sub

Private Sub Form_Load()
   
On Error GoTo ErrHandler
   ' Load flash movie to object
   NoFlash = False
   'SF1.LoadMovie 0, App.Path & "\Data\Graphics\Shockwave Flash\Startup.swf"
   SF1.Movie = App.Path & "\Data\Graphics\Shockwave Flash\Startup.swf"
   SF1.Play
   SF1.Visible = True
   
   ' AnimateWindow(Me.hWnd, 200, AW_BLEND)
   If SF1.PercentLoaded = 0 Then GoTo ErrHandler
   
   Exit Sub
    
ErrHandler:
   NoFlash = True
   SF1.Stop
   SF1.Movie = ""
   SF1.Visible = False

End Sub


Private Sub Form_Activate()

   If NoFlash Then
      AskBox Me, "The flash movie 'Startup.swf' was not found.", App.ProductName & " - Loading": NoFlash = False
   End If
   SF1.Play
   If bRunning Then frmDXForm.WindowState = FormWindowStateConstants.vbMaximized

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   For i = 0 To 2
      btnMenu(i).ForeColor = &HFFC0C0
   Next i

End Sub


Private Sub tmrDebug_Timer()
On Error Resume Next
   
   If frmSolo.Visible Then
      'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, mStars.BackColour, 1#, 0
      'Call ClearAndBeginRender: Call EndRender
     'ShowCursor True
   End If
   
   ' Idle music
   If Not DJ.Playing Then
      DJ.NewPlay mMusic.MenuMusic, -1
   End If

End Sub
