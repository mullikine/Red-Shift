VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Swflash.ocx"
Begin VB.Form frmAutorun 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Red  Shift"
   ClientHeight    =   4185
   ClientLeft      =   5115
   ClientTop       =   4230
   ClientWidth     =   5520
   Icon            =   "frmAutorun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Browes"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swfAuto 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      _cx             =   4204067
      _cy             =   4201739
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
      Stacking        =   "below"
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmAutorun.frx":08CA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5580
   End
End
Attribute VB_Name = "frmAutorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NoFlash As Boolean

Private Sub cmdBrowes_Click()

   sh32.Open App.Path

End Sub

Private Sub cmdExit_Click()

   End

End Sub

Private Sub cmdPlay_Click()

   Me.Hide
   LoadGame

End Sub

Private Sub Form_Activate()

   If NoFlash Then
      AskBox Me, "The flash movie 'Autorun.swf' was not found.", App.ProductName & " - Loading": NoFlash = False
   End If
   swfAuto.Play

End Sub

Private Sub Form_Load()

   SetTranslucent Me.hWnd, vbGreen, 245, LWA_COLORKEY Or LWA_ALPHA

On Error GoTo ErrHandler
   ' Load flash movie to object
   NoFlash = False
   swfAuto.Movie = App.Path & "\Data\Graphics\Shockwave Flash\Autorun.swf"
   swfAuto.Play
   swfAuto.Visible = True
   
   ' AnimateWindow(Me.hWnd, 200, AW_BLEND)
   If swfAuto.PercentLoaded = 0 Then GoTo ErrHandler
   
   Exit Sub
    
ErrHandler:
   NoFlash = True
   swfAuto.Stop
   swfAuto.Movie = ""
   swfAuto.Visible = False

End Sub
