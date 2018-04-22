VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   746
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser bsrHelp 
      CausesValidation=   0   'False
      Height          =   7095
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   8295
      ExtentX         =   14631
      ExtentY         =   12515
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()

    frmMainMenu.Enabled = True
    Me.Hide

End Sub


Private Sub Form_Initialize()
    
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    cmdClose.Left = Me.ScaleWidth / 2 - cmdClose.Width / 2
    bsrHelp.Navigate App.Path & "\How to play.html"
    
    Form_Resize

End Sub

Private Sub Form_Load()

   SetTranslucent Me.hWnd, &H0, 230, LWA_ALPHA

End Sub

Private Sub Form_Resize()

    bsrHelp.Width = Me.ScaleWidth + 4
    bsrHelp.Height = Me.ScaleHeight + 4

End Sub
