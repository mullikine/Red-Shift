VERSION 5.00
Begin VB.Form frmDXForm 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Red Shift"
   ClientHeight    =   5010
   ClientLeft      =   5370
   ClientTop       =   3660
   ClientWidth     =   5040
   Icon            =   "frmDXForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDX 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmDXForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picDX_KeyDown(KeyCode As Integer, Shift As Integer)
        
   Keys(KeyCode) = True

End Sub

Private Sub picDX_KeyUp(KeyCode As Integer, Shift As Integer)

    Keys(KeyCode) = False

End Sub

Private Sub Form_Load()

   picDX.Width = 1024
   picDX.Height = 768
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   mCursor.x = Screen.Width / Screen.TwipsPerPixelX / 2
   mCursor.y = Screen.Height / Screen.TwipsPerPixelY / 2

End Sub

Private Sub picDX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   ' if in game
   If mGame.bRunning Then
      ' hyperspace select from map
      If mMap.HighlightedSystem <> -1 Then
         HyperspaceTo mMap.HighlightedSystem, You.Ship
         'mMap.MapUp = False
      End If
   End If

End Sub
