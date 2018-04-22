VERSION 5.00
Begin VB.Form frmAskBox 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Alert"
   ClientHeight    =   3135
   ClientLeft      =   4545
   ClientTop       =   3975
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmAskBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label cmdExitEx 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      X1              =   304
      X2              =   296
      Y1              =   8
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      X1              =   296
      X2              =   304
      Y1              =   8
      Y2              =   16
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   4350
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   315
   End
   Begin VB.Label cmd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "YES"
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
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label cmd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label cmd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
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
      Index           =   0
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shp 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Header 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Title"
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
      TabIndex        =   1
      Top             =   0
      Width           =   4725
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   4605
   End
End
Attribute VB_Name = "frmAskBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)

    AskBoxReturnVal = Index
    Select Case Index
    Case 0
        mSounds.Play sndNeutron
    Case 1, 2
        mSounds.Play sndProton
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmd(Index)
      .ForeColor = vbBlack
      shp(Index).BackColor = vbWhite
      .Font.Bold = True
   End With

End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

   With cmd(Index)
      .ForeColor = vbWhite
      shp(Index).BackColor = vbBlack
      .Font.Bold = False
   End With

End Sub

Private Sub cmdExitEx_Click()

   Me.Hide

End Sub

Private Sub cmdExitEx_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   Shape1.BackColor = &HDA&

End Sub

Private Sub cmdExitEx_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   Shape1.BackColor = &HC0&

End Sub

Private Sub Form_Load()

    lblMessage.Alignment = 2
    SetTranslucent Me.hWnd, &H0, 230, LWA_ALPHA

End Sub
