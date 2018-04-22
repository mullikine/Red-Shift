VERSION 5.00
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   6735
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   8535
   ClipControls    =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEditor 
      Caption         =   "&Launch"
      CausesValidation=   0   'False
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
      Left            =   5640
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Saves all values to the data files"
      Top             =   1560
      Width           =   1455
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
   Begin VB.TextBox txtStars 
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
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "123"
      Top             =   1680
      Width           =   855
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
      TabIndex        =   9
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
   Begin VB.CheckBox chkMIDI 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEFFFF&
      Caption         =   "MIDI Music"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkSFX 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEFFFF&
      Caption         =   "Game SFX"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Options"
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
      TabIndex        =   14
      Top             =   0
      Width           =   8565
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
      Height          =   345
      Left            =   120
      TabIndex        =   12
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
      Index           =   2
      X1              =   376
      X2              =   552
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   368
      X2              =   368
      Y1              =   96
      Y2              =   376
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   192
      X2              =   368
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "# stars"
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
      Left            =   3840
      TabIndex        =   11
      Top             =   1725
      Width           =   1575
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Editor"
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
      Index           =   2
      Left            =   5640
      TabIndex        =   8
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Setup"
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
      TabIndex        =   7
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Media"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "You can change game variables here"
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
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   6405
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
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEditor_Click()

   mSounds.Play sndSelect
   frmOptions.Enabled = False
   frmEditor.Show

End Sub

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

    frmMainMenu.Enabled = True

End Sub

Private Sub LoadValues()

   txtStars.Text = mStars.StarCount
   chkSFX.Value = -mSounds.bSFXon
   chkMIDI.Value = -mMusic.bMusicOn

End Sub

Private Sub SaveValues()

   mStars.StarCount = Val(txtStars.Text)
   mSounds.bSFXon = -chkSFX.Value
   mMusic.bMusicOn = -chkMIDI.Value

End Sub
