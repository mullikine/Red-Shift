VERSION 5.00
Begin VB.Form frmGameMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Red Shift - Loading..."
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFC0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   746
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   Begin VB.Label lblMissionInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   10575
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   10215
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   10095
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   11055
      Left            =   240
      Top             =   240
      Width           =   14895
   End
End
Attribute VB_Name = "frmGameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

   lblInfo.Caption = "LOADING..."

End Sub

Private Sub Form_Load()
   
   Me.Caption = App.ProductName
   Open App.Path & "\Data\Config\Pregame Text.txt" For Input As #1
      lblMissionInfo.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
         String(2, vbCrLf) & _
         String(10, "-") & _
         String(2, vbCrLf) & _
         Input(LOF(1), #1)
   Close #1

End Sub

Sub PrintText(ByVal aText As String)

    lblInfo.Caption = lblInfo.Caption & vbCrLf & aText
    DoEvents

End Sub
