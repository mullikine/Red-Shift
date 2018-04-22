Attribute VB_Name = "mCursor"
Option Explicit

' Visibility
'Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Dim pVisible As Boolean

' Position
Private Declare Sub SetCursorPos Lib "user32" (ByVal x As Integer, ByVal y As Integer)
Private Declare Function GetCursorPos Lib "USER32.DLL" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

Dim Pos As POINTAPI
Dim ReturnValue As Long

Property Get x() As Integer
   
   ReturnValue = GetCursorPos(Pos)
   x = CInt(Pos.x)

End Property

Property Let x(ByVal Value As Integer)
   
   Pos.x = CLng(Value)
   SetCursorPos CLng(Pos.x), CLng(Pos.y)

End Property

Property Get y() As Integer
   
   ReturnValue = GetCursorPos(Pos)
   y = CInt(Pos.y)

End Property

Property Let y(ByVal Value As Integer)
   
   Pos.y = CLng(Value)
   SetCursorPos CLng(Pos.x), CLng(Pos.y)

End Property


'Property Get Visible() As Boolean
'
'   Visible = pVisible
'
'End Property

'Property Let Visible(ByVal Value As Boolean)
'
'   pVisible = Value
'   RefreshCursor
'
'End Property

'Private Sub RefreshCursor()
'
'   While ShowCursor(pVisible) = False
'      DoEvents
'   Wend
'
'End Sub

Sub Draw()

   EnableBlendNormal
   
   DrawLine mCursor.x, mCursor.y - 10, mCursor.x, mCursor.y + 10
   DrawLine mCursor.x - 10, mCursor.y, mCursor.x + 10, mCursor.y

End Sub


Private Sub Class_Initialize()

   ReturnValue = GetCursorPos(Pos)
   
   'pVisible = True
   'RefreshCursor

End Sub

Time - LatentTime) / FadeTime, 1, 0))
   'DrawBasicCircle X, Y, 5, LightRed, 16, D3DPT_LINESTRIP

End Sub


Sub Init()

   ReturnValue = GetCursorPos(Pos)
   LatentTime = 2
   FadeTime = 1
   LastMoveTime = Timer
   
   'pVisible = True
   'RefreshCursor

End Sub

