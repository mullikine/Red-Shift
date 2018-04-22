Attribute VB_Name = "mWindows"
Option Explicit

Public Const AW_HOR_POSITIVE = &H1
Public Const AW_HOR_NEGATIVE = &H2
Public Const AW_VER_POSITIVE = &H4
Public Const AW_VER_NEGATIVE = &H8
Public Const AW_CENTER = &H10
Public Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Public Const AW_ACTIVATE = &H20000
Public Const AW_SLIDE = &H40000
Public Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwflags As Long) As Boolean


Const GWL_EXSTYLE = (-20)
Const WS_EX_TRANSPARENT = &H20&
Const SWP_FRAMECHANGED = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWME = SWP_FRAMECHANGED Or _
   SWP_NOMOVE Or SWP_NOSIZE
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


' Topmost form code here
Function AlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Function FadeOutWindow(ByVal hWnd As Long, ByVal Period As Integer) As Boolean

   FadeOutWindow = AnimateWindow(hWnd, Period, AW_HIDE Or AW_BLEND)

End Function

Function FadeInWindow(ByVal hWnd As Long, ByVal Period As Integer) As Boolean

   FadeInWindow = AnimateWindow(hWnd, Period, AW_BLEND)

End Function

Sub SetTranslucent(ThehWnd As Long, color As Long, nTrans As Integer, flag As Byte)
    On Error GoTo ErrorRtn

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    'anything with color value color will completely disappear if flag = 1 or flag = 3
    SetLayeredWindowAttributes ThehWnd, color, nTrans, flag
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
    
End Sub

Sub RemoveTranslucent(ThehWnd As Long)
    On Error GoTo ErrorRtn

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Xor WS_EX_LAYERED
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
    
End Sub
