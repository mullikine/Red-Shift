Attribute VB_Name = "mEnvironment"
Option Explicit

Option Explicit

Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean


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
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


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


