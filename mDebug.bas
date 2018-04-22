Attribute VB_Name = "mDebug"
'---------------------------------------------------------------------------------------
' Module    : mDebug
' DateTime  : 12/16/2004 21:13
' Author    : Shane Mulligan
' Purpose   : In game debugging module
'---------------------------------------------------------------------------------------

Option Explicit

Public strPrint As String


Sub Draw()

   If Len(strPrint) Then DrawText "Debug: " & strPrint, 5, 630, 16, 16, White

End Sub


Sub sPrint(ByVal strString As String)

    strPrint = strString

End Sub
