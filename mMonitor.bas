Attribute VB_Name = "mMonitor"
'---------------------------------------------------------------------------------------
' Module    : mMonitor
' DateTime  : 12/16/2004 21:23
' Author    : Shane Mulligan
' Purpose   : Monitors the application's performance
'---------------------------------------------------------------------------------------

Option Explicit

Public LastProcStartTime As Single

Public FrameRateAverageRange As Integer
Public ProcCount As Long

Public LastFPS As Integer
Private aLastFPS As Single


Sub FPS()

Static ticker As Single
Dim temp As Single

   temp = Timer - ticker
   
   ticker = Timer
   
   If temp <= 0 Then Exit Sub
   
   aLastFPS = CSng(aLastFPS * FrameRateAverageRange + 1 / temp) / (FrameRateAverageRange + 1)
   LastFPS = Int(aLastFPS)

End Sub
