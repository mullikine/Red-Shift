Attribute VB_Name = "mStatusMenu"
'---------------------------------------------------------------------------------------
' Module    : mStatusMenu
' DateTime  : 12/16/2004 21:18
' Author    : Shane Mulligan
' Purpose   : Implements the status menu when paused etc...
'---------------------------------------------------------------------------------------

Option Explicit

Public MenuUp As Boolean

Sub Init()

   mStatusMenu.MenuUp = False

End Sub

Sub DoDraw()
   
   If MenuUp Then
      Draw
   End If

End Sub


Private Sub Draw()

   With Ships(You.Ship)
   
      EnableBlendNormal
      
      DrawText "Player Name: " & PlayerSim.Name, 1, 10, 14, 14, Blue
      DrawText "Ship Type: " & ShipTypes(.ShipType).ClassName, 1, 26, 14, 14, Red
      DrawText "Current objectives:  " & Ships(You.Ship).ObjectiveType & Ships(You.Ship).ObjectiveIndex, 1, 42, 14, 14, White
      
      EnableBlendColour
      
      DrawRectangle ViewImages(0), StatusbarDims.x / 2 - Len("PAUSED") * 18 - 6, ScreenDims.Height / 2 - 18 - 6, Len("PAUSED") * 36 + 12, 48, Red, , True
      DrawText "PAUSED", StatusbarDims.x / 2 - Len("PAUSED") * 18, ScreenDims.Height / 2 - 18, 36, 36, Yellow
      
   End With

End Sub
End With

End Sub
itor.ProcCount, 1, 108, 14, 14, White
      DrawText "nProjectiles: " & nProjectiles, 1, 132, 14, 14, White
      DrawText "nShips: " & nShips & ", nShipRel: " & UBound(ShipRelations, 1), 1, 146, 14, 14, White
      DrawText "nGuns: " & nGuns, 1, 184, 14, 14, White
      
      'EnableBlendColour
      
      'DrawRectangle ViewImages(0), ScreenDims.Width / 2 - Len("PAUSED") * 18 - 6, ScreenDims.Height / 2 - 18 - 6, Len("PAUSED") * 36 + 12, 48, True, True, Blue, Red, Red, Blue
      'DrawText "PAUSED", ScreenDims.Width / 2 - Len("PAUSED") * 18, ScreenDims.Height / 2 - 18, 36, 36, Yellow
      
   End With

End Sub
