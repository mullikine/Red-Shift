Attribute VB_Name = "mScanner"
'---------------------------------------------------------------------------------------
' Module    : mScanner
' DateTime  : 3/25/2005 12:44
' Author    : Shane Mulligan
' Purpose   : Implements the game scanner
'---------------------------------------------------------------------------------------

Option Explicit

Public CompassX As Integer
Public CompassY As Integer
Public CompassRadius As Integer

Sub Init()

   CompassX = ScannerDims.x + ScannerDims.Width * (6 / 8)
   CompassY = ScannerDims.y + ScannerDims.Height * (6 / 8)
   CompassRadius = ScannerDims.Width * (2 / 8)

End Sub

Sub Draw()

   With Ships(You.Ship)
      
      EnableBlendColour
      
      DrawCompass
      DrawKinetics
   
   End With

End Sub

Private Sub DrawCompass()

   With Ships(You.Ship)
   
      DrawArrow InvertY(Atan(.x, .y)), White
      
      DrawArrow InvertY(.maBearing) + 1, Cyan
      DrawArrow InvertY(.maBearing), Cyan
      DrawArrow InvertY(.maBearing) - 1, Cyan
      
      DrawArrow InvertY(.maArg), Green
      
      If .CurrentShipSelection <> -1 Then DrawArrow PolarVectorVectorAddToArg(1, InvertY(CartToArg((.x - Ships(.CurrentShipSelection).x), (.y - Ships(.CurrentShipSelection).y))), -1, InvertY(.maArg)), Red
      
      'DrawCircle CompassX, CompassY, CompassRadius, &H80009050, , D3DPT_LINESTRIP
      
   End With

End Sub

Private Sub DrawKinetics()

   With Ships(You.Ship)
   
      DrawText "V " & Round(Ships(You.Ship).maMod, 0) & ",  " & Round(Ships(You.Ship).maBearing, 0), ScannerDims.x + 1, ScannerDims.y + ScannerDims.Height - 13, 12, 12, White, 0
      DrawText "X " & Round(.x, 0), ScannerDims.x + 1, ScannerDims.y + 1, 12, 12, White, 0
      DrawText "Y " & Round(.y), ScannerDims.x + 1, ScannerDims.y + 13, 12, 12, White, 0
      
   End With

End Sub

Private Sub DrawArrow(ByVal Bearing As Single, ByVal Colour As Long)
Dim TempVerts(1) As TLVERTEX
   
   TempVerts(0).x = CompassX
   TempVerts(0).y = CompassY
   TempVerts(0).tu = 1
   TempVerts(0).color = Colour
   
   TempVerts(1).x = CompassX + CompassRadius * Sin(ToRadians(Bearing))
   TempVerts(1).y = CompassY + CompassRadius * Cos(ToRadians(Bearing))
   TempVerts(1).tu = 1
   TempVerts(1).color = Colour
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, TempVerts(0), Len(TempVerts(0))

End Sub
