Attribute VB_Name = "mStatusBar"
'---------------------------------------------------------------------------------------
' Module    : mStatusBar
' DateTime  : 2/20/2005 10:51
' Author    : Shane Mulligan
' Purpose   : Procedures for the implementation of the status bar
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Draw()
   
   EnableBlendNormal

   DrawBack
   mScanner.Draw
   DrawMisc
   
   EnableBlendOne
   DrawBars
   DrawSelectBoxes

End Sub

Sub DrawBack()

   DrawXYWH tBack, StatusbarDims, White, False, True

End Sub

Sub DrawMisc()

   DrawText "View " & Round(SpaceZoom, 2) * 100, StatusbarDims.x + 10, StatusbarDims.Height - 60, 14, 14, White
   DrawText "Radar " & Round(RadarZoom, 2) * 100, StatusbarDims.x + 10, StatusbarDims.Height - 45, 14, 14, White

End Sub


Private Sub DrawBars()
Dim tmpDestXYWH As XYWH

   With ShipTypes(Ships(You.Ship).ShipType)
      
      ' Hull integrity and Shield remaining
      DrawXYWH tBar1, Bar1Dims, DarkGrey, , True
      tmpDestXYWH = MultXYWH(Bar1Dims, NewXYWH(1, 1, PositivePart(Ships(You.Ship).Hull) / .MaxHull, 1))
         DrawTexture tBarBase, SlideEffect(tmpDestXYWH, -6), XYWHTofRECT(tmpDestXYWH), , False, Blue
      tmpDestXYWH = MultXYWH(Bar1Dims, NewXYWH(1, 1, PositivePart(Ships(You.Ship).Shield) / .MaxShield, 1))
         DrawTexture txrFlares(1), Pattern(tmpDestXYWH, Right_Align), XYWHTofRECT(tmpDestXYWH), , False, Blue
      ' Cloak
      DrawXYWH tBar1, Bar2Dims, DarkGrey, , True
      DrawRectangle tBar0, Bar2Dims.x, Bar2Dims.y, Bar2Dims.Width * (PositivePart(Ships(You.Ship).Cloak) / .MaxCloak), Bar2Dims.Height, Yellow, , True
      ' Battery and fuel remaining
      DrawXYWH tBar1, Bar3Dims, DarkGrey, , True
      tmpDestXYWH = MultXYWH(Bar3Dims, NewXYWH(1, 1, PositivePart(Ships(You.Ship).Battery) / ShipTypes(Ships(You.Ship).ShipType).MaxBattery, 1))
         DrawTexture tBarBase, SlideEffect(tmpDestXYWH, -6), XYWHTofRECT(tmpDestXYWH), , False, Green
      tmpDestXYWH = MultXYWH(Bar3Dims, NewXYWH(1, 1, PositivePart(Ships(You.Ship).FuelLeft) / .MaxFuel, 1))
         DrawTexture txrFlares(1), Pattern(tmpDestXYWH, Right_Align), XYWHTofRECT(tmpDestXYWH), , False, Green
      'ViewImages (0)
      ' Hyperspace cruise distance remaining
      If Ships(You.Ship).InHyperspace Then
         DrawRectangle tBar1, 0, 768 - 15, StatusbarDims.x, 15, Blue, , True
         DrawRectangle tBar0, 0, 768 - 15, StatusbarDims.x * (PositivePart(Ships(You.Ship).HyperspaceCruiseDistanceLeft) / Ships(You.Ship).InitialHyperspaceCruiseDistanceLeft), 15, White, , True
      End If
      
   End With

End Sub

Sub DrawSelectBoxes()
Dim MinDim As Integer
Dim MaxDim As Integer
   
   MaxDim = ShipSelDims.Width

   If Ships(You.Ship).CurrentShipSelection = -1 Then
      DrawText Italics("No ship"), ShipSelDims.x + ShipSelDims.Width / 2 - Len("No ship") * 12 / 2, ShipSelDims.y + ShipSelDims.Height / 2 - 6, 12, 12, White, 1
   Else
      With ShipTypes(Ships(Ships(You.Ship).CurrentShipSelection).ShipType)
      
         DrawXYWH ViewImages(0), ShipSelDims, RelationColour(ShipRelations(You.Ship, Ships(You.Ship).CurrentShipSelection)), , False
         
         MinDim = 0
         If .Size < MaxDim Then
            MinDim = (MaxDim - .Size) / 2
            MaxDim = .Size
         End If
         EnableRadarBlend
         DrawTexture mShips.Image(Ships(Ships(You.Ship).CurrentShipSelection).ShipType, 5), srcRECTNorm, NewfRECT(ShipSelDims.y + MinDim, ShipSelDims.y + MaxDim + MinDim, ShipSelDims.x + MinDim, ShipSelDims.x + MaxDim + MinDim), 0, False, &HFF70FF70
         
         DrawText Governments(Ships(Ships(You.Ship).CurrentShipSelection).Government).Name, ShipSelDims.x + 1, ShipSelDims.y + 2, 10, 10, White
         
         DrawText .ClassName, ShipSelDims.x + 1, ShipSelDims.y + 14, 10, 10, White
         DrawText "Serial: " & Ships(You.Ship).CurrentShipSelection, ShipSelDims.x + 1, ShipSelDims.y + 26, 10, 10, White
         
         If Ships(Ships(You.Ship).CurrentShipSelection).System = Ships(You.Ship).System Then
            ' distance
            DrawText Int(DistanceFromShip(You.Ship, Ships(You.Ship).CurrentShipSelection)) & "m", ShipSelDims.x + 1, ShipSelDims.y + 38, 10, 10, White
            ' shield and hull status
            DrawTexture ViewImages(0), srcRECTNorm, NewfRECT(ShipSelDims.y + ShipSelDims.Height - 12, ShipSelDims.y + ShipSelDims.Height - 6, ShipSelDims.x, ShipSelDims.x + (ShipSelDims.Width) * (PositivePart(Ships(Ships(You.Ship).CurrentShipSelection).Shield) / .MaxShield)), 0, False, LightBlue
            DrawTexture ViewImages(0), srcRECTNorm, NewfRECT(ShipSelDims.y + ShipSelDims.Height - 6, ShipSelDims.y + ShipSelDims.Height, ShipSelDims.x, ShipSelDims.x + (ShipSelDims.Width) * (PositivePart(Ships(Ships(You.Ship).CurrentShipSelection).Hull) / .MaxHull)), 0, False, LightGreen
         Else
            DrawText "Out of range", ShipSelDims.x + 1, ShipSelDims.y + 38, 10, 10, White
         End If
         
         DrawText RelationsToString(ShipRelations(You.Ship, Ships(You.Ship).CurrentShipSelection)), ShipSelDims.x + 1, ShipSelDims.y + 50, 10, 10, White
         DrawText GetCombatRating(Ships(Ships(You.Ship).CurrentShipSelection).Kills), ShipSelDims.x + 1, ShipSelDims.y + 66, 10, 10, White
      
      End With
   End If
   
   MaxDim = SOSelDims.Width
   
   If Ships(You.Ship).CurrentStellarObjectSelection = -1 Then
      DrawText Italics("No object"), ShipSelDims.x + 5, SOSelDims.y + SOSelDims.Height / 2 - 7, 14, 14, White
   Else
      With StellarObjects(Ships(You.Ship).CurrentStellarObjectSelection)
         
         Select Case StellarObjects(Ships(You.Ship).CurrentStellarObjectSelection).Government
         Case -1
            DrawXYWH ViewImages(0), ShipSelDims, White, , False
         Case Else
            DrawXYWH ViewImages(0), ShipSelDims, RelationColour(GovRelations(StellarObjects(Ships(You.Ship).CurrentStellarObjectSelection).Government, Ships(You.Ship).Government)), , False
         End Select
         
         MinDim = 0
         If .Size < MaxDim Then
            MinDim = (MaxDim - .Size) / 2
            MaxDim = .Size
         End If
         
         EnableRadarBlend
         Select Case .Image
         Case -1
            ' draw nothing
         Case -2
            ' draw circle
            DrawCircle SOSelDims.x + (MaxDim + MinDim) / 2, SOSelDims.y + (MaxDim + MinDim) / 2, (MaxDim + MinDim) / 2, White, Int(PI * (MaxDim + MinDim) / 4), D3DPT_LINESTRIP
         Case Else
            ' draw texture
            DrawTexture StellarObjectImages(.Image), srcRECTNorm, NewfRECT(SOSelDims.y + MinDim, SOSelDims.y + MaxDim + MinDim, SOSelDims.x + MinDim, SOSelDims.x + MaxDim + MinDim), .Bearing, False, &HFF70FF70
         End Select
         
         Select Case StellarObjects(Ships(You.Ship).CurrentStellarObjectSelection).Government
         Case -1
            DrawText "No government", SOSelDims.x + 1, SOSelDims.y + 2, 10, 10, White
         Case Else
            DrawText Governments(StellarObjects(Ships(You.Ship).CurrentStellarObjectSelection).Government).Name, SOSelDims.x + 1, SOSelDims.y + 2, 10, 10, White
         End Select
         
         DrawText .Name, SOSelDims.x + 1, SOSelDims.y + 2, 12, 12, White
         DrawText "Serial: " & Ships(You.Ship).CurrentStellarObjectSelection, SOSelDims.x + 1, SOSelDims.y + 16, 10, 10, White
         'DrawText RelationsToString(ShipRelations(You.Ship, Ships(You.Ship).CurrentShipSelection)), ShipSelDims.X + 1, ShipSelDims.Y + 50, 10, 10, White
         
         DrawText Int(DistanceFromStellarObject(You.Ship, Ships(You.Ship).CurrentStellarObjectSelection)) & "m", SOSelDims.x + 1, SOSelDims.y + 28, 10, 10, White
         
      End With
   End If

End Sub


Sub CleanUp()

    Set tBack = Nothing
    Set tBar0 = Nothing
    Set tBar1 = Nothing

End Sub
