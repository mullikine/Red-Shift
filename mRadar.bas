Attribute VB_Name = "mRadar"
'---------------------------------------------------------------------------------------
' Module    : mRadar
' DateTime  : 12/16/2004 21:20
' Author    : Shane Mulligan
' Purpose   : Implements the radar
'---------------------------------------------------------------------------------------

Option Explicit

Public RadarUp As Boolean


Sub Init()
 
   RadarZoom = 0.05
   SpaceZoom = 1
   RadarUp = True

End Sub


Sub DoDraw()

   If RadarUp Then Draw

End Sub


Sub Draw()
   
   ' Draw radar screen (greeny background and zoom text)
   'DrawTexture ViewImages(0), 0, 1, 0, 1, 0, 768, 0, 1024 - mStatusBar.Width, 0, False, &H1010, 0
   
   ' Draw thingies on radar
   
   EnableBlendOne
   
   mStellarObjects.DrawRadars
   mShips.DrawRadars

End Sub


Sub DrawToRadar(ByVal txrTexture As Direct3DBaseTexture8, ByVal pX As Long, ByVal pY As Long, ByVal pSize As Single, ByVal pColour As Long)
Const SmallestRadius As Integer = 3

   pSize = BoundMin(RadarZoom * pSize / 2, SmallestRadius)
   
   pX = ZoomXR(pX + SpaceOffset.x)
   pY = ZoomYR(-pY + SpaceOffset.y)
   
   ' check if need to draw
   If pX > MapDims.x And pX < MapDims.x + MapDims.Width And pY > MapDims.y And pY < MapDims.y + MapDims.Height Then
      DrawTexture txrTexture, srcRECTNorm, NewfRECT(pY - pSize, pY + pSize, pX - pSize, pX + pSize), , , pColour
      'DrawCircle pX, pY, pSize, pColour
   End If

End Sub

, pX - pSize, pX + pSize), , , pColour
      'DrawBasicCircle pX, pY, pSize, pColour
   End If

End Sub

peed * Timer) Mod 360) / 360) * 2 * PI), RelationColour(SelectRelations)
   
   For i = 0 To nCombatRadarPoints
      With CombatRadarPoints(i)
         If .Active Then
            DrawTexture txrFlares(1), srcRECTNorm, NewfRECT(-.Y + RadarSelDims.Y + RadarSelDims.Height / 2 - 4, -.Y + RadarSelDims.Y + RadarSelDims.Height / 2 + 4, .X + RadarSelDims.X + RadarSelDims.Width / 2 - 4, .X + RadarSelDims.X + RadarSelDims.Width / 2 + 4), 0, False, .Colour
         End If
      End With
   Next i

End Sub

Sub DoCRPhysics()
Dim i As Integer

   ' Erase Old Ones
   For i = 0 To nCombatRadarPoints
      If BetweenDegs(Int(ShipTypes(Ships(You.Ship).ShipType).RadarSpeed * Timer) Mod 360, (ShipTypes(Ships(You.Ship).ShipType).RadarSearchDegs + Int(ShipTypes(Ships(You.Ship).ShipType).RadarSpeed * Timer)) Mod 360, CartToArg(CombatRadarPoints(i).X, CombatRadarPoints(i).Y)) Then
         CombatRadarPoints(i).Active = False
      End If
   Next i
   
   ' Make New Ones
   For i = 0 To nShips
      If Ships(i).Alive And Not Ships(i).CloakOn Then
         If Ships(i).system = Ships(You.Ship).system Then
            If DistanceFromShipToShip(You.Ship, i) < ShipTypes(Ships(You.Ship).ShipType).RadarRange Then
               If BetweenDegs(Int(ShipTypes(Ships(You.Ship).ShipType).RadarSpeed * Timer) Mod 360, (ShipTypes(Ships(You.Ship).ShipType).RadarSearchDegs + Int(ShipTypes(Ships(You.Ship).ShipType).RadarSpeed * Timer)) Mod 360, DirectionToShip(i, You.Ship)) Then
                  NewCRPoint ((Ships(i).X - Ships(You.Ship).X) / ShipTypes(Ships(You.Ship).ShipType).RadarRange) * RadarSelDims.Width / 2, ((Ships(i).Y - Ships(You.Ship).Y) / ShipTypes(Ships(You.Ship).ShipType).RadarRange) * RadarSelDims.Height / 2, RelationColour(ShipRelations(You.Ship, i))
               End If
            End If
         End If
      End If
   Next i

End Sub

Sub NewCRPoint(ByVal X As Integer, ByVal Y As Integer, ByVal Colour As Long)
Dim Selection As Integer
Dim i As Integer

   Selection = -1
   For i = 0 To nCombatRadarPoints
      If CombatRadarPoints(i).Active = False Then
         Selection = i
         Exit For
      End If
   Next i
   
   If Selection = -1 Then ' if it cant find a place, make a new slot
      nCombatRadarPoints = nCombatRadarPoints + 1
      ReDim Preserve CombatRadarPoints(nCombatRadarPoints)
      Selection = nCombatRadarPoints
   End If
   
   With CombatRadarPoints(Selection)
      .X = X
      .Y = Y
      .Colour = Colour
      .Active = True
   End With

End Sub


Sub DoDraw()

   If RadarZoom > SpaceZoom Then RadarZoom = SpaceZoom
   If RadarUp Then Draw
   
   DoCRPhysics
   DrawCombatRadar

End Sub


Sub Draw()
   
   ' Draw radar screen (greeny background and zoom text)
   'DrawTexture ViewImages(0), 0, 1, 0, 1, 0, 768, 0, 1024 - mStatusBar.Width, 0, False, &H1010, 0
   
   ' Draw thingies on radar
   
   EnableBlendOne
   
   mStellarObjects.DrawRadars
   mShips.DrawRadars

End Sub


Sub DrawToRadar(ByVal txrTexture As Direct3DBaseTexture8, ByVal pX As Long, ByVal pY As Long, ByVal pSize As Single, ByVal pColour As Long)
Const SmallestRadius As Integer = 3

   pSize = BoundMin(RadarZoom * pSize / 2, SmallestRadius)
   
   pX = ZoomXR(pX + SpaceOffset.X)
   pY = ZoomYR(-pY + SpaceOffset.Y)
   
   ' check if need to draw
   If pX > ScreenDims.X - pSize / 2 And pX < ScreenDims.Width + pSize / 2 And pY > ScreenDims.Y - pSize / 2 And pY < ScreenDims.Height + pSize / 2 Then
      DrawTexture txrTexture, srcRECTNorm, NewfRECT(pY - pSize, pY + pSize, pX - pSize, pX + pSize), , , pColour
      'DrawBasicCircle pX, pY, pSize, pColour
   End If

End Sub

