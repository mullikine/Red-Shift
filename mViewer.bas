Attribute VB_Name = "mViewer"
'---------------------------------------------------------------------------------------
' Module    : mViewer
' DateTime  : 12/16/2004 21:21
' Author    : Shane Mulligan
' Purpose   : Defines the offsets used for drawing images to the screen at the correct locations
'---------------------------------------------------------------------------------------

Option Explicit

Public SpaceZoom As Single
Public RadarZoom As Single

Public SpaceCenter As Point
Public SpaceOffset As Point

Sub Init()

   SpaceCenter.x = StatusbarDims.x / 2
   SpaceCenter.y = ScreenDims.Height / 2

End Sub

Sub DoDraw()

   SetOffsets
   
   mStars.Draw
   mStellarObjects.DrawBodies
   mProjectiles.DrawBodies
   mShips.Draw
   mExplosions.Draw
   mTabSelect.DoDraw
   mRadar.DoDraw

End Sub

Sub SetOffsets()
   
   With Ships(You.Ship)
   
      SpaceZoom = Bound(SpaceZoom, BoundMax(BoundMax(ShipImageSets(ShipTypes(.ShipType).ShipImage).Size, 100) / ShipTypes(.ShipType).Size, 1), BoundMax((BoundMax(ShipImageSets(ShipTypes(.ShipType).ShipImage).Size, 100) / 2) / ShipTypes(.ShipType).Size, 1))
      RadarZoom = Bound(RadarZoom, BoundMax(BoundMax(ShipImageSets(ShipTypes(.ShipType).ShipImage).Size, 100) / ShipTypes(.ShipType).Size, 1), BoundMax((BoundMax(ShipImageSets(ShipTypes(.ShipType).ShipImage).Size, 100) / 6) / ShipTypes(.ShipType).Size, 1))
      
      SpaceOffset.x = SpaceCenter.x - .x
      SpaceOffset.y = SpaceCenter.y + .y
      
   End With

End Sub

Function ZoomMod(ByVal OriginalMod As Long) As Long

    ZoomMod = OriginalMod * SpaceZoom

End Function

Function ZoomX(ByVal OriginalX As Long) As Long

    ZoomX = OriginalX * SpaceZoom + SpaceCenter.x * (1 - SpaceZoom)

End Function

Function ZoomY(ByVal OriginalY As Long) As Long

    ZoomY = OriginalY * SpaceZoom + SpaceCenter.y * (1 - SpaceZoom)

End Function

Function ZoomXR(ByVal OriginalX As Long) As Long

    ZoomXR = OriginalX * RadarZoom + SpaceCenter.x * (1 - RadarZoom)

End Function

Function ZoomYR(ByVal OriginalY As Long) As Long

    ZoomYR = OriginalY * RadarZoom + SpaceCenter.y * (1 - RadarZoom)

End Function
om

End Function

Function ZoomXR(ByVal OriginalX As Long) As Long

    ZoomXR = OriginalX * RadarZoom + SpaceCenter.X * (1 - RadarZoom)

End Function

Function ZoomYR(ByVal OriginalY As Long) As Long

    ZoomYR = OriginalY * RadarZoom + SpaceCenter.Y * (1 - RadarZoom)

End Function
aceCenter.X * (1 - RadarZoom)

End Function

Function ZoomYR(ByVal OriginalY As Long) As Long

    ZoomYR = OriginalY * RadarZoom + SpaceCenter.Y * (1 - RadarZoom)

End Function
