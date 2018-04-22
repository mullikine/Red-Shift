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

