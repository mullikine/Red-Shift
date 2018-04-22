Attribute VB_Name = "mGameMenu"
Option Explicit

Sub Show()
' c - iterator, l - location on screen, i - current selected index
Dim c As Integer, l As Integer, i As Integer
Const ShipAngleN As Integer = 90
Dim ShipAngle As Integer
Const ShipSize As Integer = 50
Const Indent As Integer = 20
Const LatentAngle As Integer = 50
Dim sTempString As String

   While bMenu And DoEvents()
   
      ClearAndBeginRender
      EnableBlendNormal
      DrawTexture txrTitles(0), srcRECTNorm, NewfRECT(50, 180, (ScreenDims.Width - 700) / 2, (ScreenDims.Width + 700) / 2)
      
      i = iMenuSelectedShip
      For c = 0 To UBound(PlayerSims)
         l = c - i
         ShipAngle = ShipAngleN + Bound((c - i) / 5, 1, -1) * LatentAngle
         With PlayerSims(c)
            
            DrawTexture mShips.Image(.ShipType, ShipAngle), srcRECTNorm, XYWHTofRECT(NewXYWH(100 - Bound(Abs(c - i) / 5, 1, 0) * Indent, ScreenDims.Height / 2 + ShipSize * l, ShipSize * (1 - Bound(Abs(c - i) / 5, 1, 0)), ShipSize * (1 - Bound(Abs(c - i) / 5, 1, 0)))), ShipAngle - Round(ShipAngle / ShipImageSets(ShipTypes(.ShipType).ShipImage).DeltaDegs, 0) * ShipImageSets(ShipTypes(.ShipType).ShipImage).DeltaDegs, False, Blend(0, White, Bound(Abs(c - i) / 5, 1, 0))
            DrawText ShipTypes(.ShipType).ClassName, 100 + ShipSize - Bound(Abs(c - i) / 5, 1, 0) * Indent, ScreenDims.Height / 2 + l * ShipSize + (ShipSize - 12) / 2, 12, 12, Blend(0, Red, Bound(Abs(c - i) / 5, 1, 0)), 1
            
            If c = i Then 'current selection
               sTempString = Systems(.system).Name
               DrawText sTempString, ScreenDims.Width - Len(sTempString) * 20 - 100, (ScreenDims.Height - 20) / 2, 20, 20, White, 1
               sTempString = "(" & .X & ", " & .Y & ")"
               DrawText sTempString, ScreenDims.Width - Len(sTempString) * 16 - 100, (ScreenDims.Height - 20) / 2 + 20, 16, 16, LightRed, 1
            End If
         
         End With
      Next c
      
      EndRender 'endscene and present
      
      mKeyboard.DoMenuKeys
      
   Wend

End Sub
