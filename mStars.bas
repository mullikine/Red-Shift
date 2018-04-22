Attribute VB_Name = "mStars"
'---------------------------------------------------------------------------------------
' Module    : mStars
' DateTime  : 12/16/2004 21:19
' Author    : Shane Mulligan
' Purpose   : Implements the star background
'---------------------------------------------------------------------------------------

Option Explicit

Public BackColour As Long

Public StarCount As Integer
Public StarRadius As Integer
Dim RandomMargin As Single

Private Stars() As TLVERTEX
Private ScreenStars() As TLVERTEX

Sub Init()

    ReDim Stars(StarCount) As TLVERTEX
    ReDim ScreenStars(StarCount) As TLVERTEX
    StarRadius = 11000

End Sub

Sub InitialPhysics()
    
   For i = 0 To StarCount - 1
      With Stars(i)
      
         ' Evenly distributes the stars
         .x = Int(Rnd * 2 * StarRadius) - StarRadius
         .y = Int(Rnd * 2 * StarRadius) - StarRadius
         While Sqr(.x ^ 2 + .y ^ 2) > StarRadius
            .x = Int(Rnd * 2 * StarRadius) - StarRadius
            .y = Int(Rnd * 2 * StarRadius) - StarRadius
         Wend
         .x = .x + Ships(You.Ship).x
         .y = .y + Ships(You.Ship).y
         .color = NewGrade(Rnd * 255)
         .rhw = 1
         .specular = 0
         .z = 0
         
      End With
   Next i

End Sub

Sub Draw()
   
   EnableBlendNormal
   
   For i = 0 To StarCount - 1
      ScreenStars(i).x = ZoomX(Stars(i).x - Ships(You.Ship).x + StatusbarDims.x / 2)
      ScreenStars(i).y = ZoomY(-(Stars(i).y - Ships(You.Ship).y) + ScreenDims.Height / 2)
      ScreenStars(i).color = Stars(i).color
      ScreenStars(i).rhw = Stars(i).rhw
   Next i
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, StarCount, ScreenStars(0), Len(ScreenStars(0))

End Sub

Sub DoPhysics()
Dim randMargin As Single
Dim randArg As Single
   
   RandomMargin = Ships(You.Ship).maMod
   
   For i = 0 To StarCount - 1
      With Stars(i)
      
         randMargin = Rnd * RandomMargin
         randArg = Rnd * 180 - 90
         If Sqr((.x - Ships(You.Ship).x) ^ 2 + (.y - Ships(You.Ship).y) ^ 2) > StarRadius Then
            .x = Ships(You.Ship).x + PolToX(StarRadius - randMargin, Ships(You.Ship).maArg + randArg)
            .y = Ships(You.Ship).y + PolToY(StarRadius - randMargin, Ships(You.Ship).maArg + randArg)
         End If
         
      End With
   Next i

End Sub

Sub CleanUp()

    Erase Stars

End Sub

