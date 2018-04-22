Attribute VB_Name = "mSpaceDust"
'---------------------------------------------------------------------------------------
' Module    : mSpaceDust
' DateTime  : 2/11/2007 5:04
' Author    : Shane Mulligan
' Purpose   : Engine for single pixel particles (Very similar to explosions)
'---------------------------------------------------------------------------------------

Option Explicit

Private Type DustType
   Alive As Boolean
   system As Integer
   x As Single
   y As Single
   Colour As Long
   InitialColour As Long
   FinalColour As Long
   InitialTimeLeft As Integer
   TimeLeft As Integer
   i As Single ' dust moves by these values
   j As Single
End Type

Public Dusts() As DustType

Dim RefDust As Integer

Sub Draw()
EnableBlendNormal

Dim nVertices As Integer

Dim TempVerts() As TLVERTEX

   nVertices = -1
   For RefDust = 0 To UBound(Dusts)
      If Dusts(RefDust).Alive And Dusts(RefDust).system = Ships(You.Ship).system Then
         nVertices = nVertices + 1
         ReDim Preserve TempVerts(nVertices)
         With TempVerts(nVertices)
            
            .x = ZoomX(Dusts(RefDust).x + SpaceOffset.x)
            .y = ZoomY(-Dusts(RefDust).y + SpaceOffset.y)
            .color = Dusts(RefDust).Colour
            .rhw = 1
            .specular = 0
            .tu = 0
            .tv = 0
            
         End With
      End If
   Next RefDust
   
   If nVertices > -1 Then
      'ChangeTextureFactor Colour
      
      D3DDevice.SetTexture 0, ViewImages(0)
      D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, nVertices, TempVerts(0), Len(TempVerts(0))
   End If
    
End Sub

Sub DoPhysics()

   For RefDust = 0 To UBound(Dusts)
      With Dusts(RefDust)
      
         If .Alive Then
            .TimeLeft = .TimeLeft - 1
            .x = .x + .i
            .y = .y + .j
            .Alive = .TimeLeft >= 0
            If .system = Ships(You.Ship).system And .Alive Then
               .Colour = Blend(.InitialColour, .FinalColour, (.TimeLeft / .InitialTimeLeft))
            End If
         End If
       
      End With
   Next RefDust

End Sub

Sub MakeDust(ByVal x As Single, ByVal y As Single, ByVal system As Integer, ByVal time As Integer, Optional ByVal InitialColour As Long = &HFFFFFF, Optional ByVal FinalColour As Long = &H0, Optional ByVal i As Single = 0, Optional ByVal j As Single = 0)
Dim Selection As Integer

   Selection = -1
   For RefDust = 0 To UBound(Dusts)
      If Dusts(RefDust).Alive = False Then
         Selection = RefDust
         Exit For
      End If
   Next RefDust
   
   If Selection = -1 Then Exit Sub ' Dont bother
   
   With Dusts(Selection)
      .system = system
      .x = x
      .y = y
      .i = i
      .j = j
      .InitialColour = InitialColour
      .FinalColour = FinalColour
      .InitialTimeLeft = time
      .TimeLeft = time
      .Alive = True
   End With

End Sub

Sub MakeFlurry(ByVal x As Single, ByVal y As Single, ByVal system As Integer, ByVal time As Integer, ByVal intensity As Integer, Optional ByVal InitialColour As Long = &HFFFFFF, Optional ByVal FinalColour As Long = &H0)
   For i = 0 To intensity
      MakeDust x, y, system, time, InitialColour, FinalColour, Rnd * Sqr(intensity) - Sqr(intensity) / 2, Rnd * Sqr(intensity) - Sqr(intensity) / 2
   Next i
End Sub

Sub ChangeMax(ByVal nMax As Integer)

    ReDim Preserve Dusts(nMax - 1)

End Sub

Public Sub CleanUp()

    Erase ExplsnImages
    Erase Explosions

End Sub


  For i = 0 To intensity * SpaceZoom
      RandMod = ToRadians(Rnd * 360)
      MakeDust X, Y, system, time, InitialColour, FinalColour, Rnd * Sqr(intensity) * Cos(RandMod), Rnd * Sqr(intensity) * Sin(RandMod)
   Next i
End Sub

Sub ChangeMax(ByVal nMax As Integer)

    ReDim Preserve Dusts(nMax - 1)

End Sub

Public Sub CleanUp()

    Erase Dusts

End Sub


