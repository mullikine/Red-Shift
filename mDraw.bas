Attribute VB_Name = "mDraw"
'---------------------------------------------------------------------------------------
' Module    : mPaintbrush
' DateTime  : 12/16/2004 21:21
' Author    : Shane Mulligan
' Purpose   : Contains basic drawing functions
'---------------------------------------------------------------------------------------

Option Explicit

'This is the Flexible-Vertex-Format description for a 2D vertex (Transformed and Lit)
Public Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'This structure describes a transformed and lit vertex - it's identical to the DirectX7 type "D3DTLVERTEX"
Public Type TLVERTEX
   x As Single
   y As Single
   z As Single
   rhw As Single
   color As Long
   specular As Long
   tu As Single
   tv As Single
End Type

Function NewTLVERTEX(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal rhw As Single, ByVal Colour As Long, ByVal specular As Long, ByVal tu As Single, ByVal tv As Single) As TLVERTEX

   With NewTLVERTEX
      .color = Colour
      .rhw = rhw
      .specular = specular
      .tu = tu
      .tv = tv
      .x = x
      .y = y
      .z = z
   End With

End Function

Sub DrawTexture(ByVal Texture As Direct3DTexture8, ByRef srcRECT As fRECT, ByRef destRECT As fRECT, Optional ByVal Angle As Single = 0, Optional bSpaceBody As Boolean = False, Optional Colour As Long = &HFFFFFFFF)
Dim TempVerts(4) As TLVERTEX
   ' triangle fan, fast and easy
   ' 0----1
   ' |\   |
   ' | \  |
   ' |  \ |
   ' |   \|
   ' 3----2
Dim HalfImageHeight As Single
Dim HalfImageWidth As Single

   HalfImageHeight = Abs(destRECT.Bottom - destRECT.Top) \ 2
   HalfImageWidth = Abs(destRECT.Right - destRECT.Left) \ 2
   
   If bSpaceBody Then
      HalfImageWidth = ZoomMod(HalfImageWidth)
      HalfImageHeight = ZoomMod(HalfImageHeight)
      'Angle = InvertY(Angle)
      Angle = Angle + 180
      srcRECT.Left = 1 - srcRECT.Left
      srcRECT.Right = 1 - srcRECT.Right
      destRECT.Top = ZoomY(-destRECT.Bottom + SpaceOffset.y)
      destRECT.Bottom = ZoomY(-destRECT.Top + SpaceOffset.y)
      destRECT.Left = ZoomX(destRECT.Left + SpaceOffset.x)
      destRECT.Right = ZoomX(destRECT.Right + SpaceOffset.x)
   End If
   
   destRECT.Left = destRECT.Left + 1
   'destRight = destRight + 1
   
   'x = xCosq - ySinq
   'y = xSinq + yCosq

   TempVerts(0) = NewTLVERTEX(destRECT.Left + HalfImageWidth + RotateX(Angle, -HalfImageWidth, -HalfImageHeight), destRECT.Top + HalfImageHeight + RotateY(Angle, -HalfImageWidth, -HalfImageHeight), 0, 1, Colour, 0, srcRECT.Left, srcRECT.Top)
   TempVerts(1) = NewTLVERTEX(destRECT.Left + HalfImageWidth + RotateX(Angle, HalfImageWidth, -HalfImageHeight), destRECT.Top + HalfImageHeight + RotateY(Angle, HalfImageWidth, -HalfImageHeight), 0, 1, Colour, 0, srcRECT.Right, srcRECT.Top)
   TempVerts(2) = NewTLVERTEX(destRECT.Left + HalfImageWidth + RotateX(Angle, HalfImageWidth, HalfImageHeight), destRECT.Top + HalfImageHeight + RotateY(Angle, HalfImageWidth, HalfImageHeight), 0, 1, Colour, 0, srcRECT.Right, srcRECT.Bottom)
   TempVerts(3) = NewTLVERTEX(destRECT.Left + HalfImageWidth + RotateX(Angle, -HalfImageWidth, HalfImageHeight), destRECT.Top + HalfImageHeight + RotateY(Angle, -HalfImageWidth, HalfImageHeight), 0, 1, Colour, 0, srcRECT.Left, srcRECT.Bottom)
   TempVerts(4) = TempVerts(0)
   
   ChangeTextureFactor Colour
   
   D3DDevice.SetTexture 0, Texture
   D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, TempVerts(0), Len(TempVerts(0))

End Sub

Sub DrawCircle(ByVal x As Integer, ByVal y As Integer, ByVal Radius As Integer, Optional ByVal Colour As Long = &HFFFFFFFF, Optional ByVal nVertices As Integer = 60, Optional ByVal Method As CONST_D3DPRIMITIVETYPE = D3DPT_LINESTRIP)
ReDim TempVerts(nVertices) As TLVERTEX

   For i = 0 To nVertices
      With TempVerts(i)
         
         .x = x + Int(Radius * Sin((i / (nVertices / 2)) * PI))
         .y = y + Int(Radius * Cos((i / (nVertices / 2)) * PI))
         .color = Colour
         .rhw = 1
         .specular = 0
         
      End With
   Next i
   
   ChangeTextureFactor Colour
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP Method, nVertices, TempVerts(0), Len(TempVerts(0))

End Sub

Sub DrawVector(ByVal x As Integer, ByVal y As Integer, ByVal Distance As Integer, ByVal Bearing As Integer, Optional Colour As Long = &HFFFFFFFF, Optional ColourSecondary As Long = -1)

   DrawLine x, y, x + Distance * Sin(ToRadians(Bearing)), y + Distance * Cos(ToRadians(Bearing)), Colour, ColourSecondary

End Sub

Sub DrawLine(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, Optional Colour As Long = &HFFFFFFFF, Optional ColourSecondary As Long = -1)
ReDim TempVerts(1) As TLVERTEX

   If ColourSecondary = -1 Then ColourSecondary = Colour

   With TempVerts(0)
      
      .x = x1
      .y = y1
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 0
      
   End With
   
   With TempVerts(1)
      
      .x = x2
      .y = y2
      .color = ColourSecondary
      .rhw = 1
      .specular = 0
      .tu = 1
      .tv = 1
      
   End With
   
   ChangeTextureFactor Colour
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, TempVerts(0), Len(TempVerts(0))

End Sub

Sub DrawXYWH(ByVal Texture As Direct3DTexture8, ByRef pXYWH As XYWH, Optional Colour As Long = &HFFFFFFFF, Optional Linewidth As Integer = 1, Optional Filled As Boolean = False)

   With pXYWH
   
      DrawRectangle Texture, .x, .y, .Width, .Height, Colour, Linewidth, Filled
      
   End With

End Sub

Sub DrawfRECT(ByVal Texture As Direct3DTexture8, ByRef pRECT As fRECT, Optional Colour As Long = &HFFFFFFFF, Optional Linewidth As Integer = 1, Optional Filled As Boolean = False)

   With pRECT
   
      DrawXYWH Texture, fRECTToXYWH(pRECT), Colour, Linewidth, Filled
      
   End With

End Sub

Sub DrawRectangle(ByVal Texture As Direct3DTexture8, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional Colour As Long = &HFFFFFFFF, Optional Linewidth As Integer = 1, Optional Filled As Boolean = False)
ReDim TempVerts(4) As TLVERTEX

   With TempVerts(0)
      
      .x = x
      .y = y
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 0
      
   End With
   
   With TempVerts(1)
      
      .x = x + Width
      .y = y
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 1
      .tv = 0
      
   End With
   
   With TempVerts(2)
      
      .x = x + Width
      .y = y + Height
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 1
      .tv = 1
      
   End With
   
   With TempVerts(3)
      
      .x = x
      .y = y + Height
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 1
      
   End With
   
   With TempVerts(4)
      
      .x = x
      .y = y
      .color = Colour
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 0
      
   End With

   
   D3DDevice.SetTexture 0, Texture
   If Filled Then
      D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, TempVerts(0), Len(TempVerts(0))
   Else
      D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, TempVerts(0), Len(TempVerts(0))
   End If

End Sub

Function SlideEffect(ByRef pDestXYWH As XYWH, ByVal PixelsPerSecond As Single) As fRECT
Dim fOffset As Single

   With pDestXYWH
   
      fOffset = (Timer * PixelsPerSecond Mod .Height) / .Height
   
   End With
   
   With SlideEffect
   
      .Top = 0
      .Bottom = 1
      .Left = 0 + fOffset
      .Right = pDestXYWH.Width / pDestXYWH.Height + fOffset
      
   End With

End Function

Function Pattern(ByRef pDestXYWH As XYWH, ByVal Align As eAlign) As fRECT
   
   With Pattern
   
      .Top = 0
      .Bottom = 1
      Select Case Align
      Case Left_Align
         .Left = 0
         .Right = pDestXYWH.Width / pDestXYWH.Height
      Case Right_Align
         .Left = -pDestXYWH.Width / pDestXYWH.Height
         .Right = 0
      Case Else
         .Left = 0
         .Right = 1
      End Select
      
   End With

End Function

Function LoadTexture(SrcFile As String, Transparent_Colour As Long) As Direct3DTexture8
    
    Set LoadTexture = D3DX.CreateTextureFromFileEx(D3DDevice, SrcFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                                            D3DX_DEFAULT, 0, DispMode.Format, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, Transparent_Colour, _
                                                                            ByVal 0, ByVal 0)
End Function 'D3DFMT_UNKNOWN

Sub ClearAndBeginRender()

   ' clear the device
   D3DDevice.Clear 1, ByVal 0, D3DCLEAR_TARGET, mStars.BackColour, 1#, 0
   
   ' call begin scene
   D3DDevice.BeginScene

End Sub

Sub EndRender(Optional hWndOveride As Long)

   ' end the scene
   D3DDevice.EndScene
   
   ' present the backbuffer to the screen
   If frmDXForm.Visible Then D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub EnableBlendOne()
   
   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
       
End Sub

Public Sub EnableBlendNormal()
   
   If AlphaMode Then
      EnableBlendColour
   Else
      D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
      D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
      
      D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
   End If

End Sub

Public Sub EnableBlendColour()

   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

End Sub

Public Sub EnableBlendInvColour()
   
   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVSRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVDESTCOLOR
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
       
End Sub

Sub EnableRadarBlend()

   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
   
End Sub

Sub ChangeTextureFactor(ByVal Colour As Long)

   If AlphaMode Then
      D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, Colour
      D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
   End If
   
End Sub

Sub DisableTextureFactor()

   'D3DDevice.Reset D3DWindow
       
End Sub

Public Sub DisableBlend()
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
           
End Sub

Sub CleanUp()

   Call ClearAndBeginRender
   Call EndRender

End Sub
 .tu = 1
      .tv = 0
      
   End With
   
   With TempVerts(2)
      
      .x = x + Width
      .y = y + Height
      .color = ColourBR
      .rhw = 1
      .specular = 0
      .tu = 1
      .tv = 1
      
   End With
   
   With TempVerts(3)
      
      .x = x
      .y = y + Height
      .color = ColourBL
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 1
      
   End With
   
   With TempVerts(4)
      
      .x = x
      .y = y
      .color = ColourTL
      .rhw = 1
      .specular = 0
      .tu = 0
      .tv = 0
      
   End With
   
   If Filled Then
      D3DDevice.SetTexture 0, Texture
      D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 2, TempVerts(0), Len(TempVerts(0))
   End If
   
   If Border Then
      D3DDevice.SetTexture 0, ViewImages(0)
      D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, TempVerts(0), Len(TempVerts(0))
   End If

End Sub

Function SlideEffect(ByRef pDestXYWH As XYWH, ByVal PixelsPerSecond As Single) As fRECT
Dim fOffset As Single

   With pDestXYWH
   
      fOffset = (Timer * PixelsPerSecond Mod .Height) / .Height
   
   End With
   
   With SlideEffect
   
      .Top = 0
      .Bottom = 1
      .Left = 0 + fOffset
      .Right = pDestXYWH.Width / pDestXYWH.Height + fOffset
      
   End With

End Function

Function Pattern(ByRef pDestXYWH As XYWH, ByVal Align As eAlign) As fRECT
   
   With Pattern
   
      .Top = 0
      .Bottom = 1
      Select Case Align
      Case Left_Align
         .Left = 0
         .Right = pDestXYWH.Width / pDestXYWH.Height
      Case Right_Align
         .Left = -pDestXYWH.Width / pDestXYWH.Height
         .Right = 0
      Case Else
         .Left = 0
         .Right = 1
      End Select
      
   End With

End Function

Function LoadTexture(SrcFile As String, Transparent_Colour As Long) As Direct3DTexture8
    
    Set LoadTexture = D3DX.CreateTextureFromFileEx(D3DDevice, SrcFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                                            D3DX_DEFAULT, 0, DispMode.Format, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, Transparent_Colour, _
                                                                            ByVal 0, ByVal 0)
End Function 'D3DFMT_UNKNOWN

Sub ClearAndBeginRender()

   ' clear the device
   D3DDevice.Clear 1, ByVal 0, D3DCLEAR_TARGET, mStars.BackColour, 1#, 0
   
   ' call begin scene
   D3DDevice.BeginScene

End Sub

Sub EndRender(Optional hWndOveride As Long)

   ' end the scene
   D3DDevice.EndScene
   
   ' present the backbuffer to the screen
   If frmDXForm.Visible Then D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Sub EnableBlendOne()
   
   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
       
End Sub

Public Sub EnableBlendNormal()
   
   If AlphaMode Then
      EnableBlendColour
   Else
      D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
      D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
      
      D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
   End If

End Sub

Public Sub EnableBlendColour()

   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

End Sub

Public Sub EnableBlendInvColour()
   
   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVSRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVDESTCOLOR
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
       
End Sub

Sub EnableRadarBlend()

   D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
   D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
   
End Sub

Sub ChangeTextureFactor(ByVal Colour As Long)

   If AlphaMode Then
      D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, Colour
      D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
   End If
   
End Sub

Sub DisableTextureFactor()

   'D3DDevice.Reset D3DWindow
       
End Sub

Public Sub DisableBlend()
   
   D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
           
End Sub

Sub CleanUp()

   Call ClearAndBeginRender
   Call EndRender

End Sub
