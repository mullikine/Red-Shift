Attribute VB_Name = "mWinDims"
'---------------------------------------------------------------------------------------
' Module    : mWinDims
' DateTime  : 12/16/2004 17:18
' Author    : Shane Mulligan
' Purpose   : Consists of dimentions for various objects during gameplay
'---------------------------------------------------------------------------------------

Option Explicit

Public srcRECTNorm As fRECT

Public ScreenDims As XYWH
Public StatusbarDims As XYWH
Public Bar1Dims As XYWH
Public Bar2Dims As XYWH
Public Bar3Dims As XYWH
Public MapDims As XYWH
Public SOSelDims As XYWH
Public ShipSelDims As XYWH
Public LoadingDims As XYWH
Public ScannerDims As XYWH

Public Type XYWH                 ' Not defined in Win32API.txt
   x As Single                  '
   y As Single                  '
   Width As Single              '
   Height As Single             '
End Type

Public Type fRECT                 ' Not defined in Win32API.txt
   Top As Single                  '
   Bottom As Single                  '
   Left As Single              '
   Right As Single             '
End Type

Sub Init()

   srcRECTNorm = NewfRECT(0, 1, 0, 1)

   ScreenDims = NewXYWH(0, 0, 1024, 768)
   StatusbarDims = NewXYWH(ScreenDims.Width - 194, 0, 194, 768)
   Bar1Dims = NewXYWH(StatusbarDims.x + 35, 198, 148, 6)
   Bar2Dims = NewXYWH(StatusbarDims.x + 35, 215, 148, 6)
   Bar3Dims = NewXYWH(StatusbarDims.x + 35, 233, 148, 6)
   MapDims = NewXYWH(30, 30, 1024 - StatusbarDims.Width - 60, 768 - 60)
   SOSelDims = NewXYWH(1024 - StatusbarDims.Width + 5, StatusbarDims.Width + 180, 160, 160)
   ShipSelDims = NewXYWH(StatusbarDims.x + 9, 329, 174, 114)
   LoadingDims = NewXYWH(19, ScreenDims.Height - 70, ScreenDims.Width - 38, 30)
   ScannerDims = NewXYWH(StatusbarDims.x + 9, 9, 175, 175)

End Sub


Function NewRECT(ByVal Top As Long, ByVal Bottom As Long, ByVal Left As Long, ByVal Right As Long) As RECT

   NewRECT.Top = Top
   NewRECT.Bottom = Bottom
   NewRECT.Left = Left
   NewRECT.Right = Right

End Function

Function NewfRECT(ByVal Top As Single, ByVal Bottom As Single, ByVal Left As Single, ByVal Right As Single) As fRECT

   NewfRECT.Top = Top
   NewfRECT.Bottom = Bottom
   NewfRECT.Left = Left
   NewfRECT.Right = Right

End Function

Function NewXYWH(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As XYWH

   NewXYWH.x = x
   NewXYWH.y = y
   NewXYWH.Width = Width
   NewXYWH.Height = Height

End Function

Function XYWHTofRECT(ByRef uXYWH As XYWH) As fRECT

   XYWHTofRECT.Top = uXYWH.y
   XYWHTofRECT.Bottom = uXYWH.y + uXYWH.Height
   XYWHTofRECT.Left = uXYWH.x
   XYWHTofRECT.Right = uXYWH.x + uXYWH.Width

End Function

Function fRECTToXYWH(ByRef uRECT As fRECT) As XYWH

   fRECTToXYWH.y = uRECT.Top
   fRECTToXYWH.Height = uRECT.Bottom - uRECT.Top
   fRECTToXYWH.x = uRECT.Left
   fRECTToXYWH.Width = uRECT.Right - uRECT.Left

End Function

Function fRECTToRECT(ByRef uRECT As fRECT) As RECT

   fRECTToRECT.Top = uRECT.Top
   fRECTToRECT.Bottom = uRECT.Bottom
   fRECTToRECT.Left = uRECT.Left
   fRECTToRECT.Right = uRECT.Right

End Function

Function RECTTofRECT(ByRef uRECT As RECT) As fRECT

   RECTTofRECT.Top = uRECT.Top
   RECTTofRECT.Bottom = uRECT.Bottom
   RECTTofRECT.Left = uRECT.Left
   RECTTofRECT.Right = uRECT.Right

End Function

Function MultXYWH(ByRef pXYWH1 As XYWH, ByRef pXYWH2 As XYWH) As XYWH

   With MultXYWH
   
      .x = pXYWH1.x * pXYWH2.x
      .y = pXYWH1.y * pXYWH2.y
      .Width = pXYWH1.Width * pXYWH2.Width
      .Height = pXYWH1.Height * pXYWH2.Height
      
   End With

End Function

Function AddXYWH(ByRef pXYWH1 As XYWH, ByRef pXYWH2 As XYWH) As XYWH

   With AddXYWH
   
      .x = pXYWH1.x + pXYWH2.x
      .y = pXYWH1.y + pXYWH2.y
      .Width = pXYWH1.Width + pXYWH2.Width
      .Height = pXYWH1.Height + pXYWH2.Height
      
   End With

End Function

Function MultfRECT(ByRef pRECT1 As fRECT, ByRef pRECT2 As fRECT) As fRECT

   With MultfRECT
   
      .Bottom = pRECT1.Bottom * pRECT2.Bottom
      .Left = pRECT1.Left * pRECT2.Left
      .Right = pRECT1.Right * pRECT2.Right
      .Top = pRECT1.Top * pRECT2.Top
      
   End With

End Function

Function AddfRECT(ByRef pRECT1 As fRECT, ByRef pRECT2 As fRECT) As fRECT

   With AddfRECT
   
      .Bottom = pRECT1.Bottom + pRECT2.Bottom
      .Left = pRECT1.Left + pRECT2.Left
      .Right = pRECT1.Right + pRECT2.Right
      .Top = pRECT1.Top + pRECT2.Top
      
   End With

End Function
