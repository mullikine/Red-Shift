Attribute VB_Name = "mWallpapers"
'---------------------------------------------------------------------------------------
' Module    : mWallpapers
' DateTime  : 12/16/2004 21:16
' Author    : Shane Mulligan
' Purpose   : Mainly just draws textures to screen
'---------------------------------------------------------------------------------------

Option Explicit

Private MaxWallpaper As Integer
Private Wallpaper() As Direct3DTexture8


Sub Init()

   MaxWallpaper = 4
   
   ReDim Wallpaper(0 To MaxWallpaper)
   For j = 0 To MaxWallpaper
      Set Wallpaper(j) = LoadTexture(App.Path _
         & "\Data\Graphics\Wallpapers\" & j & ".tga", 0)
      IncLoadStatus
   Next j

End Sub

Sub DrawStaticWallpaper(pWallpaper As eWallpaper, ByVal Colour As Long)
   
   ClearAndBeginRender
      DrawTexture Wallpaper(pWallpaper), srcRECTNorm, NewfRECT(-1, 768, -1, 1024), 0, False, Colour
   EndRender

End Sub

Sub DrawWallpaper(pWallpaper As eWallpaper, WallDims As XYWH, Optional ByVal Colour As Long = &HFFFFFF)

   With WallDims
   
      DrawTexture Wallpaper(pWallpaper), srcRECTNorm, NewfRECT(.y, .y + .Height, .x, .x + .Width), 0, False, Colour
   
   End With

End Sub

Sub CleanUp()

   Erase Wallpaper

End Sub

