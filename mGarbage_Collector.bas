Attribute VB_Name = "mGarbage_Collector"
'---------------------------------------------------------------------------------------
' Module    : mGarbage_Collector
' DateTime  : 12/16/2004 21:25
' Author    : Shane Mulligan
' Purpose   : Initializes the garbage collector
'---------------------------------------------------------------------------------------

Option Explicit

Sub CollectGarbage()
On Local Error Resume Next

    mStars.CleanUp
    mText.CleanUp
    mWallpapers.CleanUp
    mStatusBar.CleanUp
    mTextures.CleanUp
    mExplosions.CleanUp
    mDirectX.CleanUp

End Sub
