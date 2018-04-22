Attribute VB_Name = "mRes"
'---------------------------------------------------------------------------------------
' Module    : mRes
' DateTime  : 1/22/2005 13:26
' Author    : Shane Mulligan
' Purpose   : Loads data from the resource files etc...
'---------------------------------------------------------------------------------------

Option Explicit

' n ---------------------------
Public nStellarObjectImages As Integer
Public nShipImages As Integer
Public nProjectileImages As Integer
' textures --------------------
Public StellarObjectImages() As Direct3DBaseTexture8
Public ShipImageSets() As tAniSprite
Public ProjectileImageSets() As tAniSprite
Public ViewImages(1) As Direct3DBaseTexture8


Type tAniSprite
   Frames As Integer
   DeltaDegs As Integer
   Image() As Direct3DBaseTexture8
   Extension As String
   MaskColour As Long
   FlipX As Boolean
   Size As Integer
End Type


Sub Init()

   mOptions.Init
   'mOptions.LoadPrefs NormalGamePrefs
   mOptions.LoadPrefs GetPrefsFromFile(App.Path & "\Data\Config\Prefs.txt")
   mInitTypes.Init
   mInitStaticTypes.Init

End Sub
nitStaticTypes.Init

End Sub
