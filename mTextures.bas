Attribute VB_Name = "mTextures"
'---------------------------------------------------------------------------------------
' Module    : mTextures
' DateTime  : 12/16/2004 21:17
' Author    : Shane Mulligan
' Purpose   : Loads and holds flare textures
'---------------------------------------------------------------------------------------

Option Explicit

Public tBack As Direct3DTexture8, tBar0 As Direct3DTexture8, tBar1 As Direct3DTexture8, tBarBase As Direct3DTexture8
Public txrPlay256 As Direct3DTexture8, txrPlay64 As Direct3DTexture8, txrCircle256 As Direct3DTexture8

Public txrFlares(2) As Direct3DBaseTexture8
Public txrTitles(0) As Direct3DBaseTexture8


Sub Init()

   LoadShapeTextures
   LoadTitleTextures
   LoadProgressBarTextures
   LoadFontTextures
   LoadViewTextures

End Sub

Private Sub LoadShapeTextures()

   For i = 0 To UBound(txrFlares)
      Set txrFlares(i) = LoadTexture(App.Path _
            & "\Data\Graphics\Radar\" & i & ".tga", vbGreen)
   Next i

End Sub

Private Sub LoadTitleTextures()

   For i = 0 To UBound(txrTitles)
      Set txrTitles(i) = LoadTexture(App.Path _
            & "\Data\Graphics\Titles\" & i & ".jpg", vbGreen)
   Next i

End Sub

Private Sub LoadProgressBarTextures()

   Set txrPlay64 = LoadTexture(App.Path & "\Data\Graphics\Misc\Play64.bmp", -1)
   Set txrPlay256 = LoadTexture(App.Path & "\Data\Graphics\Misc\Play256.bmp", -1)
   Set txrCircle256 = LoadTexture(App.Path & "\Data\Graphics\Misc\Circle256.bmp", -1)
   
   Set tBar0 = LoadTexture(App.Path & "\Data\Graphics\StatusBar\Bar0.bmp", vbGreen)
   Set tBar1 = LoadTexture(App.Path & "\Data\Graphics\StatusBar\Bar1.bmp", vbGreen)
   Set tBarBase = LoadTexture(App.Path & "\Data\Graphics\StatusBar\BarBase.bmp", vbGreen)

End Sub

Private Sub LoadFontTextures()

   Set fntTex(0) = LoadTexture(App.Path & "\Data\Graphics\Fonts\0.bmp", vbGreen)
   Set fntTex(1) = LoadTexture(App.Path & "\Data\Graphics\Fonts\1.bmp", vbGreen)

End Sub

Private Sub LoadViewTextures()

   Set ViewImages(0) = LoadTexture(App.Path & "\Data\Graphics\View\White.bmp", vbGreen)
   Set ViewImages(1) = LoadTexture(App.Path & "\Data\Graphics\View\Circle.bmp", vbGreen)

End Sub

Sub LoadTextures()

   mWallpapers.Init
   
   ' Loads StellarObject images
   ReDim StellarObjectImages(0 To nStellarObjectImages) As Direct3DBaseTexture8
   For i = 0 To nStellarObjectImages
      Set StellarObjectImages(i) = LoadTexture(App.Path _
      & "\Data\Graphics\StellarObjects\" & CStr(i) & ".bmp", vbGreen)
      IncLoadStatus
   Next i
   
   LoadProjectileImageSets
   LoadShipImageSets
   
   ' Loads statusbar images
   Set tBack = LoadTexture(App.Path & "\Data\Graphics\StatusBar\Status Bar.bmp", vbGreen)
   IncLoadStatus
   
   ' Loads explosion image
   Set ExplsnImages(0) = LoadTexture(App.Path & "\Data\Graphics\Explosions\0.bmp", vbGreen)
   IncLoadStatus
   Set ExplsnImages(1) = LoadTexture(App.Path & "\Data\Graphics\Explosions\1.bmp", vbGreen)
   IncLoadStatus

End Sub

Private Sub LoadProjectileImageSets()
Dim aInput() As String
Dim GetColour As RGBA

   Open App.Path & "\Data\Graphics\Projectiles\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim ProjectileImageSets(aInput(1))
   For i = 0 To UBound(ProjectileImageSets)
      With ProjectileImageSets(i)
      
         Open App.Path & "\Data\Graphics\Projectiles\" & i & "\Specs.txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Frames = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .DeltaDegs = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .Extension = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            GetColour.Red = aInput(1)
            GetColour.Green = aInput(2)
            GetColour.Blue = aInput(3)
            .MaskColour = RGB(GetColour.Red, GetColour.Green, GetColour.Blue)
            aInput = Tokenize(ReadStr(1), " ")
            .FlipX = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .Size = aInput(1)
         Close #1
         ' Resizes the image array
         ReDim .Image(.Frames)
         ' Loads images
         For j = 0 To .Frames
            Set .Image(j) = LoadTexture(App.Path _
               & "\Data\Graphics\Projectiles\" & i _
               & "\Frames\" & j & .Extension, .MaskColour)
            IncLoadStatus
         Next j
         
      End With
   Next i

End Sub

Private Sub LoadShipImageSets()
Dim aInput() As String
Dim GetColour As RGBA

   Open App.Path & "\Data\Graphics\Ships\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim ShipImageSets(aInput(1))
   For i = 0 To UBound(ShipImageSets)
      With ShipImageSets(i)
      
         Open App.Path & "\Data\Graphics\Ships\" & i & "\Specs.txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Frames = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .DeltaDegs = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .Extension = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            GetColour.Red = aInput(1)
            GetColour.Green = aInput(2)
            GetColour.Blue = aInput(3)
            .MaskColour = RGB(GetColour.Red, GetColour.Green, GetColour.Blue)
            aInput = Tokenize(ReadStr(1), " ")
            .FlipX = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .Size = aInput(1)
         Close #1
         ' Resizes the image array
         ReDim .Image(.Frames)
         ' Loads images
         For j = 0 To .Frames
            Set .Image(j) = LoadTexture(App.Path _
               & "\Data\Graphics\Ships\" & i _
               & "\Frames\" & j & .Extension, .MaskColour)
            IncLoadStatus
         Next j
         
      End With
   Next i

End Sub

Sub CleanUp()

    Erase txrFlares

End Sub
dStatus
         Next j
         
      End With
   Next i

End Sub


Sub CleanUp()

    Erase txrFlares

End Sub
