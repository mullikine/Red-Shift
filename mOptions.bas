Attribute VB_Name = "mOptions"
'---------------------------------------------------------------------------------------
' Module    : mOptions
' DateTime  : 2/20/2005 10:50
' Author    : Shane Mulligan
' Purpose   : Allows quick swapping of important game variables
'---------------------------------------------------------------------------------------

Option Explicit

Public NormalGamePrefs As tGamePrefs

Public Type tGamePrefs
   nStellarObjectImages As Integer
   iMaxExplosions As Integer
   nBGStars As Integer
   iBGStarsMargin As Integer
   iProcsPerFrame As Integer
   iMinFPS As Integer
   iMaxFPS As Integer
   iFrameRateAverageRange As Integer
   bSoundsOn As Boolean
   bMusicOn As Boolean
   iMaxShips As Integer
   bInvertSpeakers As Boolean
End Type


Sub Init()

   With NormalGamePrefs
   
      .bMusicOn = True
      .bSoundsOn = True
      .iBGStarsMargin = 200
      .iFrameRateAverageRange = 50
      .iMaxFPS = 250
      .iMinFPS = 50
      .iProcsPerFrame = 1
      .iMaxExplosions = 5
      .nBGStars = 1000
      .nStellarObjectImages = 4
      .iMaxShips = 100
      .bInvertSpeakers = False
      
   End With
   
End Sub

Function GetPrefsFromFile(ByVal sFileName As String) As tGamePrefs
Dim aInput() As String
   
   Open sFileName For Input As #1
      With GetPrefsFromFile
      
         aInput = Tokenize(ReadStr(1), " ")
         .bMusicOn = CBool(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .bSoundsOn = CBool(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iBGStarsMargin = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iFrameRateAverageRange = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iMaxFPS = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iMinFPS = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iProcsPerFrame = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iMaxExplosions = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .nBGStars = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .nStellarObjectImages = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .iMaxShips = Val(aInput(1))
         aInput = Tokenize(ReadStr(1), " ")
         .bInvertSpeakers = CBool(aInput(1))
         
      End With
   Close #1

End Function

Sub LoadPrefs(ByRef aGamePrefs As tGamePrefs)

   With aGamePrefs
   
      nStellarObjectImages = .nStellarObjectImages
      mExplosions.ChangeMax .iMaxExplosions
      mStars.StarCount = .nBGStars
      mGame.ProcsPerFrame = .iProcsPerFrame
      mGame.MinFPS = .iMinFPS
      mGame.MaxFPS = .iMaxFPS
      mMonitor.FrameRateAverageRange = .iFrameRateAverageRange
      mSounds.bSFXon = .bSoundsOn
      mMusic.bMusicOn = .bMusicOn
      mShips.iMaxShips = .iMaxShips
      mSounds.InvertSpeakers = .bInvertSpeakers
      
   End With
   
End Sub
 mSounds.bSFXon = .bSoundsOn
      mMusic.bMusicOn = .bMusicOn
      mShips.iMaxShips = .iMaxShips
      mSounds.InvertSpeakers = .bInvertSpeakers
      
   End With
   
End Sub
