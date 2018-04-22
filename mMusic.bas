Attribute VB_Name = "mMusic"
Option Explicit

Public StartTheme As String
Public MenuMusic As String
Public Playlist() As String

Public bMusicOn As Boolean

Public DJ As cDJ

Sub Init()
Dim aInput() As String
   
   Open App.Path & "\Data\Music\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
      StartTheme = Replace(aInput(1), "~", " ")
      aInput = Tokenize(ReadStr(1), " ")
      MenuMusic = Replace(aInput(1), "~", " ")
   Close #1

   Open App.Path & "\Data\Music\Playlist\Specs.txt" For Input As #1
      Playlist = Split(Input(LOF(1), #1), vbCrLf)
   Close #1
   
   Set DJ = New cDJ
   DJ.ChangeVolume -1000

End Sub

Sub DoMusic()

   If (Not DJ.Playing) And bMusicOn Then
      mStatusMenu.MenuUp = True
      DJ.NewPlay "Playlist\" & Playlist(Int(Rnd * (UBound(Playlist) + 1))), 0
      While DJ.Playing = False
         DoEvents
      Wend
      mStatusMenu.MenuUp = False
   End If

End Sub


End Sub
