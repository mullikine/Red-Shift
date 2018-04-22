Attribute VB_Name = "mSounds"
'---------------------------------------------------------------------------------------
' Module    : mSounds
' DateTime  : 12/16/2004 21:19
' Author    : Shane Mulligan
' Purpose   : Implements SFX
'---------------------------------------------------------------------------------------

Option Explicit

Public Const MaxFrequency As Long = 90000
Public Const MinFrequency As Long = 0
Public Const MaxVolume As Integer = 0
Public Const MinVolume As Integer = -9000
Public Const MaxPan As Integer = 9000
Public Const MinPan As Integer = -9000

Public SoundNameStrings() As String
Public bSFXon As Boolean

Private DS As DirectSound8
Private DSToneBuffers() As DirectSoundSecondaryBuffer8

Private desc As DSBUFFERDESC

Public nSounds As Integer
Public nBuffersPerSound As Integer
Public MaxRange As Single

Public InvertSpeakers As Boolean

Sub Init()
Dim aInput() As String

   ' Get the dimensions
   Open App.Path & "\Data\Sounds\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
      nSounds = Val(aInput(1))
      aInput = Tokenize(ReadStr(1), " ")
      nBuffersPerSound = Val(aInput(1))
       aInput = Tokenize(ReadStr(1), " ")
      MaxRange = Val(aInput(1))
   Close #1
   
   ' Set the buffers
   Set DS = Dx.DirectSoundCreate("")
   
   DS.SetCooperativeLevel frmDXForm.picDX.hWnd, DSSCL_NORMAL
   desc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
   
   ReDim DSToneBuffers(nSounds, nBuffersPerSound)
   For i = 0 To nSounds
      For j = 0 To nBuffersPerSound
         Set DSToneBuffers(i, j) = DS.CreateSoundBufferFromFile(App.Path & "\Data\Sounds\" & i & ".wav", desc)
      Next j
   Next i

End Sub

Sub Play(ByVal eSoundName As SoundName, Optional ByVal Volume As Long = -1000, Optional ByVal Pan As Integer = 0, Optional ByVal Frequency As Long = 0, Optional ByVal bLoop As Boolean)

    If bSFXon Then
      
      If InvertSpeakers Then Pan = -Pan
      
      For i = 0 To 25
         If Not DSToneBuffers(eSoundName, i).GetStatus = DSBSTATUS_PLAYING Then
            With DSToneBuffers(eSoundName, i)
            
               If Volume < MinVolume Then Volume = MinVolume
               If Volume > MaxVolume Then Volume = MaxVolume
               .SetVolume Volume
               If Frequency < MinFrequency Then Frequency = MinFrequency
               If Frequency > MaxFrequency Then Frequency = MaxFrequency
               .SetFrequency Frequency
               If Pan < MinPan Then Pan = MinPan
               If Pan > MaxPan Then Pan = MaxPan
               .SetPan Pan
               If bLoop Then
                  .Play DSBPLAY_LOOPING
               Else
                  .Play DSBPLAY_DEFAULT
               End If
               
            End With
            Exit For
         End If
      Next
   End If

End Sub

Sub CleanUp()

   Erase DSToneBuffers

End Sub

   End If
               
            End With
            Exit For
         End If
      Next
   End If

End Sub

Sub CleanUp()

   Erase DSToneBuffers

End Sub

