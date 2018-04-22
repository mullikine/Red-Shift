VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8                    'creates the "DirectXEvent8_DXCallback" function.

'---------------------------------------------------------------------------------------
' Module    : cDJ
' DateTime  : 3/25/2005 10:49
' Author    : Shane Mulligan
' Purpose   : Implements music using DirectX
'---------------------------------------------------------------------------------------

'DirectMusic Objects:
'Private oDX As DirectX8                     'the root of all DX8 applications
Private oDMPerf As DirectMusicPerformance8  'the master performance
Private oDMLoader As DirectMusicLoader8     'helps load music into buffers
Private oDMSeg As DirectMusicSegment8       'actually stores the music to be played

'Internal Variables
Private hEvent As Long  'An Event handle
Private bInit As Boolean 'did we initialise successfully
Private lMusicTimeAtStart As Long   'what the music time offset was when the audio piece started...

Public SegState As DirectMusicSegmentState8

Sub NewPlay(ByVal lpszSoundName As String, ByVal Repeat As Long)

   If InitDMusic(lpszSoundName) Then
      bInit = True
      PlayMusic Repeat
   End If

End Sub

Function InitDMusic(ByVal lpszSoundName As String) As Boolean
Dim dmParams As DMUS_AUDIOPARAMS

   On Error GoTo BailOut:
      
      If Not bMusicOn Then GoTo BailOut
      
      Set oDMPerf = Dx.DirectMusicPerformanceCreate
      Set oDMLoader = Dx.DirectMusicLoaderCreate
      
      'You can play around with these settings to see what happens...
      oDMPerf.InitAudio frmDXForm.hWnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
      oDMPerf.SetMasterAutoDownload True
      
   '//2. Setup the messaging system
      'this next part will allow DM8 to talk back to us; with most DX programs
      'we talk to it (function calls etc...), but we can also let it call our
      'functions (the DXCALLBACK function). For those that know about this sort
      'of thing, it's just a fairly simple Callback structure.
      'oDMPerf.AddNotificationType DMUS_NOTIFY_ON_SEGMENT 'relay messages about the segment
      'hEvent = Dx.CreateEvent(frmDXForm)
      'oDMPerf.SetNotificationHandle hEvent 'used to identify the messages
      
      Debug.Print "DirectMusic8 Initialised."

   '//3. Configure the loader
       oDMLoader.SetSearchDirectory App.Path & "\Data\Music\"
         'the path above specifies where we are loading
         'the audio file from. If you load it from a different
         'folder, "\Music\Techno\" for example, then you need to
         'change the above parameter.
         
   '//4. Load in the audio file
      Set oDMSeg = oDMLoader.LoadSegment(App.Path & "\Data\Music\" & lpszSoundName)
         'the loader will correctly parse the following file formats:
         '.WAV; .MID; .SEG; .RMI
         'IT WONT LOAD MP3 files, if you require MP3 audio then look
         'at the DirectShow8 tutorials.
          
      oDMSeg.SetStandardMidiFile
         'for .MID / .MIDI / .RMI files we must make this additional call.
         'for other formats (SEG/WAV) dont make this call.
         
   '//5. setup tempo and volume
      ChangeTempo 1
      ChangeVolume 0
      
      Debug.Print "Audio Segment loaded..."
   InitDMusic = True
   Exit Function
BailOut:
   InitDMusic = False
End Function

Private Function TerminateDMusic() As Boolean
On Error GoTo BailOut:

   oDMPerf.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
   Dx.DestroyEvent hEvent
   
   If ObjPtr(oDMSeg) Then Set oDMSeg = Nothing
   If ObjPtr(oDMLoader) Then Set oDMLoader = Nothing
   If Not (oDMPerf Is Nothing) Then
      oDMPerf.CloseDown
      Set oDMPerf = Nothing
   End If
   'If ObjPtr(DX) Then Set DX = Nothing
   
   
   TerminateDMusic = True
   Exit Function
BailOut:
   TerminateDMusic = False
End Function

Sub PlayMusic(ByVal Repeat As Long)

   If Not bInit Then Exit Sub
   
   Debug.Print "PLAY"
   
   oDMSeg.SetRepeats Repeat ' -1 = play infinitely = loop
   
   Set SegState = oDMPerf.PlaySegmentEx(oDMSeg, DMUS_SEGF_DEFAULT, 0)
   lMusicTimeAtStart = oDMPerf.GetMusicTime

End Sub

Sub StopMusic()

    If Not bInit Then Exit Sub
    
    Debug.Print "STOP"
    
    oDMPerf.StopEx oDMSeg, 0, DMUS_SEGF_DEFAULT
    
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
If eventid = hEvent Then
    'the message is for us
    Dim dmMSG As DMUS_NOTIFICATION_PMSG
    
    If Not oDMPerf.GetNotificationPMSG(dmMSG) Then
        'error!
        Debug.Print "MESSAGE RECIEVED: Error in Parsing"
    Else
        Select Case dmMSG.lNotificationOption
            Case DMUS_NOTIFICATION_SEGABORT
                Debug.Print "MESSAGE: DMUS_NOTIFICATION_SEGABORT"
            Case DMUS_NOTIFICATION_SEGALMOSTEND
                Debug.Print "MESSAGE: DMUS_NOTIFICATION_SEGALMOSTEND"
            Case DMUS_NOTIFICATION_SEGEND
                Debug.Print "MESSAGE: DMUS_NOTIFICATION_SEGEND"
            Case DMUS_NOTIFICATION_SEGLOOP
                Debug.Print "MESSAGE: DMUS_NOTIFICATION_SEGLOOP"
                lMusicTimeAtStart = oDMPerf.GetMusicTime
            Case DMUS_NOTIFICATION_SEGSTART
                Debug.Print "MESSAGE: DMUS_NOTIFICATION_SEGSTART"
                lMusicTimeAtStart = oDMPerf.GetMusicTime
            Case Else
                Debug.Print "MESSAGE: Unknown"
        End Select
    End If
End If
End Sub

Sub ChangeTempo(ByVal NewTempo As Single)
Const Min As Single = 0.01
Const Max As Single = 4

   If Not bInit Then Exit Sub
   
   If NewTempo < Min Then NewTempo = Min
   If NewTempo > Max Then NewTempo = Max
      
   oDMPerf.SetMasterTempo NewTempo
   
   Debug.Print "Master Tempo Changed: " & NewTempo
   
End Sub

Function GetTempo() As Single

   GetTempo = oDMPerf.GetMasterTempo

End Function

Sub ChangeVolume(ByVal NewVolume As Integer)
Const Min As Integer = -2500
Const Max As Integer = 200

   If Not bInit Then Exit Sub
   
   If NewVolume < Min Then NewVolume = Min
   If NewVolume > Max Then NewVolume = Max
        
   oDMPerf.SetMasterVolume NewVolume
    
   Debug.Print "Master Volume Changed: " & NewVolume
    
End Sub

Function GetVolume() As Integer

   GetVolume = CInt(oDMPerf.GetMasterVolume)

End Function

Function MusicProgress() As String

    MusicProgress = "Time: " & oDMPerf.GetMusicTime - lMusicTimeAtStart & " / " & oDMSeg.GetLength

End Function

Private Sub Class_Terminate()

   If Not TerminateDMusic() Then
      Debug.Print "Unable to terminate DirectMusic8 Successfully..."
   End If

End Sub

Function Playing() As Boolean

   If Not bInit Then Exit Function
   
   Playing = oDMPerf.IsPlaying(oDMSeg, SegState)

End Function
