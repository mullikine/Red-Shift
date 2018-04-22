Attribute VB_Name = "mApplication"
'---------------------------------------------------------------------------------------
' Module    : mApplication
' DateTime  : 4/13/2005 22:58
' Author    : Shane Mulligan
' Purpose   : Contains Sub Main
'---------------------------------------------------------------------------------------

Public sh32 As New Shell32.Shell
Public AlphaMode As Boolean
Public iLoadingPosition As Integer

Sub Main()

   If App.PrevInstance Then Exit Sub
   Select Case Trim(Command)
   Case "AUTORUN"
      frmAutorun.Show
   Case Else
      LoadGame
   End Select

End Sub

Sub LoadGame()
   
   ' Init
   Randomize
   
   mFSO.Init
   mRes.Init
   mPlayerSims.LoadPlayerSims
   mWinDims.Init
   
   mDirectX.InitD3D
   mDirectX.Initialise frmDXForm.picDX.hWnd
   mSounds.Init
   mMusic.Init
   
   mTextures.Init
   DJ.NewPlay mMusic.StartTheme, 0
   
   ' fade in
   frmDXForm.Show
   For i = 0 To 255 Step 6
      SetTranslucent frmDXForm.hWnd, &H0, i, LWA_ALPHA
   Next i
   RemoveTranslucent frmDXForm.hWnd
   
   '// blacken background if required
   'mEngine.ClearAndBeginRender
   '   DrawTexture ViewImages(0), srcRECTNorm, NewfRECT(0, ScreenDims.Height, 0, ScreenDims.Width), , , &HFF000000
   'mEngine.EndRender
   mTextures.LoadTextures
   
   '// wait for music to end
   'While oDMPerf.IsPlaying(oDMSeg, SegState)
   '   DoEvents
   'Wend
   
   AlphaMode = False
   'AlphaMode = AskBox(frmDXForm, "Choose between correct blending or incorrect blending:" & vbCrLf & _
   '      "Yes - correct" & vbCrLf & _
   '      "No - incorrect", App.ProductName, vbYesNo) = vbNo
   
   frmMainMenu.Show
   frmDXForm.Hide

End Sub


Sub CloseApp()

   mGarbage_Collector.CollectGarbage
   End

End Sub


Sub IncLoadStatus()
Dim tmpLoadingDims As XYWH
Const LastLoadingPosition As Integer = 507

   DoEvents

   tmpLoadingDims = LoadingDims
   
   iLoadingPosition = iLoadingPosition + 1
   ClearAndBeginRender
      DrawTexture txrTitles(0), srcRECTNorm, NewfRECT(50, 180, 162, 862)
      EnableBlendColour
      DrawXYWH tBar1, LoadingDims, &H80303090, , True
      DrawTexture txrFlares(0), SlideEffect(tmpLoadingDims, 100), XYWHTofRECT(LoadingDims), , , Red
      tmpLoadingDims.Width = LoadingDims.Width * (iLoadingPosition / LastLoadingPosition)
      DrawTexture txrFlares(1), SlideEffect(tmpLoadingDims, 100), XYWHTofRECT(tmpLoadingDims), , , White
      'DrawXYWH tBar0, tmpLoadingDims, &H80008000, , True
      DrawXYWH ViewImages(0), LoadingDims, &H80505050
      EnableBlendNormal
   EndRender
   Debug.Print iLoadingPosition

End Sub
