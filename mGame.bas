Attribute VB_Name = "mGame"
'---------------------------------------------------------------------------------------
' Module    : mGame
' DateTime  : 12/6/2004 21:27
' Author    : Shane Mulligan
' Purpose   : Controls the game flow
'---------------------------------------------------------------------------------------

Option Explicit

Public i As Integer, j As Integer, k As Integer 'General iterators
Public s As Single                              'General single

Public ProcsPerFrame As Integer
Public MinFPS As Integer
Public MaxFPS As Integer

Public You As tPlayer

Public bRunning As Boolean, bActive As Boolean

' The dying fade out
Const InitDying As Integer = 32
Public Dying As Integer

Sub RunGame()

   Dying = InitDying
   
   '  Prepare routine
   
   frmDXForm.Show
   frmGameMenu.Show
   
   frmMainMenu.Hide
   frmSolo.Hide
      'Unload frmMainMenu
      'Unload frmSolo
   Initialize_Game
   
   frmDXForm.Show
   frmGameMenu.Hide
   
   mSounds.Play sndAfterburner
   
   bRunning = True
   
   
   '  Play game
   While bRunning And DoEvents()
      If Not mStatusMenu.MenuUp Then
         mMonitor.FPS
         mGame.DoPhysics
         mMusic.DoMusic
      End If
      mMonitor.ProcCount = mMonitor.ProcCount + 1
      mGame.Draw
      mKeyboard.DoKeys
      ' if you die, game ends
      If Ships(You.Ship).Died Then
         Dying = Dying - 1
      Else
         Dying = InitDying
      End If
      bRunning = bRunning And Not Dying < 1
   Wend
   EnableBlendOne
   
   ' Stop some processes
   DJ.StopMusic
   
   ' Clear the device
   'Call ClearAndBeginRender: Call EndRender
   
   'mEngine.Initialise frmDXForm.picdx.hWnd  ' Make anew
   
   SavePlayerSims
   
   frmMainMenu.Show
   frmSolo.Show
   'frmDXForm.Hide
   
   mKeyboard.FormatKeys

End Sub


Sub Draw()
   
   If mMonitor.ProcCount Mod ProcsPerFrame <> 0 Then Exit Sub
   
   If ProcCount > 1 Then ClearAndBeginRender
   
   ' ************************Start drawing
   
   mViewer.DoDraw
   mStatusBar.Draw
   mHUD.DoDraw
   mMap.DoDraw
   mStatusMenu.DoDraw
   
   mDebug.Draw
   
   If Ships(You.Ship).Died Then
      EnableBlendOne
      ' dying fade out
      DrawTexture ViewImages(0), srcRECTNorm, NewfRECT(0, ScreenDims.Height, 0, ScreenDims.Width), , , &HFF000000 + ((255 - (Dying * 8 - 1)) / 255) * &HFF0000
      EnableBlendNormal
   End If
   
   'mCursor.Draw
   ' ***********************Finish drawing
   
   If ProcCount > 1 Then EndRender

End Sub


Sub DoPhysics()

    mCheats.Process
    
    '> Think methods
    mShips.DoPhysics
    mProjectiles.DoPhysics
    mStellarObjects.DoPhysics
    
    ' Must come after mShips.DoPhysics
    ' due to star reset @ fast speeds
    mStars.DoPhysics
    mExplosions.DoPhysics

End Sub


Private Sub Initialize_Game()

   With frmGameMenu
      
      .PrintText "Formatting keys..."
      mKeyboard.FormatKeys
      
      .PrintText "Initializing viewer..."
      mViewer.Init   'Must come before the initialization of visual elements esp. Stars
      
      .PrintText "Initializing guns..."
      mGuns.Init
      
      .PrintText "Initializing ships..."
      mShips.Init
      
      .PrintText "Initializing projectiles..."
      mProjectiles.Init
      
      .PrintText "Initializing stars..."
      mStars.Init
      
      .PrintText "Playing Music..."
      DJ.NewPlay Playlist(Int(Rnd * (UBound(Playlist) + 1))), 0
      
      .PrintText "Initializing radar functions..."
      mRadar.Init
      
      .PrintText "Initializing scanner functions..."
      mScanner.Init
      
      .PrintText "Initializing status menu..."
      mStatusMenu.Init
      
      .PrintText "Initializing tab select..."
      mTabSelect.Init
      
      .PrintText "Finalizing..."
      mStars.InitialPhysics

   End With

End Sub
StatusMenu.Init
   Call SmallIncLoadStatus: mTabSelect.Init
   Call SmallIncLoadStatus: mKeyboard.InitKeyConfig
   Call SmallIncLoadStatus: mStars.InitialPhysics

End Sub
