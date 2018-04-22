Attribute VB_Name = "mKeyboard"
'---------------------------------------------------------------------------------------
' Module    : mKeyboard
' DateTime  : 12/16/2004 21:24
' Author    : Shane Mulligan
' Purpose   : Handles key events
'---------------------------------------------------------------------------------------

Option Explicit

Public Keys(256) As Boolean


Sub FormatKeys()

   For i = 0 To 256
      Keys(i) = False
   Next i

End Sub

Sub DoKeys()

   With Ships(You.Ship)
      
      ' Inside of Control Key scope
      '------------------------------
      
      If Keys(vbKeyControl) Then
         
         If Keys(vbKeyAdd) Then
            mGame.ProcsPerFrame = mGame.ProcsPerFrame + 1
            Keys(vbKeyAdd) = False
         End If
         
         If Keys(vbKeySubtract) Then
            mGame.ProcsPerFrame = mGame.ProcsPerFrame - 1
            If mGame.ProcsPerFrame = 0 Then mGame.ProcsPerFrame = 1
            Keys(vbKeySubtract) = False
         End If
         
         If Keys(vbKeyD) Then
            .Hull = BoundMin(.Hull - ShipTypes(.ShipType).MaxHull / 100, -1)
            .Shield = 0
         End If
         
         'Eject
         If Keys(vbKeyE) Then mDebug.sPrint "Pilot ejection system not implemented."
         
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
         
         ' Select next ship
         If Keys(vbKeyS) Then
            .CurrentShipSelection = NextShip(You.Ship, SelectRelations)
            Keys(vbKeyS) = False
         End If
         
         ' Select fleet master ship
         If Keys(vbKeyQ) Then
            .CurrentShipSelection = .OwnorShip
            Keys(vbKeyQ) = False
         End If
         
         ' Make selected ship hostile
         If Keys(vbKeyF) Then
            If Not .CurrentShipSelection = -1 Then
               ShipRelations(You.Ship, .CurrentShipSelection) = Hostile
            End If
            Keys(vbKeyF) = False
         End If
         
      Else
         
         ' Toggle autopilot
         If Keys(vbKeyC) Then
            You.Autopilot = Not You.Autopilot
            Keys(vbKeyC) = False
         End If
         
         ' Fighters hot toggle
         If Keys(vbKeyO) Then
            .FightersHot = Not .FightersHot
            Keys(vbKeyO) = False
         End If
         
         ' Switch ships
         If Keys(vbKeyPageUp) Then
            You.Ship = You.Ship + 1
            If You.Ship > nShips Then You.Ship = 0
            mStars.InitialPhysics
            'Ships(You.Ship).ObjectiveType = vbNullString
            Keys(vbKeyPageUp) = False
         End If
         
         ' toggle cloak
         If Keys(vbKeyU) Then
            .CloakOn = Not .CloakOn
            Keys(vbKeyU) = False
         End If
         
         'Select stellar object
         If Keys(vbKeyTab) Then
            .CurrentStellarObjectSelection = NextStellarObject
            Keys(vbKeyTab) = False
         End If
         
         '\\\\\\\\\\\\\\\\\\\\\
         
         ' Cycle Select Relations
         If Keys(vbKeyA) Then
            SelectRelations = SelectRelations + 1
            If SelectRelations = 5 Then SelectRelations = 0
            Keys(vbKeyA) = False
         End If
         
         ' Select closest ship
         If Keys(vbKeyS) Then
            .CurrentShipSelection = ClosestShip(You.Ship, SelectRelations)
            Keys(vbKeyS) = False
         End If
         
         '/////////////////////
         
         ' Chases selected ship
         If Keys(vbKeyDivide) And .CurrentShipSelection <> -1 Then
            .ObjectiveType = "ChSh"
            .ObjectiveIndex = .CurrentShipSelection
         End If
            
      End If
      
      ' Outside of Control Key scope
      '------------------------------
      
      ' Toggle Pause (Status menu)
      If Keys(vbKeyP) Then
         mStatusMenu.MenuUp = Not mStatusMenu.MenuUp
         Keys(vbKeyP) = False
      End If
      
      ' Toggle Map
      If Keys(vbKeyM) Then
         mMap.MapUp = Not mMap.MapUp
         Keys(vbKeyM) = False
      End If
      
      ' Toggle Radar
      If Keys(vbKeyMultiply) Then
         mRadar.RadarUp = Not mRadar.RadarUp
         mTabSelect.TabsOn = Not mTabSelect.TabsOn
         Keys(vbKeyMultiply) = False
      End If
      
      If mMap.MapUp Then
         '///MAP
         ' Zoom Map Out
         If Keys(vbKeySubtract) Then
            mMap.ScaleFactor = mMap.ScaleFactor + 1
         End If
         
         ' Zoom Map In
         If Keys(vbKeyAdd) Then
            mMap.ScaleFactor = mMap.ScaleFactor - 1
         End If
      Else
         If mRadar.RadarUp And (Keys(vbKeyEnd) Or (Not Keys(vbKeyHome) And Not Keys(vbKeyEnd))) Then
            '///RADAR
            If Keys(vbKeySubtract) Then
               RadarZoom = RadarZoom / 1.1
            End If
            If Keys(vbKeyAdd) Then
               RadarZoom = RadarZoom * 1.1
            End If
         End If
         If Keys(vbKeyHome) Or (Not Keys(vbKeyHome) And Not Keys(vbKeyEnd)) Then
            '///VIEWSCREEN
            If Keys(vbKeySubtract) Then
               SpaceZoom = SpaceZoom / 1.1
            End If
            If Keys(vbKeyAdd) Then
               SpaceZoom = SpaceZoom * 1.1
            End If
         End If
         
      End If
      
      If Keys(vbKeyF5) Then
         DJ.ChangeTempo DJ.GetTempo + 0.1
      End If
      
      If Keys(vbKeyF6) Then
         DJ.ChangeTempo DJ.GetTempo - 0.1
      End If
      
      If Keys(vbKeyF7) Then
         DJ.ChangeVolume DJ.GetVolume + 100
      End If
      
      If Keys(vbKeyF8) Then
         DJ.ChangeVolume DJ.GetVolume - 100
      End If
      
      If Keys(vbKeyF12) Then
         RadarZoom = SpaceZoom
      End If
      
      ' Quit
      If Keys(vbKeyEscape) Then
         bRunning = False
      End If
   
   End With
   
End Sub
ggleFightersHot) = False
         End If
         
         ' Switch ships
         If Keys(keyJumpShip) Then
            You.Ship = You.Ship + 1
            If You.Ship > nShips Then You.Ship = 0
            mStars.InitialPhysics
            'Ships(You.Ship).ObjectiveType = vbNullString
            Keys(keyJumpShip) = False
         End If
         
         ' toggle cloak
         If Keys(keyToggleCloak) Then
            .CloakOn = Not .CloakOn
            Keys(keyToggleCloak) = False
         End If
         
         'Select stellar object
         If Keys(keyNextPlanet) Then
            .CurrentStellarObjectSelection = NextStellarObject
            Keys(keyNextPlanet) = False
         End If
         
         '\\\\\\\\\\\\\\\\\\\\\
         
         ' Cycle Select Relations
         If Keys(keyCycleSelRelations) Then
            SelectRelations = SelectRelations + 1
            If SelectRelations = 5 Then SelectRelations = 0
            Keys(keyCycleSelRelations) = False
         End If
         
         ' Select closest ship
         If Keys(keyClosestShip) Then
            .CurrentShipSelection = ClosestShip(You.Ship, SelectRelations)
            Keys(keyClosestShip) = False
         End If
         
         '/////////////////////
         
         ' Chases selected ship
         If Keys(keyChaseShip) And .CurrentShipSelection <> -1 Then
            .ObjectiveType = "ChSh"
            .ObjectiveIndex = .CurrentShipSelection
            Keys(keyChaseShip) = False
         End If
            
      End If
   
   End With
   
End Sub
