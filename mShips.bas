Attribute VB_Name = "mShips"
'---------------------------------------------------------------------------------------
' Module    : mShips
' DateTime  : 4/12/2005 12:21
' Author    : Shane Mulligan
' Purpose   : Ship physics and drawing
'---------------------------------------------------------------------------------------

Option Explicit

Public Ships() As tShip
Public nShips As Integer
Public iMaxShips As Integer

Public RefShip As Integer

Sub Init()
   
   nShips = -1
   
   You.Ship = NewShip(PlayerSim.ShipType, -1, PlayerSim.System, PlayerSim.x, PlayerSim.y, 0, 0, 0, 0, Trader, , PlayerSim.Credit)
   
End Sub

Function NewShip(ByVal ShipType As Integer, ByVal OwnorShip As Integer, ByVal System As Integer, ByVal x As Long, ByVal y As Long, ByVal maMod As Single, ByVal maArg As Single, ByVal maBearing As Single, ByVal pGovernment As Integer, ByVal pCareer As eCareer, Optional ByVal pRefShip As Integer = -1, Optional ByVal pCredit As Integer = 0, Optional ByVal pHyperspaceTo As Integer = -1) As Integer

   If Not nShips < iMaxShips Then NewShip = -1: Exit Function
   
   If pRefShip = -1 Then
      For i = 0 To nShips
         If Ships(i).Died Then
            pRefShip = i
            GoTo WithShip
         End If
      Next i
      nShips = nShips + 1
      ReDim Preserve Ships(nShips) As tShip
      pRefShip = nShips
   End If
   
WithShip:
   With Ships(pRefShip)
   
      .Died = False
      .Kills = 0
      .Credit = 0
      .ShipType = ShipType
      .OwnorShip = OwnorShip
      .System = System
      .x = x
      .y = y
      .maMod = maMod
      .maArg = maArg
      .maBearing = maBearing
      .HyperspaceDestination = -1
      .Name = "Computer"
      .ObjectiveType = vbNullString
      .Credit = pCredit
      .Career = pCareer
      .Government = pGovernment
      .FightersHot = False
      .CurrentShipSelection = -1
      .InHyperspace = False
      HyperspaceTo pHyperspaceTo, pRefShip
      
      .Shield = ShipTypes(.ShipType).MaxShield
      .Hull = ShipTypes(.ShipType).MaxHull
      .FuelLeft = ShipTypes(.ShipType).MaxFuel
      .Cloak = ShipTypes(.ShipType).MaxCloak
      .Battery = ShipTypes(.ShipType).MaxBattery
      
      SelectDefaultGuns pRefShip

   End With
   
   If nShips = 0 Then
         FormatShipRelations
      Else
         FormatIndividualShipRelations pRefShip
   End If
   
   ' Return pRefship
   NewShip = pRefShip

End Function

Sub Draw()

   DrawBodies

End Sub

Private Sub DrawBodies()

   For RefShip = 0 To nShips
      With Ships(RefShip)
         
         ' Makes sure the ship is within an appropriate proximity
         If .System = Ships(You.Ship).System Then
            If ZoomMod(mTabSelect.DistanceFromShip(You.Ship, RefShip)) < 1024 And Not .Died Then
               DrawJetstream
               DrawBody
            End If
         End If
         
      End With
   Next RefShip

End Sub


Private Sub DrawBody()
Dim CloakColour As Long
Dim srcRECT As fRECT
Dim destRECT As fRECT

   With Ships(RefShip)
      
      Select Case True
      Case .InHyperspace
         CloakColour = Blue
         EnableBlendOne
      Case .CloakOn
         CloakColour = Green - &H80000000
         EnableBlendOne
      Case Else
         CloakColour = &HFF - ((ShipTypes(.ShipType).MaxHull - .Hull) / ShipTypes(.ShipType).MaxHull) * &HFF
         If CloakColour < 128 Then CloakColour = 128
         CloakColour = &HFF000000 + RGB(CloakColour, CloakColour, CloakColour)
         EnableBlendNormal
      End Select
      
      srcRECT.Top = 1
      srcRECT.Bottom = 0
      If .maBearing > 180 And ShipImageSets(ShipTypes(.ShipType).ShipImage).FlipX Then
         srcRECT.Right = 0
         srcRECT.Left = 1
      Else
         srcRECT.Right = 1
         srcRECT.Left = 0
      End If
      
      destRECT = NewfRECT(.y - ShipTypes(.ShipType).Size / 2, .y + ShipTypes(.ShipType).Size / 2, .x - ShipTypes(.ShipType).Size / 2, .x + ShipTypes(.ShipType).Size / 2)
      
      DrawTexture Image(.ShipType, .maBearing), srcRECT, destRECT, .maBearing - Round(.maBearing / ShipImageSets(ShipTypes(.ShipType).ShipImage).DeltaDegs, 0) * ShipImageSets(ShipTypes(.ShipType).ShipImage).DeltaDegs, True, CloakColour
   
   End With

End Sub


Sub DrawRadars()

   For RefShip = 0 To nShips
      With Ships(RefShip)
         
         If Not .Died And Not .CloakOn And .System = Ships(You.Ship).System Then
            mShips.DrawRadar
         End If
         
      End With
   Next RefShip

End Sub


Private Sub DrawRadar()

   With Ships(RefShip)
      
      mRadar.DrawToRadar mTextures.txrFlares(1), .x, .y, ShipTypes(.ShipType).Size, RelationColour(ShipRelations(You.Ship, RefShip))
      
   End With

End Sub


Sub DoPhysics()
Dim ReincarnateShip As Integer

   For RefShip = 0 To nShips
      With Ships(RefShip)
      
         If .Died Then
            If RefShip = You.Ship Then
               ' Switch ships to a fleetmember if possible
               ReincarnateShip = ClosestShipByFuel(You.Ship, Member)
               If ReincarnateShip <> -1 Then
                  You.Ship = ReincarnateShip
                  DisplayMessage "You are now commanding a ship member of your fleet.", Green
               End If
            End If
         Else
            If DoDoom Then
               If RefShip = You.Ship Then
                  DisplayMessage "Your ship has been destroyed.", Red
               End If
            Else
               DoShipRefs
               If .HyperspaceDestination <> -1 Then
                  DoHyperspace
               Else
                  If RefShip = You.Ship Then
                     mStars.BackColour = &H0
                  End If
                  If RefShip = You.Ship And Not You.Autopilot Then
                     DoKeys
                  Else
                     DoThink
                     DoObjectives
                  End If
               End If
               DoHUD
               DoDamage
               DoFuelShieldCloakBattery
               DoFriction
               DoKeyOrThinkAction
               DoGravity
               DoSpinAndBearing
               DoHeadingAndBearing
               DoNewLocation
               DoGuns
            End If
         End If
         
      End With
   Next RefShip
   NewShips

End Sub

Sub DoHUD()
Dim iSO As Integer

   With Ships(RefShip)
   
      .CurrentStellarObject = -1
      For iSO = 0 To UBound(StellarObjects)
         If StellarObjects(iSO).System = .System Then
            If DistanceFromStellarObject(RefShip, iSO) <= 0 Then .CurrentStellarObject = iSO
         End If
      Next iSO
   
   End With

End Sub

Private Sub DoFriction()

   With Ships(RefShip)
      
      .maMod = .maMod * ShipTypes(.ShipType).FrictionRatio ^ .maMod
      .Spin = .Spin * ShipTypes(.ShipType).SpinFrictionRatio ^ Abs(.Spin)
      
   End With

End Sub


Private Sub DoGuns()
Dim RefGun As Integer
Dim TempPoint As Point

   With Ships(RefShip)
   
      For RefGun = 0 To UBound(Split(.Guns, ","))
         If Split(.Guns, ",")(RefGun) <> -1 Then
            If Guns(Split(.Guns, ",")(RefGun)).GunType <> -1 Then
               ' Gun location
               TempPoint = RotatePoint(Split(ShipTypes(.ShipType).GunPositionsX, ",")(RefGun), Split(ShipTypes(.ShipType).GunPositionsY, ",")(RefGun), .maBearing)
               
               ' Aim gun
               If .CurrentShipSelection <> -1 And GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).Ballistic Then
                  If .CurrentShipSelection = RefShip Then
                     Guns(Split(.Guns, ",")(RefGun)).Bearing = .maBearing
                  Else
                     Guns(Split(.Guns, ",")(RefGun)).Bearing = CartToArg(Ships(.CurrentShipSelection).x - (.x + TempPoint.x), Ships(.CurrentShipSelection).y - (.y + TempPoint.y))
                  End If
               Else
                  Guns(Split(.Guns, ",")(RefGun)).Bearing = .maBearing
               End If
               
               ' Fire
               If Guns(Split(.Guns, ",")(RefGun)).Firing Then
                  Guns(Split(.Guns, ",")(RefGun)).Firing = False
                  ' Valid gun, ammo left
                  If Not (GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).FireRate = 0 Or Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining = 0) Then
                     ' Fire rate
                     If ProcCount - Guns(Split(.Guns, ",")(RefGun)).ProcLastFired >= GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).FireRate Then
                        Guns(Split(.Guns, ",")(RefGun)).ProcLastFired = ProcCount
                        mProjectiles.MakeProjectile RefShip, RefGun, .System, .x + TempPoint.x, .y + TempPoint.y, .CurrentShipSelection, Guns(Split(.Guns, ",")(RefGun)).Bearing
                        Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining = Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining - 1
                     End If
                  End If
               End If
               
               ' Energy weapon recharge
               If ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).ProjectileType).WeaponClass = Energy Or ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).ProjectileType).WeaponClass = Beam Then
                  If mMonitor.ProcCount Mod GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).RechargeRate = 0 Then
                     If Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining < GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).MaxAmmo Then
                        If .Battery > GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).RechargePowerCost Then
                           Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining = Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining + 1
                           .Battery = .Battery - GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).RechargePowerCost
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next RefGun
      
   End With

End Sub


Private Sub DoSpinAndBearing()

   With Ships(RefShip)
      
      .maBearing = .maBearing + .Spin
      
      DoHeadingAndBearing
      
   End With

End Sub


Private Sub DoObjectives()
Dim ModArg As Pol2

   With Ships(RefShip)
      
      Select Case .ObjectiveType
      Case "SlfD" ' self destruct
            .Hull = BoundMin(.Hull - 1, -1)
            .Shield = 0
         
      Case "RFSO" ' refuel at stellar object
         If .System = StellarObjects(.ObjectiveIndex).System Then
            ModArg.A = CartToArg(StellarObjects(.ObjectiveIndex).x - .x, StellarObjects(.ObjectiveIndex).y - .y)
            ModArg.M = DistanceFromStellarObject(RefShip, .ObjectiveIndex)
            ModArg = Pol2Pol2Add(ModArg, NewPol2(-.maMod * 2, .maArg))
            
            TurnToBearing ModArg.A
            
            .AfterburnerOn = False ' dont waste fuel
            
            If ModArg.M > StellarObjects(.ObjectiveIndex).Size Then
               .MoveUp = True
               .MoveDown = False
            Else
               .MoveUp = False
               .MoveDown = True
            End If
            
            If .FuelLeft = ShipTypes(.ShipType).MaxFuel Then .ObjectiveType = vbNullString
         Else
            ' try to hyperspace. if not then...
            Select Case HyperspaceTo(StellarObjects(.ObjectiveIndex).System)
            Case Not_Enough_Fuel
               .ObjectiveType = "RFSO-" ' ...no refuel
            Case Not_Far_Enough_From_Center
               TurnToBearing CartToArg(.x, .y)
               .MoveUp = True
            End Select
         End If
         
      Case "FlSh" ' Follow ship
         If .System = Ships(.ObjectiveIndex).System Then
            ModArg.A = CartToArg(Ships(.ObjectiveIndex).x - .x, Ships(.ObjectiveIndex).y - .y)
            ModArg.M = DistanceFromShip(RefShip, .ObjectiveIndex)
            ModArg = Pol2Pol2Add(ModArg, NewPol2(-.maMod * 2, .maArg))
            
            TurnToBearing ModArg.A
            
            .MoveUp = True
            
            ' Slow down when getting close
            If .maMod <> 0 Then .MoveDown = ModArg.M / .maMod < 5
            
            ' Slow down when target ship has stopped
            If Ships(.ObjectiveIndex).maMod < 5 Then .MoveDown = True
            
            If ModArg.M < 600 Then
               If ModArg.M < 100 Then
                  .MoveDown = True
                  .MoveUp = False
                  TurnToBearing Ships(.ObjectiveIndex).maBearing
               End If
            Else
               If .maMod <> 0 Then .AfterburnerOn = ModArg.M / .maMod > 70 And Abs(DifferenceBetweenAngles(DirectionToShip(RefShip, .ObjectiveIndex), .maArg)) < 10
               .MoveDown = False
            End If
            
            ' don't stop unless forced or other priority
            If Ships(.ObjectiveIndex).CloakOn Then .ObjectiveType = vbNullString
         Else
            ' try to hyperspace. if not then...
            Select Case HyperspaceTo(StellarObjects(.ObjectiveIndex).System)
            Case Not_Enough_Fuel
               .ObjectiveType = "RFSO"
               .ObjectiveIndex = ClosestStellarObjectByFuel(RefShip, .Government)
            Case Not_Far_Enough_From_Center
               TurnToBearing CartToArg(.x, .y)
               .MoveUp = True
            End Select
         End If
         
      Case "ChSh" ' Check out ship
         If .System = Ships(.ObjectiveIndex).System Then
            ModArg.A = CartToArg(Ships(.ObjectiveIndex).x - .x, Ships(.ObjectiveIndex).y - .y)
            ModArg.M = DistanceFromShip(RefShip, .ObjectiveIndex)
            ModArg = Pol2Pol2Add(ModArg, NewPol2(.maMod, .maArg))
            
            If TurnToBearing(ModArg.A) Then
               .MoveUp = True
               .AfterburnerOn = True
               .MoveDown = False
            Else
               .MoveDown = True
            End If
            
            If ModArg.M < 400 Or Ships(.ObjectiveIndex).CloakOn Then .ObjectiveType = vbNullString
         Else
            ' try to hyperspace. if not then...
            Select Case HyperspaceTo(StellarObjects(.ObjectiveIndex).System)
            Case Not_Enough_Fuel
               .ObjectiveType = "RFSO"
               .ObjectiveIndex = ClosestStellarObjectByFuel(RefShip, .Government)
            Case Not_Far_Enough_From_Center
               TurnToBearing CartToArg(.x, .y)
               .MoveUp = True
            End Select
         End If
      
      Case "DsSh" ' Destroy ship
         If .System = Ships(.ObjectiveIndex).System Then
            If Not Ships(.ObjectiveIndex).CloakOn Then
               .CurrentShipSelection = .ObjectiveIndex
               
               ModArg.A = CartToArg(Ships(.ObjectiveIndex).x - .x, Ships(.ObjectiveIndex).y - .y)
               ModArg.M = DistanceFromShip(RefShip, .ObjectiveIndex)
               ModArg = Pol2Pol2Add(ModArg, Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(Ships(.ObjectiveIndex).maMod, Ships(.ObjectiveIndex).maArg)))
               
               TurnToBearing ModArg.A
               
               .MoveUp = True
               
               ' Slow down when getting close
               If .maMod <> 0 Then .MoveDown = ModArg.M / .maMod < 5
               
               For i = 0 To UBound(Split(.Guns, ","))
                  If Split(.Guns, ",")(i) <> -1 Then
                     If Guns(Split(.Guns, ",")(i)).GunType <> -1 Then
                        If ModArg.M <= ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ProjectileType).EffectiveRange Then
                           If ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ProjectileType).Homing Then
                              Guns(Split(.Guns, ",")(i)).Firing = True
                           Else
                              If GunTypes(Guns(Split(.Guns, ",")(i)).GunType).Ballistic Then
                                 Guns(Split(.Guns, ",")(i)).Firing = True
                              Else
                                 If TurnToBearing(ModArg.A) Then
                                    Guns(Split(.Guns, ",")(i)).Firing = True
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               Next i
               
               ' If in range
               If ModArg.M < 400 Then
                  If ModArg.M > 0 And ModArg.M < 180 Then
                     .MoveDown = True
                     If ModArg.M < 150 Then .MoveUp = False
                  End If
               Else
                  If .maMod <> 0 Then
                     If ModArg.M / .maMod > 70 And Abs(Mod360(DirectionToShip(RefShip, .ObjectiveIndex) - .maArg) - 180) < 10 Then
                        .AfterburnerOn = True
                     End If
                  End If
                  .MoveDown = False
               End If
            End If
            
            If Ships(.ObjectiveIndex).Died Then .ObjectiveType = vbNullString
         Else
            ' try to hyperspace. if not then...
            Select Case HyperspaceTo(StellarObjects(.ObjectiveIndex).System)
            Case Not_Enough_Fuel
               .ObjectiveType = "RFSO"
               .ObjectiveIndex = ClosestStellarObjectByFuel(RefShip, .Government)
            Case Not_Far_Enough_From_Center
               TurnToBearing CartToArg(.x, .y)
               .MoveUp = True
            End Select
         End If
         
      Case "ChSO" ' Check out stellar object
         If .System = StellarObjects(.ObjectiveIndex).System Then
            ModArg.A = CartToArg(StellarObjects(.ObjectiveIndex).x - .x, StellarObjects(.ObjectiveIndex).y - .y)
            ModArg.M = DistanceFromStellarObject(RefShip, .ObjectiveIndex)
            
            If TurnToBearing(ModArg.A) Then
               .MoveUp = True
               .AfterburnerOn = True
               .MoveDown = False
            Else
               .MoveDown = True
            End If
            
            If .CurrentStellarObject = .ObjectiveIndex Then .ObjectiveType = vbNullString
         Else
            ' try to hyperspace. if not then...
            Select Case HyperspaceTo(StellarObjects(.ObjectiveIndex).System)
            Case Not_Enough_Fuel
               .ObjectiveType = "RFSO"
               .ObjectiveIndex = ClosestStellarObjectByFuel(RefShip, .Government)
            Case Not_Far_Enough_From_Center
               TurnToBearing CartToArg(.x, .y)
               .MoveUp = True
            End Select
         End If
         
      End Select
      
   End With

End Sub


Private Sub DoKeys()

   With Ships(RefShip)
      
      .AfterburnerOn = Keys(vbKeyZ)
      .MoveUp = Keys(vbKeyUp)
      .MoveDown = Keys(vbKeyDown)
      .MoveLeft = Keys(vbKeyLeft)
      .MoveRight = Keys(vbKeyRight)
      .Align = Keys(vbKeyEnd)
      
      For i = 0 To UBound(Split(.Guns, ","))
         If Split(.Guns, ",")(i) <> -1 Then
            If Guns(Split(.Guns, ",")(i)).GunType <> -1 Then
               Guns(Split(.Guns, ",")(i)).Firing = Keys(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).KeyTrigger) Or Keys(49 + i)
            End If
         End If
      Next i
      
   End With

End Sub


Private Sub DoThink()
Dim iObjective As Integer

   With Ships(RefShip)
      
      ' When out of fuel and stuffed
      If .FuelLeft = 0 And .maMod < 2 Then
         .ObjectiveType = "SlfD" ' self destruct
      Else
         ' Emergency refuel
         If .FuelLeft <= FuelToStellarObject(RefShip, ClosestStellarObjectByFuel(RefShip, .Government)) And Not .ObjectiveType = "RFSO-" Then
            .ObjectiveType = "RFSO"
            .ObjectiveIndex = ClosestStellarObjectByFuel(RefShip, .Government)
         Else
            ' If there is a leader
            If .OwnorShip <> -1 Then
               ' Be hostile against anyone who the leader is hostile to
               For i = 0 To nShips
                  If ShipRelations(.OwnorShip, i) = Hostile Then
                     ShipRelations(RefShip, i) = Hostile
                  ElseIf ShipRelations(RefShip, i) = Hostile Then
                     ShipRelations(.OwnorShip, i) = Hostile
                  End If
               Next i
               ' Attack only with leader's permission
               If Ships(.OwnorShip).FightersHot = False Then GoTo SkipAttack
            End If
            
            ' Attack closest hostiles
            If ClosestShip(RefShip, Hostile) <> -1 Then
               .ObjectiveType = "DsSh"
               .ObjectiveIndex = ClosestShip(RefShip, Hostile)
               .FightersHot = True
               Exit Sub
            Else
               .FightersHot = False
            End If
         
SkipAttack:
            ' Follow ownor ship
            If .OwnorShip <> -1 Then
               .ObjectiveType = "FlSh"
               .ObjectiveIndex = .OwnorShip
            Else
TryAgain:
               ' Random objective
               iObjective = Int(Rnd * 5)
               Select Case iObjective
               Case 0, 1, 2
                  .ObjectiveType = "ChSO"
                  .ObjectiveIndex = RandomSO(.System)
                  If .ObjectiveIndex = -1 Then GoTo TryAgain
               Case 3
                  .ObjectiveType = "ChSO"
                  .ObjectiveIndex = RandomSO(RandomSys(Systems(.System).Galaxy))
                  If .ObjectiveIndex = -1 Then GoTo TryAgain
               Case 4
                  .ObjectiveType = "ChSh"
                  .ObjectiveIndex = ClosestShip(RefShip, Neutral)
                  If .ObjectiveIndex = -1 Then GoTo TryAgain
               'Case 5, 6
               '   .ObjectiveType = "DsSh"
               '   .ObjectiveIndex = ClosestShip(RefShip, Hostile)
               '   If .ObjectiveIndex = -1 Then GoTo TryAgain
               End Select
            End If
         End If
      End If
      
   End With

End Sub


Private Sub DoKeyOrThinkAction()

   With Ships(RefShip)
      
      If .MoveUp Then ChangeMod TotalAcceleration(RefShip)
      If .InHyperspace Then ChangeMod ShipTypes(.ShipType).HyperspaceAccel + 0.3 * .maMod
      If .MoveRight And Not .MoveLeft Then .Spin = .Spin + TotalSpinAcceleration(RefShip)
      If .MoveLeft And Not .MoveRight Then .Spin = .Spin - TotalSpinAcceleration(RefShip)
      If .MoveDown Then .maMod = .maMod * ShipTypes(.ShipType).FrictionRatio ^ (10 * .maMod)
      If .Align Then .Spin = .Spin * ShipTypes(.ShipType).SpinFrictionRatio ^ (10 * Abs(.Spin))
      
   End With

End Sub


Private Sub DoGravity()
Dim ModArg As Pol2
Dim iStellarObject As Integer

   With Ships(RefShip)
   
      For iStellarObject = 0 To UBound(StellarObjects)
         If StellarObjects(iStellarObject).System = .System Then
            If DistanceBetween(NewRect2(StellarObjects(iStellarObject).x, StellarObjects(iStellarObject).y), NewRect2(.x, .y)) < (ShipTypes(.ShipType).Size + StellarObjects(iStellarObject).Size) / 2 Then Exit Sub
         End If
      Next iStellarObject
   
      ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NetGravityAt(.System, NewRect2(.x, .y)))
      .maMod = ModArg.M
      .maArg = ModArg.A
      
   End With

End Sub

Private Function DoDoom() As Boolean
   
   With Ships(RefShip)
   
      If .Hull <= 0 Then
         MakeExplosion CSng(.x), CSng(.y), .System, ShipTypes(.ShipType).Size / 2, 5, 20, CatalystEx, 1
         .Died = True
         DoDoom = True
      End If
      
   End With

End Function


Private Sub DoDamage()
Dim RefProjectile As Integer
Dim ModArg As Pol2

   With Ships(RefShip)
      
      ' Damage from projectiles
      For RefProjectile = 0 To nProjectiles
         If Projectiles(RefProjectile).Exists Then
            ' If in the same system
            If Projectiles(RefProjectile).System = .System Then
               ' If active
               If Projectiles(RefProjectile).RemainingTimeTillHot = 0 Then
                  ' If in range
                  If Sqr((Projectiles(RefProjectile).x - .x) ^ 2 + (Projectiles(RefProjectile).y - .y) ^ 2) - ShipTypes(.ShipType).Size / 2 - ProjectileTypes(Projectiles(RefProjectile).ProjectileType).Size / 2 <= ProjectileTypes(Projectiles(RefProjectile).ProjectileType).TerminalProximity Then
                     
                     ' Explosion
                     If ProjectileTypes(Projectiles(RefProjectile).ProjectileType).BlastExists Then
                        MakeExplosion Projectiles(RefProjectile).x, Projectiles(RefProjectile).y, Projectiles(RefProjectile).System, ProjectileTypes(Projectiles(RefProjectile).ProjectileType).InitialBlastSize, ProjectileTypes(Projectiles(RefProjectile).ProjectileType).DeltaBlastSize, ProjectileTypes(Projectiles(RefProjectile).ProjectileType).BlastTimeLength, CatalystEx, ProjectileTypes(Projectiles(RefProjectile).ProjectileType).BlastVolume
                     End If
                     
                     ' Shield + Hull
                     .Shield = .Shield - ProjectileTypes(Projectiles(RefProjectile).ProjectileType).ShieldDamage
                     If .Shield <= 0 Then .Hull = .Hull - ProjectileTypes(Projectiles(RefProjectile).ProjectileType).HullDamage
                     
                     If .Hull < 0 And Projectiles(RefProjectile).OwnorShip <> -1 Then
                        IncKills Projectiles(RefProjectile).OwnorShip
                     End If
                     
                     ' Tractor
                     ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(ProjectileTypes(Projectiles(RefProjectile).ProjectileType).TractorForce / ShipTypes(.ShipType).Mass, Projectiles(RefProjectile).maBearing + 180))
                     .maMod = ModArg.M
                     .maArg = ModArg.A
                     
                     Projectiles(RefProjectile).Exists = False
                     
                     ' Relations
                     If Projectiles(RefProjectile).OwnorShip <> -1 Then
                        If ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Neutral Or _
                        ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Forbiddon Then
                           ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Hostile
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next RefProjectile
      
Damagefromexplosions:
      ' Damage from explosions
      '''
      
SmokeAndSparks:
      ' Smoke and sparks
      If .Hull < ShipTypes(.ShipType).MaxHull And Int(Rnd * (.Hull)) < 1 And Not .Died Then
         MakeExplosion CSng(.x + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), CSng(.y + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), .System, 10, -1, 6, SmokeEx, 0
         If Int(Rnd * 10) = 0 Then
            MakeExplosion CSng(.x + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), CSng(.y + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), .System, 6, 1, 4, CatalystEx, 0.1
            '.Hull = .Hull - 1
         End If
      End If
      
   End With

End Sub


Private Sub DrawJetstream()
Dim JetSize As Integer
Dim Size As Single
Dim ModFromCenter As Single
Dim x As Single
Dim y As Single

EnableBlendOne

   With Ships(RefShip)
      
      If .MoveUp And .System = Ships(You.Ship).System And .FuelLeft > 0 Then
         JetSize = ShipTypes(.ShipType).Size / 9
         If .AfterburnerOn Then JetSize = 1.3 * JetSize
         
         For i = 0 To 10
            ModFromCenter = ShipTypes(.ShipType).Size / 4 + Rnd * ShipTypes(.ShipType).Size / 3
            If .AfterburnerOn Then ModFromCenter = ModFromCenter * 1.2
            
            Size = JetSize * ((ShipTypes(.ShipType).Size * (1 / 4 + 1 / 3)) / ModFromCenter)
            x = .x - PolToX(ModFromCenter, Mod360(.maBearing))
            y = .y - PolToY(ModFromCenter, Mod360(.maBearing))
            DrawTexture mTextures.txrFlares(1), NewfRECT(1, 0, 1, 0), _
               NewfRECT(y - Size / 2, y + Size / 2, x - Size / 2, x + Size / 2), _
               0, True, NewColour(255 / 3, 255 / 9, 0, 0)
         Next i 'Mod360(.maBearing + 180) ' CartToArg(.LastX - .X, .LastY - .Y)
      End If
      
   End With

EnableBlendNormal

End Sub


Private Sub DoFuelShieldCloakBattery()

   With Ships(RefShip)
      
      ' Shield
      If .Shield < ShipTypes(.ShipType).MaxShield And .Battery >= ShipTypes(.ShipType).ShieldRechargePowerCost Then
         .Shield = .Shield + ShipTypes(.ShipType).ShieldRechargeValue
         .Battery = .Battery - ShipTypes(.ShipType).ShieldRechargePowerCost
      End If
      .Shield = Bound(.Shield, ShipTypes(.ShipType).MaxShield, 0)
      
      ' Cloak
      If .CloakOn Then
         ' Cloak draining
         .Cloak = .Cloak - ShipTypes(.ShipType).CloakDrainRate
         If .Cloak < 0 Then
            .CloakOn = False
         End If
      Else
         ' Cloak recharging
         If .Cloak < ShipTypes(.ShipType).MaxCloak And .Battery >= ShipTypes(.ShipType).CloakRechargePowerCost Then
            .Cloak = .Cloak + ShipTypes(.ShipType).CloakRechargeValue
            .Battery = .Battery - ShipTypes(.ShipType).CloakRechargePowerCost
         End If
      End If
      .Cloak = Bound(.Cloak, ShipTypes(.ShipType).MaxCloak, 0)
      
      ' Fuel
      If .FuelLeft > 0 Then
         If .MoveUp Then
            .FuelLeft = .FuelLeft - ShipTypes(.ShipType).Acceleration
            If .AfterburnerOn Then
               .FuelLeft = .FuelLeft - ShipTypes(.ShipType).AfterburnerAcceleration
            End If
         End If
         If .MoveLeft Or .MoveRight Then
            .FuelLeft = .FuelLeft - ToRadians(ShipTypes(.ShipType).SpinAcceleration) * (ShipTypes(.ShipType).Size / 5)
         End If
      End If
      ' refuel at planet or station
      If .CurrentStellarObject <> -1 Then
         If StellarObjects(.CurrentStellarObject).Government = .Government Then
            .FuelLeft = .FuelLeft + 5
         End If
      End If
      .FuelLeft = Bound(.FuelLeft, ShipTypes(.ShipType).MaxFuel, 0)
      
      ' Battery / generator
      If .Battery < ShipTypes(.ShipType).MaxBattery And .FuelLeft >= ShipTypes(.ShipType).BatteryReGenerateFuelCost Then
         .Battery = .Battery + ShipTypes(.ShipType).BatteryReGenerateValue
         .FuelLeft = .FuelLeft - ShipTypes(.ShipType).BatteryReGenerateFuelCost
      End If
      .Battery = Bound(.Battery, ShipTypes(.ShipType).MaxBattery, 0)
   
   End With

End Sub


Private Sub DoHeadingAndBearing()

   With Ships(RefShip)
      
      .maBearing = Mod360(.maBearing)
      .maArg = Mod360(.maArg)
      
   End With

End Sub


Private Sub DoNewLocation()

    With Ships(RefShip)
        
        .LastX = .x
        .LastY = .y
        .x = .x + PolToX(.maMod, .maArg)
        .y = .y + PolToY(.maMod, .maArg)
        
    End With

End Sub


Function HyperspaceTo(ByVal System As Integer, Optional ByVal pShip As Integer = -1) As eHyperspaceToReturn

   If System = -1 Then Exit Function
   
   If pShip = -1 Then pShip = RefShip
   
   With Ships(pShip)
   
      If .InHyperspace Then
         HyperspaceTo = Already_Entering_Hyperspace
         Exit Function
      End If
      
      If Sqr(.x ^ 2 + .y ^ 2) < Systems(.System).HyperspaceDepartDistance Then
         If pShip = You.Ship Then DisplayMessage "Not far enough from system center", White
         HyperspaceTo = Not_Far_Enough_From_Center
         Exit Function
      End If
      
      HyperspaceTo = Entering_Hyperspace
      .HyperspaceDestination = System
      .InitialHyperspaceCruiseDistanceLeft = Sqr((Systems(.System).x - Systems(.HyperspaceDestination).x) ^ 2 + (Systems(.System).y - Systems(.HyperspaceDestination).y) ^ 2)
      .HyperspaceCruiseDistanceLeft = .InitialHyperspaceCruiseDistanceLeft
      
      ' if not enough fuel then abort
      If .FuelLeft < .InitialHyperspaceCruiseDistanceLeft Then
         If pShip = You.Ship Then DisplayMessage "Insufficient fuel to complete jump to " & Systems(.HyperspaceDestination).Name, White
         .HyperspaceDestination = -1
         HyperspaceTo = Not_Enough_Fuel
      End If
   
   End With

End Function


Sub DoHyperspace()

   With Ships(RefShip)
      
      .MoveUp = False
      .MoveDown = True
      
      If Not TurnToBearing(CartToArg(-(Systems(.System).x - Systems(.HyperspaceDestination).x), Systems(.System).y - Systems(.HyperspaceDestination).y)) And Not .InHyperspace Then Exit Sub
      
      .MoveDown = False
      
      .InHyperspace = True
      
      If Not EnteredHyperspace Then Exit Sub
      
      If CruisingHyperspace Then Exit Sub
      
      .System = .HyperspaceDestination
      
      .x = PolToX(Systems(.HyperspaceDestination).HyperspaceArriveDistance, .maArg + 180)
      .y = PolToY(Systems(.HyperspaceDestination).HyperspaceArriveDistance, .maArg + 180)
      
      .maMod = 100
      
      mTabSelect.Reset RefShip
      
      .InHyperspace = False
      .HyperspaceDestination = -1
      
      ' Player only
      If RefShip = You.Ship Then
         mStars.InitialPhysics
         mSounds.Play sndJumpArive
      Else
         If ShipRelations(You.Ship, RefShip) = Hostile And Ships(You.Ship).System = Ships(RefShip).System Then
            mSounds.Play sndHostileJumpArive
         End If
      End If
      
   End With

End Sub


Sub DoShipRefs()

   With Ships(RefShip)
      
      If .CurrentShipSelection <> -1 Then
         If Ships(.CurrentShipSelection).Died Or Ships(.CurrentShipSelection).System <> Ships(.CurrentShipSelection).System Then
            .CurrentShipSelection = -1
         End If
      End If
      
      If .OwnorShip <> -1 Then
         If Ships(.OwnorShip).Died Then
            .OwnorShip = -1
         End If
      End If
      
   End With

End Sub


Private Sub ChangeMod(ByVal Rate As Single)
Dim ModArg As Pol2

   With Ships(RefShip)
      
      ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(Rate, .maBearing))
      .maMod = ModArg.M
      .maArg = ModArg.A
      
   End With

End Sub


Function TotalAcceleration(ByVal RefShip As Integer) As Single

   With Ships(RefShip)
   
      If .FuelLeft > 0 Then
         TotalAcceleration = ShipTypes(.ShipType).Acceleration
         If .AfterburnerOn Then
            TotalAcceleration = TotalAcceleration + ShipTypes(.ShipType).AfterburnerAcceleration
         End If
      End If
      
   End With

End Function


Function TotalSpinAcceleration(ByVal RefShip As Integer) As Single

   With Ships(RefShip)
   
      If .FuelLeft > 0 Then
         TotalSpinAcceleration = ShipTypes(.ShipType).SpinAcceleration
      End If
      
   End With

End Function


Private Function TurnToBearing(ByVal Bearing As Single) As Boolean

   With Ships(RefShip)
   
      Bearing = DifferenceBetweenAngles(Mod360(.maBearing), Mod360(Bearing))
      If Abs(Bearing) <= 1 And Abs(.Spin) < ShipTypes(.ShipType).SpinAcceleration Then
         .Spin = 0
         .maBearing = .maBearing + Bearing
         TurnToBearing = True
      Else
         .MoveLeft = False
         .MoveRight = False
         .Align = Abs(Bearing) < 5
         If Bearing < 0 Then
            .MoveLeft = True
         Else
            .MoveRight = True
         End If
         .Spin = BoundMax(Abs(Bearing) / 2, 1) * .Spin
      End If
      
   End With

End Function


Private Function EnteredHyperspace() As Boolean

   With Ships(RefShip)
      
      If You.Ship = RefShip Then
         mSounds.Play sndJumpLeave
      ElseIf ShipRelations(You.Ship, RefShip) = Hostile And Ships(You.Ship).System = Ships(RefShip).System Then
         mSounds.Play sndHostileJumpLeave
      End If
      
      If .maMod > ShipTypes(.ShipType).HyperspaceCruiseSpeed Then
         EnteredHyperspace = True
         If RefShip = You.Ship Then Play sndJumpLeave
      Else
         If RefShip = You.Ship Then
            mStars.BackColour = Int((.maMod / ShipTypes(.ShipType).HyperspaceCruiseSpeed) * 155)
            mStars.BackColour = RGB(mStars.BackColour, mStars.BackColour, mStars.BackColour)
         End If
      End If
      
   End With

End Function


Private Function CruisingHyperspace() As Boolean

   With Ships(RefShip)
      
      .maMod = ShipTypes(.ShipType).HyperspaceCruiseSpeed
      
      .HyperspaceCruiseDistanceLeft = .HyperspaceCruiseDistanceLeft - ShipTypes(.ShipType).HyperspaceCruiseSpeed / 10
      
      ' burn fuel
      .FuelLeft = .FuelLeft - ShipTypes(.ShipType).HyperspaceCruiseSpeed / 10
      
      If .HyperspaceCruiseDistanceLeft > 0 Then
         If RefShip = You.Ship Then
            mStars.BackColour = RGB(155, 155, 155)
         End If
         CruisingHyperspace = True
      End If
      
   End With

End Function

Sub NewShips()
Dim TempRefShip As Integer
Dim RSO As Integer

   If mMonitor.LastFPS > mGame.MinFPS Then
      If Int(Rnd * 20) = 0 Then
         RSO = RandomSO(RandomSys(RandomGxy))
         If RSO <> -1 Then
            With StellarObjects(RSO)
            
               If .Landable Then
TryAgain:
                  i = Int(Rnd * (UBound(ShipTypes) + 1))
                  If ShipTypes(i).SpeciesSpecific <> -1 And ShipTypes(i).SpeciesSpecific <> Governments(.Government).Species Then GoTo TryAgain
                  If ShipTypes(i).GovernmentSpecific <> -1 And ShipTypes(i).GovernmentSpecific <> .Government Then GoTo TryAgain
                  TempRefShip = NewShip(i, ClosestShipFromSO(RSO, .Government), .System, .x + Int(Rnd * .Size) - .Size / 2, .y + Int(Rnd * .Size) - .Size / 2, 0, 0, Rnd * 360, .Government, Privateer, , 1000)
               End If
            
            End With
         End If
      End If
   End If

End Sub

Sub SelectDefaultGuns(ByRef iShip As Integer)
Dim TmpGuns() As Variant

   With Ships(iShip)
   
      ReDim TmpGuns(UBound(Split(ShipTypes(.ShipType).DefaultGunTypes, ",")))
      ' Select default weapons
      For i = 0 To UBound(TmpGuns)
         TmpGuns(i) = NewGun(Split(ShipTypes(.ShipType).DefaultGunTypes, ",")(i))
      Next i
      .Guns = Join(TmpGuns, ",")
   
   End With

End Sub

Sub IncKills(ByVal iShip As Integer, Optional ByVal nKills As Long = 1)
   
   If Ships(iShip).OwnorShip <> -1 Then
      IncKills Ships(iShip).OwnorShip, nKills
   End If
   
   Ships(iShip).Kills = Ships(iShip).Kills + nKills

End Sub

Function TopSpeed(ByVal ShipType As Integer) As Single

   With ShipTypes(ShipType)
   
      TopSpeed = -.Acceleration / (.FrictionRatio - 1)
   
   End With

End Function

Function Image(ByVal ShipType As Integer, ByVal f0_360 As Single) As Direct3DBaseTexture8

   If ShipImageSets(ShipTypes(ShipType).ShipImage).FlipX Then
      If f0_360 > 180 Then
         f0_360 = 180 - (f0_360 - 180)
      End If
   End If
   
   f0_360 = Round(f0_360 / ShipImageSets(ShipTypes(ShipType).ShipImage).DeltaDegs, 0)
   
   If f0_360 = ShipImageSets(ShipTypes(ShipType).ShipImage).Frames + 1 Then f0_360 = 0
   
   Set Image = ShipImageSets(ShipTypes(ShipType).ShipImage).Image(f0_360)

End Function

Function NetGravityAt(ByVal System As Integer, ByRef Location As Rect2) As Pol2
Dim iStellarObject As Integer
Dim GravityTemp As Single
Dim Distance As Single

   For iStellarObject = 0 To UBound(StellarObjects)
      If StellarObjects(iStellarObject).System = System Then
         Distance = DistanceBetween(Location, NewRect2(StellarObjects(iStellarObject).x, StellarObjects(iStellarObject).y))
         GravityTemp = BoundMax(StellarObjects(iStellarObject).GravitationalFieldStrength / (Distance / 1000) ^ 2, StellarObjects(iStellarObject).MaxGravityAcceleration)
         NetGravityAt = Pol2Pol2Add(NetGravityAt, NewPol2(GravityTemp, CartToArg(StellarObjects(iStellarObject).x - Location.x, StellarObjects(iStellarObject).y - Location.y)))
      End If
   Next iStellarObject

End Function
                    ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(ProjectileTypes(Projectiles(RefProjectile).ProjectileType).TractorForce / ShipTypes(.ShipType).Mass, Projectiles(RefProjectile).maBearing + 180))
                     .maMod = ModArg.M
                     .maArg = ModArg.A
                     
                     Projectiles(RefProjectile).Exists = False
                     
                     ' Relations
                     If Projectiles(RefProjectile).OwnorShip <> -1 Then
                        If ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Neutral Or _
                        ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Forbiddon Then
                           ShipRelations(RefShip, Projectiles(RefProjectile).OwnorShip) = Hostile
                        End If
                     End If
                  End If
               End If
            End If
         End If
      Next RefProjectile
      
Damagefromexplosions:
      ' Damage from explosions
      '''
      
SmokeAndSparks:
      ' Smoke and sparks
      If (.Hull / ShipTypes(.ShipType).MaxHull) < 0.5 And Not .Died Then
         If (Rnd * ((.Hull / ShipTypes(.ShipType).MaxHull) / 0.5) * 10) < 1 Then
            MakeExplosion CSng(.x + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), CSng(.y + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), .system, 10, -1, 6, SmokeEx, 0
            If Int(Rnd * 5) = 0 Then
               'MakeExplosion CSng(.x + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), CSng(.y + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), .system, 6, 1, 4, CatalystEx, 0.05
               MakeFlurry CSng(.x + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), CSng(.y + Rnd * ShipTypes(.ShipType).Size / 3 - ShipTypes(.ShipType).Size / 6), .system, 8, Int(ShipTypes(.ShipType).Size / 10), Cyan, Blue 'come out as opposite colours Yellow and Red
               '.Hull = .Hull - 1
            End If
         End If
      End If
      
   End With

End Sub


Private Sub DrawJetstream()
Dim JetSize As Integer
Dim Size As Single
Dim ModFromCenter As Single
Dim x As Single
Dim y As Single

Dim ModArg As Pol2

EnableBlendOne

   With Ships(RefShip)
      
      If .MoveUp And .system = Ships(You.Ship).system And .FuelLeft > 0 Then
         JetSize = ShipTypes(.ShipType).Size / 9
         If .AfterburnerOn Then JetSize = 1.3 * JetSize
         If .InHyperspace Then
            JetSize = 1.5 * JetSize
            For i = 0 To Int(.maMod / 2)
               ModArg.M = ShipTypes(.ShipType).Size / 2 + Rnd * .maMod
               ModArg.A = Mod360(.maBearing + 180 + Rnd * 18 - 9)
               MakeDust .x + Pol2ToRect2(ModArg).x, _
                  .y + Pol2ToRect2(ModArg).y, _
                  .system, _
                  3, _
                  vbBlue, _
                  vbBlack, _
                  Rnd * 6 - 3, _
                  Rnd * 6 - 3
            Next i
         End If
         
         For i = 0 To 10
            ModFromCenter = ShipTypes(.ShipType).Size / 4 + Rnd * ShipTypes(.ShipType).Size / 3
            If .AfterburnerOn Then ModFromCenter = ModFromCenter * 1.2
            
            Size = JetSize * ((ShipTypes(.ShipType).Size * (1 / 4 + 1 / 3)) / ModFromCenter)
            x = .x - PolToX(ModFromCenter, Mod360(.maBearing))
            y = .y - PolToY(ModFromCenter, Mod360(.maBearing))
            
            If .InHyperspace Then
               DrawTexture mTextures.txrFlares(1), NewfRECT(1, 0, 1, 0), _
                  NewfRECT(y - Size / 2, y + Size / 2, x - Size / 2, x + Size / 2), _
                  0, True, NewColour(0, 255 / 9, 255 / 3, 0)
            Else
               DrawTexture mTextures.txrFlares(1), NewfRECT(1, 0, 1, 0), _
                  NewfRECT(y - Size / 2, y + Size / 2, x - Size / 2, x + Size / 2), _
                  0, True, NewColour(255 / 3, 255 / 9, 0, 0)
            End If
         Next i 'Mod360(.maBearing + 180) ' CartToArg(.LastX - .X, .LastY - .Y)
      End If
      
   End With

EnableBlendNormal

End Sub


Private Sub DoFuelShieldCloakBattery()

   With Ships(RefShip)
      
      ' Shield
      If .Shield < ShipTypes(.ShipType).MaxShield And .Battery >= ShipTypes(.ShipType).ShieldRechargePowerCost Then
         .Shield = .Shield + ShipTypes(.ShipType).ShieldRechargeValue
         .Battery = .Battery - ShipTypes(.ShipType).ShieldRechargePowerCost
      End If
      .Shield = Bound(.Shield, ShipTypes(.ShipType).MaxShield, 0)
      
      ' Cloak
      If .CloakOn Then
         ' Cloak draining
         .Cloak = .Cloak - ShipTypes(.ShipType).CloakDrainRate
         If .Cloak < 0 Then
            .CloakOn = False
         End If
      Else
         ' Cloak recharging
         If .Cloak < ShipTypes(.ShipType).MaxCloak And .Battery >= ShipTypes(.ShipType).CloakRechargePowerCost Then
            .Cloak = .Cloak + ShipTypes(.ShipType).CloakRechargeValue
            .Battery = .Battery - ShipTypes(.ShipType).CloakRechargePowerCost
         End If
      End If
      .Cloak = Bound(.Cloak, ShipTypes(.ShipType).MaxCloak, 0)
      
      ' Fuel
      If .FuelLeft > 0 Then
         If .MoveUp Then
            .FuelLeft = .FuelLeft - ShipTypes(.ShipType).Acceleration
            If .AfterburnerOn Then
               .FuelLeft = .FuelLeft - ShipTypes(.ShipType).AfterburnerAcceleration
            End If
         End If
         If .MoveLeft Or .MoveRight Then
            .FuelLeft = .FuelLeft - ToRadians(ShipTypes(.ShipType).SpinAcceleration) * (ShipTypes(.ShipType).Size / 5)
         End If
      End If
      ' refuel at planet or station
      If .CurrentStellarObject <> -1 Then
         If StellarObjects(.CurrentStellarObject).Government = .Government Then
            .FuelLeft = .FuelLeft + 5
         End If
      End If
      .FuelLeft = Bound(.FuelLeft, ShipTypes(.ShipType).MaxFuel, 0)
      
      ' Battery / generator
      If .Battery < ShipTypes(.ShipType).MaxBattery And .FuelLeft >= ShipTypes(.ShipType).BatteryReGenerateFuelCost Then
         .Battery = .Battery + ShipTypes(.ShipType).BatteryReGenerateValue
         .FuelLeft = .FuelLeft - ShipTypes(.ShipType).BatteryReGenerateFuelCost
      End If
      .Battery = Bound(.Battery, ShipTypes(.ShipType).MaxBattery, 0)
   
   End With

End Sub


Private Sub DoHeadingAndBearing()

   With Ships(RefShip)
      
      .maBearing = Mod360(.maBearing)
      .maArg = Mod360(.maArg)
      
   End With

End Sub


Private Sub DoNewLocation()

    With Ships(RefShip)
        
        .LastX = .x
        .LastY = .y
        .x = .x + PolToX(.maMod, .maArg)
        .y = .y + PolToY(.maMod, .maArg)
        
    End With

End Sub


Function HyperspaceTo(ByVal system As Integer, Optional ByVal pShip As Integer = -1) As eHyperspaceToReturn

   If system = -1 Then Exit Function
   
   If pShip = -1 Then pShip = RefShip
   
   With Ships(pShip)
   
      If .InHyperspace Then
         HyperspaceTo = Already_Entering_Hyperspace
         Exit Function
      End If
      
      If Sqr(.x ^ 2 + .y ^ 2) < Systems(.system).HyperspaceDepartDistance Then
         If pShip = You.Ship Then DisplayMessage "Not far enough from system center", White
         HyperspaceTo = Not_Far_Enough_From_Center
         Exit Function
      End If
      
      HyperspaceTo = Entering_Hyperspace
      .HyperspaceDestination = system
      .HyperspaceCruiseDistance = DistanceSysToSys(.system, .HyperspaceDestination)
      .HyperspaceCruiseDistanceCompleted = 0
      
      ' if not enough fuel then abort
      If .FuelLeft < .HyperspaceCruiseDistance Then
         If pShip = You.Ship Then DisplayMessage "Insufficient fuel to complete jump to " & Systems(.HyperspaceDestination).Name, White
         .HyperspaceDestination = -1
         HyperspaceTo = Not_Enough_Fuel
      End If
   
   End With

End Function


Sub DoHyperspace()

   With Ships(RefShip)
      
      .MoveUp = True
      '.MoveDown = True
      
      ' turn to correct direction and
      If Not .InHyperspace Then
         If TurnToBearing(CartToArg(-(Systems(.system).x - Systems(.HyperspaceDestination).x), Systems(.system).y - Systems(.HyperspaceDestination).y)) Then
            .MoveDown = False
            .InHyperspace = True
            If You.Ship = RefShip Then
               mSounds.Play sndJumpLeave
            ElseIf ShipRelations(You.Ship, RefShip) = Hostile And Ships(You.Ship).system = Ships(RefShip).system Then
               mSounds.Play sndHostileJumpLeave
            End If
            Exit Sub
         End If
      End If
      
      TurnToBearing (CartToArg(-(Systems(.system).x - Systems(.HyperspaceDestination).x), Systems(.system).y - Systems(.HyperspaceDestination).y))
      
      If .maMod > ShipTypes(.ShipType).HyperspaceCruiseSpeed Then
         .maMod = ShipTypes(.ShipType).HyperspaceCruiseSpeed
      End If
      
      If .maMod = ShipTypes(.ShipType).HyperspaceCruiseSpeed Then
         .HyperspaceCruiseDistanceCompleted = .HyperspaceCruiseDistanceCompleted + ShipTypes(.ShipType).HyperspaceCruiseSpeed / 10
      End If
      
      If .HyperspaceCruiseDistanceCompleted >= DistanceSysToSys(.system, .HyperspaceDestination) Then
         .system = .HyperspaceDestination

         .x = PolToX(Systems(.HyperspaceDestination).HyperspaceArriveDistance, .maArg + 180)
         .y = PolToY(Systems(.HyperspaceDestination).HyperspaceArriveDistance, .maArg + 180)
         
         MakeFlurry CSng(.x), CSng(.y), .system, 20, ShipTypes(.ShipType).Size * 2, Blue, DarkBlue
         MakeFlurry CSng(.x), CSng(.y), .system, 10, ShipTypes(.ShipType).Size, LightBlue, Blue
         
         .maMod = 100
         
         mTabSelect.Reset RefShip
         
         .InHyperspace = False
         .FuelLeft = .FuelLeft - .HyperspaceCruiseDistance
         .HyperspaceDestination = -1
         .HyperspaceCruiseDistance = 0
         .HyperspaceCruiseDistanceCompleted = 0
         
         ' Player only
         If RefShip = You.Ship Then
            mStars.InitialPhysics
            mSounds.Play sndJumpArive
         Else
            If ShipRelations(You.Ship, RefShip) = Hostile And Ships(You.Ship).system = Ships(RefShip).system Then
               mSounds.Play sndHostileJumpArive
            End If
         End If
      End If
      
   End With

End Sub


Sub DoShipRefs()

   With Ships(RefShip)
      
      If .CurrentShipSelection <> -1 Then
         If Ships(.CurrentShipSelection).Died Or Ships(.CurrentShipSelection).system <> Ships(.CurrentShipSelection).system Then
            .CurrentShipSelection = -1
         End If
      End If
      
      If .OwnorShip <> -1 Then
         If Ships(.OwnorShip).Died Then
            .OwnorShip = -1
         End If
      End If
      
   End With

End Sub


Private Sub ChangeMod(ByVal Rate As Single)
Dim ModArg As Pol2

   With Ships(RefShip)
      
      ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(Rate, .maBearing))
      .maMod = ModArg.M
      .maArg = ModArg.A
      
   End With

End Sub


Function TotalAcceleration(ByVal RefShip As Integer) As Single

   With Ships(RefShip)
   
      If .FuelLeft > 0 Then
         TotalAcceleration = ShipTypes(.ShipType).Acceleration
         If .AfterburnerOn Then
            TotalAcceleration = TotalAcceleration + ShipTypes(.ShipType).AfterburnerAcceleration
         End If
      End If
      
   End With

End Function


Function TotalSpinAcceleration(ByVal RefShip As Integer) As Single

   With Ships(RefShip)
   
      If .FuelLeft > 0 Then
         TotalSpinAcceleration = ShipTypes(.ShipType).SpinAcceleration
      End If
      
   End With

End Function


Private Function TurnToBearing(ByVal Bearing As Single) As Boolean

   With Ships(RefShip)
   
      Bearing = DifferenceBetweenAngles(Mod360(.maBearing), Mod360(Bearing))
      If Abs(Bearing) <= 1 And Abs(.Spin) < ShipTypes(.ShipType).SpinAcceleration Then
         .Spin = 0
         .maBearing = .maBearing + Bearing
         TurnToBearing = True
      Else
         .MoveLeft = False
         .MoveRight = False
         .Align = Abs(Bearing) < 5
         If Bearing < 0 Then
            .MoveLeft = True
         Else
            .MoveRight = True
         End If
         .Spin = BoundMax(Abs(Bearing) / 2, 1) * .Spin
      End If
      
   End With

End Function

Public Sub NewShips()
Dim TempRefShip As Integer
Dim RSO As Integer

   If mMonitor.LastFPS > mGame.MinFPS Then
      'If Int(Rnd * 50) = 0 Then
         RSO = RandomSO(RandomSys(RandomGxy))
         If RSO <> -1 Then
            With StellarObjects(RSO)
            
               If .Landable Then
TryAgain:
                  i = Int(Rnd * (UBound(ShipTypes) + 1))
                  If ShipTypes(i).SpeciesSpecific <> -1 And ShipTypes(i).SpeciesSpecific <> Governments(.Government).Species Then GoTo TryAgain
                  If ShipTypes(i).GovernmentSpecific <> -1 And ShipTypes(i).GovernmentSpecific <> .Government Then GoTo TryAgain
                  TempRefShip = NewShip(i, ClosestShipFromSO(RSO, .Government), .system, .x + Int(Rnd * .Size) - .Size / 2, .y + Int(Rnd * .Size) - .Size / 2, 0, 0, Rnd * 360, .Government, Millitary, , 1000)
               End If
            
            End With
         End If
      'End If
   End If

End Sub

Sub SelectDefaultGuns(ByRef iShip As Integer)
Dim TmpGuns() As Variant

   With Ships(iShip)
   
      ReDim TmpGuns(UBound(Split(ShipTypes(.ShipType).DefaultGunTypes, ",")))
      ' Select default weapons
      For i = 0 To UBound(TmpGuns)
         TmpGuns(i) = NewGun(Split(ShipTypes(.ShipType).DefaultGunTypes, ",")(i))
      Next i
      .Guns = Join(TmpGuns, ",")
   
   End With

End Sub

Sub IncKills(ByVal iShip As Integer, Optional ByVal nKills As Long = 1)
   
   If Ships(iShip).OwnorShip <> -1 Then
      IncKills Ships(iShip).OwnorShip, nKills
   End If
   
   Ships(iShip).Kills = Ships(iShip).Kills + nKills

End Sub

Function TopSpeed(ByVal ShipType As Integer) As Single

   With ShipTypes(ShipType)
   
      TopSpeed = -.Acceleration / (.FrictionRatio - 1)
   
   End With

End Function

Function Image(ByVal ShipType As Integer, ByVal f0_360 As Single) As Direct3DBaseTexture8

   If ShipImageSets(ShipTypes(ShipType).ShipImage).FlipX Then
      If f0_360 > 180 Then
         f0_360 = 180 - (f0_360 - 180)
      End If
   End If
   
   f0_360 = Round(f0_360 / ShipImageSets(ShipTypes(ShipType).ShipImage).DeltaDegs, 0)
   
   If f0_360 = ShipImageSets(ShipTypes(ShipType).ShipImage).Frames + 1 Then f0_360 = 0
   
   Set Image = ShipImageSets(ShipTypes(ShipType).ShipImage).Image(f0_360)

End Function

Function NetGravityAt(ByVal system As Integer, ByRef Location As Rect2) As Pol2
Dim iStellarObject As Integer
Dim GravityTemp As Single
Dim Distance As Single
   
   For iStellarObject = 0 To UBound(StellarObjects)
      With StellarObjects(iStellarObject)
         If .system = system Then
            If DistanceBetween(NewRect2(.x, .y), Location) < .Size / 2 Then
               NetGravityAt = NewPol2(0, 0)
               Exit Function ' there is definitely no force of gravity here
            End If
            Distance = DistanceBetween(Location, NewRect2(.x, .y))
            GravityTemp = BoundMax(.GravitationalFieldStrength / (Distance / 1000) ^ 2, .MaxGravityAcceleration)
            NetGravityAt = Pol2Pol2Add(NetGravityAt, NewPol2(GravityTemp, CartToArg(.x - Location.x, .y - Location.y)))
         End If
      End With
   Next iStellarObject

End Function
on As Rect2) As Pol2
Dim iStellarObject As Integer
Dim GravityTemp As Single
Dim Distance As Single
   
   For iStellarObject = 0 To UBound(StellarObjects)
      With StellarObjects(iStellarObject)
         If .system = system Then
            If DistanceBetween(NewRect2(.X, .Y), Location) < .Size / 2 Then
               NetGravityAt = NewPol2(0, 0)
               Exit Function ' there is definitely no force of gravity here
            End If
            Distance = DistanceBetween(Location, NewRect2(.X, .Y))
            GravityTemp = BoundMax(.GravitationalFieldStrength / (Distance / 1000) ^ 2, .MaxGravityAcceleration)
            NetGravityAt = Pol2Pol2Add(NetGravityAt, NewPol2(GravityTemp, CartToArg(.X - Location.X, .Y - Location.Y)))
         End If
      End With
   Next iStellarObject

End Function
