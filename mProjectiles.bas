Attribute VB_Name = "mProjectiles"
'---------------------------------------------------------------------------------------
' Module    : mProjectiles
' DateTime  : 2/20/2005 10:23
' Author    : Shane Mulligan
' Purpose   : Implements the projectile class
'---------------------------------------------------------------------------------------

Option Explicit

Public Projectiles() As tProjectile
Public nProjectiles As Integer

Public RefProjectile As Integer
Public MinProjectiles As Integer

Sub Init()

   MinProjectiles = 10
   FormatProjectiles MinProjectiles

End Sub

Sub ReboundProjectiles(ByVal pnProjectiles As Integer)

   nProjectiles = pnProjectiles
   
   ReDim Preserve Projectiles(nProjectiles)
   
End Sub

Sub ProjectileGarbageCollect()

   If Not Projectiles(nProjectiles).Exists And nProjectiles > MinProjectiles Then
      ReboundProjectiles nProjectiles - 1
   End If

End Sub

Sub FormatProjectiles(ByVal pnProjectiles As Integer)

   nProjectiles = pnProjectiles

   ReDim Projectiles(nProjectiles) As tProjectile

End Sub

Sub DoPhysics()

   For RefProjectile = 0 To nProjectiles
   
      With Projectiles(RefProjectile)
   
         If .Exists And Not DoDoom Then
            DoShipRefs
            DoHome
            DoSpin
            DoThrust
            DoFriction
            DoHeadingAndBearing
            DoNewLocation
            DoTimeTillHot
            DoJetstream
         End If
      
      End With
      
   Next RefProjectile
   
   ProjectileGarbageCollect

End Sub

Private Function DoShipRefs()

   With Projectiles(RefProjectile)
   
      If .TargetShip <> -1 Then
         If Ships(.TargetShip).Died Or .System <> Ships(.TargetShip).System Then
            .TargetShip = -1
         End If
      End If
      If .OwnorShip <> -1 Then
         If Ships(.OwnorShip).Died Then
            .OwnorShip = -1
         End If
      End If
   
   End With

End Function

Private Function DoDoom() As Boolean

   With Projectiles(RefProjectile)
   
      If .FuelLeft <= 0 And .Exists Then
         Select Case ProjectileTypes(.ProjectileType).WeaponClass
         Case Projectile
            MakeExplosion .x, .y, .System, ProjectileTypes(.ProjectileType).Size / 4, -0.5, 7, CatalystEx, ProjectileTypes(.ProjectileType).BlastVolume, &HFFC0C0C0
         End Select
         .Exists = False
      End If
      
      DoDoom = Not .Exists
   
   End With

End Function


Private Sub DoThrust()

   ChangeMod TotalAcceleration(RefProjectile)

End Sub

Function TotalAcceleration(ByVal pProjectile As Integer) As Single

   With Projectiles(pProjectile)
   
      If .FuelLeft > 0 Then
         TotalAcceleration = ProjectileTypes(.ProjectileType).Acceleration
         If ProjectileTypes(.ProjectileType).Homing And .TargetShip <> -1 Then
            If DistBtwnShipPjtl(.TargetShip, pProjectile) <= ProjectileTypes(.ProjectileType).LockOnRange Then
               TotalAcceleration = TotalAcceleration + ProjectileTypes(.ProjectileType).HLAccelBoost
               .FuelLeft = .FuelLeft - 1
            End If
         End If
         .FuelLeft = .FuelLeft - 1
      End If
      
   End With

End Function

Private Sub DoHome()
Dim fArg As Single

   With Projectiles(RefProjectile)
   
      If ProjectileTypes(.ProjectileType).Homing And .TargetShip <> -1 Then
         If DistBtwnShipPjtl(.TargetShip, RefProjectile) <= ProjectileTypes(.ProjectileType).LockOnRange Then
            If Not Ships(.TargetShip).CloakOn Then
               fArg = CartToArg(Ships(.TargetShip).x - .x, Ships(.TargetShip).y - .y)
               TurnToBearing fArg
            End If
         End If
      End If
      
   End With

End Sub

Private Sub DoFriction()

   With Projectiles(RefProjectile)
      
      .maMod = .maMod * ProjectileTypes(.ProjectileType).FrictionRatio ^ .maMod
      .Spin = .Spin * ProjectileTypes(.ProjectileType).SpinFrictionRatio ^ Abs(.Spin)
      
   End With

End Sub


Private Sub DoHeadingAndBearing()

   With Projectiles(RefProjectile)
      
      .maBearing = Mod360(.maBearing)
      .maArg = Mod360(.maArg)
      
   End With

End Sub


Private Sub DoNewLocation()

    With Projectiles(RefProjectile)
        
        .LastX = .x
        .LastY = .y
        .x = .x + PolToX(.maMod, .maArg)
        .y = .y + PolToY(.maMod, .maArg)
        
    End With

End Sub


Private Sub DoTimeTillHot()

   With Projectiles(RefProjectile)
      
      If .RemainingTimeTillHot > 0 And .OwnorShip <> -1 Then
         ' If not touching ownorship...
         If DistBtwnShipPjtl(.OwnorShip, RefProjectile) > Ships(.OwnorShip).maMod * 2 Then
            ' ...then count down
            .RemainingTimeTillHot = .RemainingTimeTillHot - 1
         End If
      End If
       
   End With

End Sub


Sub DrawBodies()
Dim picDeg As Integer

   For RefProjectile = 0 To nProjectiles
      With Projectiles(RefProjectile)
         
         If .Exists Then
            If .System = Ships(You.Ship).System Then
               If ProjectileTypes(.ProjectileType).ProjectileImage <> -1 Then
                  If ZoomMod(DistBtwnShipPjtl(You.Ship, RefProjectile)) < 1100 Then
                     Select Case ProjectileTypes(.ProjectileType).WeaponClass
                     Case Beam
                        EnableBlendOne
                        DrawLine ZoomX(.LastX + SpaceOffset.x), ZoomY(-.LastY + SpaceOffset.y), ZoomX(.x + SpaceOffset.x), ZoomY(-.y + SpaceOffset.y), 0, ProjectileTypes(.ProjectileType).Colour
                     Case Energy, Projectile, Rocket
                        EnableBlendOne
                        
                        picDeg = Round(.maBearing, 0)
                        
                        DrawTexture mProjectiles.Image(.ProjectileType, picDeg), NewfRECT(1, 0, (Not (picDeg > 180)) + 1, (picDeg > 180) + 1), NewfRECT(.y - ProjectileTypes(.ProjectileType).Size / 2, .y + ProjectileTypes(.ProjectileType).Size / 2, .x - ProjectileTypes(.ProjectileType).Size / 2, .x + ProjectileTypes(.ProjectileType).Size / 2), Int(.maBearing) - Round(.maBearing / ProjectileImageSets(ProjectileTypes(.ProjectileType).ProjectileImage).DeltaDegs, 0) * ProjectileImageSets(ProjectileTypes(.ProjectileType).ProjectileImage).DeltaDegs, True, ProjectileTypes(.ProjectileType).Colour
                     End Select
                  End If
               End If
            End If
         End If
         
      End With
   Next RefProjectile

End Sub


Private Sub ChangeMod(ByVal Rate As Single)
Dim ModArg As Pol2

   With Projectiles(RefProjectile)
      
      ModArg = Pol2Pol2Add(NewPol2(.maMod, .maArg), NewPol2(Rate, .maBearing))
      .maMod = ModArg.M
      .maArg = ModArg.A
      
   End With

End Sub


Function MakeProjectile(ByVal OwnorShip As Integer, _
                        ByVal Gun As Integer, _
                        ByVal System As Long, _
                        ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal TargetShip As Integer, _
                        ByVal Bearing As Single) As Boolean

Dim RBO As Single

   j = -1
   For i = 0 To nProjectiles
      If Projectiles(i).Exists = False Then
         j = i
         Exit For
      End If
   Next i
   
   If j = -1 Then
      mProjectiles.ReboundProjectiles nProjectiles + 1
      j = nProjectiles
   End If
   
   RBO = GunTypes(Guns(Split(Ships(OwnorShip).Guns, ",")(Gun)).GunType).RandomBearingOffset
   
   With Projectiles(j)
      .Exists = True
      .ProjectileType = GunTypes(Guns(Split(Ships(OwnorShip).Guns, ",")(Gun)).GunType).ProjectileType
      .System = System
      .OwnorShip = OwnorShip
      .TargetShip = TargetShip
      .x = x
      .y = y
      .maBearing = Bearing + Rnd * RBO - RBO / 2
      .maMod = ProjectileTypes(.ProjectileType).InitialVelocity + Ships(OwnorShip).maMod
      .maArg = .maBearing
      .FuelLeft = ProjectileTypes(.ProjectileType).MaxFuel
      .RemainingTimeTillHot = ProjectileTypes(.ProjectileType).StartTimeTillHot
      
      If (mSounds.MinVolume * (DistanceFromShip(OwnorShip, You.Ship) / mSounds.MaxRange)) / ProjectileTypes(.ProjectileType).Volume > mSounds.MinVolume And .System = Ships(You.Ship).System Then
         If ProjectileTypes(.ProjectileType).Sound > -1 Then
            mSounds.Play ProjectileTypes(.ProjectileType).Sound, (mSounds.MinVolume * (DistanceFromShip(OwnorShip, You.Ship) / mSounds.MaxRange)) / ProjectileTypes(.ProjectileType).Volume, Ships(OwnorShip).x - Ships(You.Ship).x, ProjectileTypes(.ProjectileType).Frequency
         End If
      End If
      
   End With
   
   MakeProjectile = True   ' A projectile was successfully created

End Function

Sub DoSpin()

   With Projectiles(RefProjectile)
   
      .maBearing = .maBearing + .Spin
   
   End With

End Sub

Private Function TurnToBearing(ByVal Bearing As Single) As Boolean

   With Projectiles(RefProjectile)
      
      Bearing = DifferenceBetweenAngles(.maBearing, Mod360(Bearing))
      Turn Bound(Bearing / 45, 1, -1)
      .Spin = BoundMax(Abs(Bearing) / 2, 1) * .Spin
      If Abs(Bearing) <= 1 And Abs(.Spin) < ProjectileTypes(.ProjectileType).SpinAcceleration Then
         .Spin = 0
         .maBearing = .maBearing + Bearing
         TurnToBearing = True
      End If
      
   End With

End Function

Private Sub Turn(ByVal PowerRatio As Single)

   With Projectiles(RefProjectile)
   
      .Spin = .Spin + ProjectileTypes(.ProjectileType).SpinAcceleration * PowerRatio
      
   End With
   
End Sub

Private Sub DoJetstream()
Dim ModArg As Pol2

   With Projectiles(RefProjectile)
      
      If .FuelLeft > 0 And .System = Ships(You.Ship).System Then
         Select Case ProjectileTypes(.ProjectileType).JetstreamType
         Case 0
            For i = 0 To Int(.maMod / 2)
               ModArg.M = ProjectileTypes(.ProjectileType).Size / 2 + Rnd * .maMod
               ModArg.A = Mod360(.maBearing + 180 + Rnd * 18 - 9)
               MakeExplosion .x + Pol2ToRect2(ModArg).x, _
                  .y + Pol2ToRect2(ModArg).y, _
                  .System, ProjectileTypes(.ProjectileType).JetstreamInitSize, _
                  ProjectileTypes(.ProjectileType).JetstreamDeltaSize, ProjectileTypes(.ProjectileType).JetstreamLastTime, SmokeEx, &HFF606060
            Next i
         End Select
      End If
      
   End With

End Sub

Function Image(ByVal ProjectileType As Integer, ByVal i0_360 As Integer) As Direct3DBaseTexture8

   If i0_360 > 180 Then
      i0_360 = 180 - (i0_360 - 180)
   End If
   
   i0_360 = i0_360 / ProjectileImageSets(ProjectileTypes(ProjectileType).ProjectileImage).DeltaDegs
   
   Set Image = ProjectileImageSets(ProjectileTypes(ProjectileType).ProjectileImage).Image(i0_360)

End Function
