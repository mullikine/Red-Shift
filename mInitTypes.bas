Attribute VB_Name = "mInitTypes"
Option Explicit

Public ShipTypes() As tShipType
Public GunTypes() As tGunType
Public ProjectileTypes() As tProjectileType

Sub Init()

   InitProjectileTypes
   InitGunTypes
   InitShipTypes
   InitShips

End Sub

Private Sub InitShips()

   ReDim Ships(nShips) As tShip

End Sub

Private Sub InitShipTypes()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Ship Types\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim ShipTypes(aInput(1))

   For i = 0 To UBound(ShipTypes)
      With ShipTypes(i)
      
         Open App.Path & "\Data\Classes\Ship Types\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .ClassName = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .SpeciesSpecific = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .GovernmentSpecific = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ShipImage = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxShield = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxHull = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxFuel = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxCloak = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxBattery = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .BatteryReGenerateValue = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .BatteryReGenerateFuelCost = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .CloakDrainRate = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .CloakRechargePowerCost = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .CloakRechargeValue = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ShieldRechargePowerCost = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ShieldRechargeValue = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SpinFrictionRatio = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Acceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .AfterburnerAcceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SpinAcceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .FrictionRatio = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Size = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Mass = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HyperspaceCruiseSpeed = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HyperspaceAccel = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .DefaultGunTypes = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .GunPositionsX = aInput(1)
            aInput = Tokenize(ReadStr(1), " ")
            .GunPositionsY = aInput(1)
         Close #1
         
      End With
   Next i

End Sub


Private Sub InitGunTypes()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Gun Types\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim GunTypes(aInput(1))
   
   For i = 0 To UBound(GunTypes)
      With GunTypes(i)
      
         Open App.Path & "\Data\Classes\Gun Types\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .ClassName = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .FireRate = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxAmmo = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ProjectileType = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .PurchasePrice = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SellingPrice = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .RechargeRate = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .RechargePowerCost = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .RandomBearingOffset = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Ballistic = CBool(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .KeyTrigger = Val(aInput(1))
         Close #1
      
      End With
   Next i

End Sub


Private Sub InitProjectileTypes()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Projectile Types\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim ProjectileTypes(aInput(1))

   For i = 0 To UBound(ProjectileTypes)
      With ProjectileTypes(i)
      
         Open App.Path & "\Data\Classes\Projectile Types\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .ClassName = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .WeaponClass = StringToeWeaponClass(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ProjectileImage = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxFuel = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Acceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HLAccelBoost = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .TractorForce = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SpinAcceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SpinFrictionRatio = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .TerminalProximity = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .StartTimeTillHot = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .ShieldDamage = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HullDamage = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Size = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .BlastExists = CBool(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .InitialBlastSize = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .DeltaBlastSize = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .BlastTimeLength = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .FrictionRatio = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Homing = CBool(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .LockOnRange = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Colour = CLng(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .InitialVelocity = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .JetstreamType = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .JetstreamInitSize = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .JetstreamDeltaSize = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .JetstreamLastTime = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .EffectiveRange = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Sound = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Frequency = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Volume = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .BlastVolume = Val(aInput(1))
         Close #1
         
      End With
   Next i

End Sub

