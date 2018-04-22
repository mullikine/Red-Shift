Attribute VB_Name = "mDefTypes"
Option Explicit

Public Type tPlayer
   Ship As Integer
   Autopilot As Boolean
End Type

Public Type tProfile
   Name As String
   Kills As Integer
   Credit As Integer
   ShipType As Integer
   System As Integer
   x As Integer
   y As Integer
End Type

'Public Type tThruster
'   ThrusterType As Integer
'   ThrusterOn As Boolean
'End Type

'Public Type tThrusterType
'   ClassName As String
'
'   PurchasePrice As Integer ' In credit
'   SellingPrice As Integer
'
'   PropulsionMethod As ePropulsionMethod
'
'   PropellantsPerProc As Single
'End Type

Public Type tGalaxy
   Name As String
   x As Integer
   y As Integer
End Type

Public Type tSystem
   Name As String
   x As Long
   y As Long
   Galaxy As Integer
   HyperspaceArriveDistance As Integer
   HyperspaceDepartDistance As Integer
   Government As Integer
End Type

Public Type tStellarObject
   Name As String

   Image As Integer

   Kind As eSOKind

   Government As Integer
   System As Integer

   x As Long
   y As Long

' Bitmap length or breadth
   Size As Integer

   GravitationalFieldStrength As Single
   GravGeometricRatio As Single
   MaxGravityAcceleration As Single

   Bearing As Single

' Rot variable frame increment constant
   SpinSpeed As Single

   DescriptionIndex As Integer

   Landable As Boolean
End Type

Public Type tGovernment
   Name As String
   Species As Integer
End Type

Public Type tSpecies
   Name As String
End Type

Public Type tGun
   Hot As Boolean
   ProcLastFired As Long
   Firing As Boolean
   Bearing As Single

   GunType As Integer

   AmmoRemaining As Integer
End Type

Public Type tGunType
   ClassName As String

   PurchasePrice As Integer ' In credit
   SellingPrice As Integer

   ' Gun properties
   ProjectileType As Integer
   FireRate As Integer ' Frames per projection
   MaxAmmo As Integer
   RechargeRate As Integer
   RechargePowerCost As Integer
   RandomBearingOffset As Integer
   Ballistic As Boolean
   KeyTrigger As Integer
End Type

Public Type tProjectile
   ProjectileType As Integer

   System As Integer
   x As Double
   y As Double
   LastX As Double
   LastY As Double

   RemainingTimeTillHot As Integer

   FuelLeft As Single

   maMod As Single
   maArg As Single
   maBearing As Single
   Spin As Single

   OwnorShip As Integer
   TargetShip As Integer

   MoveLeft As Boolean
   MoveRight As Boolean

   Exists As Boolean
End Type

Public Type tProjectileType
   ClassName As String
   ProjectileImage As Integer

   WeaponClass As eWeaponClass

   'Ship qualities
   TerminalProximity As Integer
   StartTimeTillHot As Integer
   ShieldDamage As Single
   HullDamage As Single
   Size As Integer
   BlastExists As Boolean
   InitialBlastSize As Single
   DeltaBlastSize As Single
   BlastTimeLength As Integer
   MaxFuel As Integer
   Acceleration As Single
   HLAccelBoost As Single
   TractorForce As Single
   SpinAcceleration As Single
   SpinFrictionRatio As Single
   FrictionRatio As Currency
   Homing As Boolean
   LockOnRange As Single
   Colour As Long
   InitialVelocity As Single
   JetstreamType As Integer
   JetstreamInitSize As Single
   JetstreamDeltaSize As Single
   JetstreamLastTime As Integer
   EffectiveRange As Single
   Sound As Integer
   Frequency As Single
   Volume As Single
   BlastVolume As Single
End Type

Public Type tShip
   ShipType As Integer
   
   Name As String
   Kills As Long
   Credit As Long
   OwnorShip As Integer
   Government As Integer
   Career As eCareer

'Private pShipRelations() As eRelations

   System As Integer

' Stellar object the ship is gravitationally bound to
   CurrentStellarObject As Integer

   x As Double
   y As Double
   LastX As Double
   LastY As Double

   HyperspaceDestination As Integer

   CurrentStellarObjectSelection As Integer
   CurrentShipSelection As Integer

   AfterburnerOn As Boolean
   CloakOn As Boolean
   MoveUp As Boolean
   MoveDown As Boolean
   MoveLeft As Boolean
   MoveRight As Boolean
   Align As Boolean

   Shield As Integer
   Hull As Integer
   FuelLeft As Single
   Cloak As Integer
   Battery As Integer

   maMod As Single
   maArg As Single

   maBearing As Single
   Spin As Single

'   CurrentSecondary As integer

' Flags
'   Docked As Boolean
   Died As Boolean
   InHyperspace As Boolean
   InitialHyperspaceCruiseDistanceLeft As Long
   HyperspaceCruiseDistanceLeft As Long

' What the AI currently has in mind of doing
   ObjectiveType As String
   ObjectiveIndex As Integer

' Fighter hot/restrain [default = restrain]
   FightersHot As Boolean

   Guns As String
End Type

Type tShipType
' Class
   ClassName As String
   SpeciesSpecific As Integer
   GovernmentSpecific As Integer
   ShipImage As Integer

' Energy
   MaxShield As Integer
   MaxHull As Integer
   MaxFuel As Integer
   MaxCloak As Integer
   MaxBattery As Integer
   ShieldRechargeValue As Integer
   ShieldRechargePowerCost As Integer
   CloakRechargeValue As Integer
   CloakDrainRate As Integer
   CloakRechargePowerCost As Integer
   BatteryReGenerateValue As Integer
   BatteryReGenerateFuelCost As Integer

' Physics
   SpinFrictionRatio As Currency
   Acceleration As Currency
   AfterburnerAcceleration As Currency
   SpinAcceleration As Currency
   FrictionRatio As Single
   Size As Integer
   Mass As Single
   HyperspaceCruiseSpeed As Integer
   HyperspaceAccel As Single

' Inventory
   DefaultGunTypes As String

' Design
   GunPositionsX As String
   GunPositionsY As String

' Other
    Proximity As Integer
End Type
