Attribute VB_Name = "mTabSelect"
'---------------------------------------------------------------------------------------
' Module    : mTabSelect
' DateTime  : 4/28/2005 23:54
' Author    : Shane Mulligan
' Purpose   : Implements ship selection and mathmatical formulae...
'---------------------------------------------------------------------------------------

Option Explicit

Public SelectRelations As Integer ' alias: eRelations

Public TabsOn As Boolean

Sub Init()
Dim iShip As Integer

   For iShip = 0 To nShips
      With Ships(iShip)
         
         .CurrentStellarObjectSelection = -1
         .CurrentShipSelection = -1
         
      End With
   Next iShip
   
   TabsOn = True

End Sub

Sub DoDraw()

   If TabsOn Then Draw

End Sub

Sub Draw()
   
   With Ships(You.Ship)
      
      DrawStellarObjectTabs .CurrentStellarObjectSelection
      DrawShipTabs .CurrentShipSelection
      
   End With

End Sub

Private Sub DrawFan(ByVal x As Integer, ByVal y As Integer, ByVal Distance As Single, ByVal Angle As Currency, ByVal Bearing As Currency, ByVal Colour As Long)
Dim TempVerts(3) As TLVERTEX
   
   EnableBlendColour
   
   Bearing = InvertY(Bearing)
   
   If Distance > 200 Then Distance = 200
   
   TempVerts(0).x = x
   TempVerts(0).y = y
   TempVerts(0).rhw = 1
   TempVerts(0).tu = 0
   TempVerts(0).tv = 0
   TempVerts(0).color = Colour
   
   TempVerts(1).x = ZoomX(Int(PolToX(Distance, Bearing + Angle) + x))
   TempVerts(1).y = ZoomY(Int(PolToY(Distance, Bearing + Angle) + y))
   TempVerts(1).tu = 0
   TempVerts(1).tv = 0
   TempVerts(1).rhw = 1
   TempVerts(1).color = &H0
   
   TempVerts(2).x = ZoomX(Int(PolToX(Distance, Bearing - Angle) + x))
   TempVerts(2).y = ZoomY(Int(PolToY(Distance, Bearing - Angle) + y))
   TempVerts(2).tu = 0
   TempVerts(2).tv = 0
   TempVerts(2).rhw = 1
   TempVerts(2).color = &H0

   TempVerts(3).x = x
   TempVerts(3).y = y
   TempVerts(3).tu = 0
   TempVerts(3).tv = 0
   TempVerts(3).rhw = 1
   TempVerts(3).color = Colour
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVerts(0), Len(TempVerts(0))

End Sub

Private Sub DrawShipTabs(ByVal TheShip As Integer)
Dim Distance As Single
Dim TempPoint As Point
Dim RefGun As Integer

   EnableBlendColour
   
   With Ships(You.Ship)
      
      If .CurrentShipSelection = -1 Then
         ' Code for selection display
         '...
      ElseIf Ships(.CurrentShipSelection).System <> .System Then
         ' Code for selection display
         '...
      Else
         ' Code for target aids
         For RefGun = 0 To UBound(Split(.Guns, ","))
            
            'If .Gun(RefGun).Firing Then
               ' Valid gun, ammo left
               If Split(.Guns, ",")(RefGun) <> -1 Then
                  If Not Guns(Split(.Guns, ",")(RefGun)).GunType = -1 Then
                     If Not (GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).FireRate = 0 Or Guns(Split(.Guns, ",")(RefGun)).AmmoRemaining = 0) Then
                        TempPoint = RotatePoint(Split(ShipTypes(.ShipType).GunPositionsX, ",")(RefGun), Split(ShipTypes(.ShipType).GunPositionsY, ",")(RefGun), .maBearing)
                        DrawVector .x + ZoomMod(TempPoint.x) + SpaceOffset.x, -(.y + ZoomMod(TempPoint.y)) + SpaceOffset.y, ZoomMod(ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).ProjectileType).InitialVelocity * 5), InvertY(Guns(Split(.Guns, ",")(RefGun)).Bearing), ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(RefGun)).GunType).ProjectileType).Colour, &H0
                     End If
                  End If
               End If
            'End If
            
         Next RefGun
            
         ' Code for drawfan
         Distance = DistanceFromShip(You.Ship, .CurrentShipSelection)
         DrawFan StatusbarDims.x / 2, ScreenDims.Height / 2, Distance, Atan(ShipTypes(Ships(TheShip).ShipType).Size / 2, Distance), CCur(CartToArg(Ships(TheShip).x - .x, Ships(TheShip).y - .y)), &HFF00FF00
         
         ' Code for selection box
         If ZoomMod(Distance) < 1024 Then
            DrawBox Ships(TheShip).x, Ships(TheShip).y, ShipTypes(Ships(TheShip).ShipType).Size, RelationColour(ShipRelations(You.Ship, TheShip))
         End If
         
         ' Code for selection display
         '...
      End If
      
   End With

End Sub

Private Sub DrawBox(ByVal pX As Long, ByVal pY As Long, ByVal pSize As Integer, ByVal pColour As Long)
Dim TempVerts(15) As TLVERTEX
Const CornerRatio As Integer = 3

   pSize = pSize / 2
   
   pX = pX + SpaceOffset.x
   pY = -pY + SpaceOffset.y
   
   For i = 0 To 15
      With TempVerts(i)
         
         .color = pColour
         .rhw = 1
         .specular = 0
         .tu = 1
         .tv = 1
         
      End With
   Next i
   
   TempVerts(0).x = pX - pSize
   TempVerts(0).y = pY - pSize
   TempVerts(1).x = pX - pSize / CornerRatio
   TempVerts(1).y = pY - pSize
   
   TempVerts(2).x = pX - pSize
   TempVerts(2).y = pY - pSize
   TempVerts(3).x = pX - pSize
   TempVerts(3).y = pY - pSize / CornerRatio
   
   
   TempVerts(4).x = pX + pSize
   TempVerts(4).y = pY - pSize
   TempVerts(5).x = pX + pSize / CornerRatio
   TempVerts(5).y = pY - pSize
   
   TempVerts(6).x = pX + pSize
   TempVerts(6).y = pY - pSize
   TempVerts(7).x = pX + pSize
   TempVerts(7).y = pY - pSize / CornerRatio
   
   
   TempVerts(8).x = pX + pSize
   TempVerts(8).y = pY + pSize
   TempVerts(9).x = pX + pSize / CornerRatio
   TempVerts(9).y = pY + pSize
   
   TempVerts(10).x = pX + pSize
   TempVerts(10).y = pY + pSize
   TempVerts(11).x = pX + pSize
   TempVerts(11).y = pY + pSize / CornerRatio
   
   
   TempVerts(12).x = pX - pSize
   TempVerts(12).y = pY + pSize
   TempVerts(13).x = pX - pSize / CornerRatio
   TempVerts(13).y = pY + pSize
   
   TempVerts(14).x = pX - pSize
   TempVerts(14).y = pY + pSize
   TempVerts(15).x = pX - pSize
   TempVerts(15).y = pY + pSize / CornerRatio
   

      
   
   For i = 0 To 15
      TempVerts(i).x = ZoomX(TempVerts(i).x)
      TempVerts(i).y = ZoomY(TempVerts(i).y)
   Next i
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 15, TempVerts(0), Len(TempVerts(0))

End Sub

Private Sub DrawStellarObjectTabs(ByVal TheStellarObject As Integer)
Dim Distance As Single

   With Ships(You.Ship)
      
      If .CurrentStellarObjectSelection <> -1 Then
         Distance = DistanceFromStellarObject(You.Ship, .CurrentStellarObjectSelection)
         DrawFan StatusbarDims.x / 2, ScreenDims.Height / 2, Distance, Atan(StellarObjects(TheStellarObject).Size / 2, Distance), CCur(CartToArg(StellarObjects(TheStellarObject).x - .x, StellarObjects(TheStellarObject).y - .y)), D3DColorRGBA(255, 0, 0, 50)
      End If
      
   End With

End Sub

Function DirectionToShip(ByVal FromShip As Integer, ByVal ToShip As Integer) As Double

   With Ships(FromShip)
      
      DirectionToShip = CartToArg(.x - Ships(ToShip).x, .y - Ships(ToShip).y)
   
   End With

End Function

Function DistanceBetween(ByRef Loc1 As Rect2, ByRef Loc2 As Rect2) As Double

   With Loc1
      
      DistanceBetween = Sqr((.x - Loc2.x) ^ 2 + (.y - Loc2.y) ^ 2)
   
   End With

End Function

Function DistBtwnShipPjtl(ByVal pShip As Integer, ByVal pProjectile As Integer) As Double

   With Projectiles(pProjectile)
   
      DistBtwnShipPjtl = BoundMin(Sqr((Ships(pShip).x - .x) ^ 2 + (Ships(pShip).y - .y) ^ 2) - ShipTypes(Ships(pShip).ShipType).Size / 2 - ProjectileTypes(.ProjectileType).Size, 0)
   
   End With

End Function

Function DistanceFromShip(ByVal FromShip As Integer, ByVal ToShip As Integer) As Double

   With Ships(FromShip)
      
      DistanceFromShip = BoundMin(Sqr((.x - Ships(ToShip).x) ^ 2 + (.y - Ships(ToShip).y) ^ 2) - ShipTypes(.ShipType).Size / 2 - ShipTypes(Ships(ToShip).ShipType).Size / 2, 0)
   
   End With

End Function

Function DistanceFromExToShip(ByVal FromEx As Integer, ByVal ToShip As Integer) As Double

   With Explosions(FromEx)
      
      DistanceFromExToShip = BoundMin(Sqr((.x - Ships(ToShip).x) ^ 2 + (.y - Ships(ToShip).y) ^ 2) - .Size / 2 - ShipTypes(Ships(ToShip).ShipType).Size / 2, 0)
   
   End With

End Function

Function DistanceFromSOToShip(ByVal FromSO As Integer, ByVal ToShip As Integer) As Double

   With StellarObjects(FromSO)
      
      DistanceFromSOToShip = BoundMin(Sqr((.x - Ships(ToShip).x) ^ 2 + (.y - Ships(ToShip).y) ^ 2) - .Size / 2 - ShipTypes(Ships(ToShip).ShipType).Size / 2, 0)
   
   End With

End Function

Function NextShip(ByVal pRefShip As Integer, Optional ByVal Relations As eRelations = eRelations.Neutral) As Integer

   With Ships(pRefShip)
      
      NextShip = .CurrentShipSelection + 1
      
      If NextShip > nShips Then
         NextShip = -1
         Exit Function
      End If
      
      While Ships(NextShip).System <> .System Or ShipRelations(pRefShip, NextShip) <> Relations Or Ships(NextShip).Died
         NextShip = NextShip + 1
         If NextShip > nShips Then
            NextShip = -1
            Exit Function
         End If
      Wend
      
   End With

End Function

Function NextStellarObject() As Integer

   With Ships(You.Ship)
      
      NextStellarObject = .CurrentStellarObjectSelection + 1
      If NextStellarObject > UBound(StellarObjects) Then NextStellarObject = -1: Exit Function
      While StellarObjects(NextStellarObject).System <> .System
         NextStellarObject = NextStellarObject + 1
         If NextStellarObject > UBound(StellarObjects) Then NextStellarObject = -1: Exit Function
      Wend
      
   End With

End Function

Function LandableSOExists(ByVal pSystem As Integer) As Boolean
Dim iSO As Integer
   
   For iSO = 0 To UBound(StellarObjects)
      If StellarObjects(iSO).System = pSystem And StellarObjects(iSO).Landable Then
         LandableSOExists = True
         Exit For
      End If
   Next iSO

End Function

Function SysExists(ByVal pGalaxy As Integer) As Boolean
Dim iSys As Integer
   
   For iSys = 0 To UBound(Systems)
      If Systems(iSys).Galaxy = pGalaxy Then
         SysExists = True
         Exit For
      End If
   Next iSys

End Function

Function RandomSO(ByVal pSystem As Integer) As Integer
      
   If LandableSOExists(pSystem) Then
      RandomSO = Int(Rnd * (UBound(StellarObjects) + 1))
      While StellarObjects(RandomSO).System <> pSystem Or Not StellarObjects(RandomSO).Landable
         RandomSO = Int(Rnd * (UBound(StellarObjects) + 1))
      Wend
   Else
      RandomSO = -1
   End If

End Function

Function RandomSys(ByVal pGalaxy As Integer) As Integer
      
   If SysExists(pGalaxy) Then
      RandomSys = Int(Rnd * (UBound(Systems) + 1))
      While Systems(RandomSys).Galaxy <> pGalaxy
         RandomSys = Int(Rnd * (UBound(Systems) + 1))
      Wend
   Else
      RandomSys = -1
   End If

End Function

Function RandomGxy() As Integer
      
   RandomGxy = Int(Rnd * (UBound(Galaxies) + 1))

End Function

Function ClosestShip(ByVal pRefShip As Integer, Optional ByVal Relations As eRelations = eRelations.Neutral) As Integer
Dim Distance As Long
Dim ShortestDistance As Long
Dim iShip As Integer

   ClosestShip = -1
   
   With Ships(pRefShip)
   
      For iShip = 0 To nShips
         If Ships(iShip).System = .System And ShipRelations(pRefShip, iShip) = Relations And Not Ships(iShip).Died And Not Ships(iShip).CloakOn Then
            Distance = DistanceFromShip(pRefShip, iShip)
            If (Distance < ShortestDistance Or ShortestDistance = 0) Then
               ShortestDistance = Distance
               ClosestShip = iShip
            End If
         End If
      Next iShip
   
   End With

End Function

Function ClosestShipFromSO(ByVal pStellarObject As Integer, ByVal Government As Integer) As Integer
Dim Distance As Long
Dim ShortestDistance As Long
Dim iShip As Integer

   ClosestShipFromSO = -1
   
   With StellarObjects(pStellarObject)
   
      For iShip = 0 To nShips
         If Ships(iShip).System = .System And Ships(iShip).Government = Government And Not Ships(iShip).Died And Not Ships(iShip).CloakOn Then
            Distance = DistanceFromSOToShip(pStellarObject, iShip)
            If (Distance < ShortestDistance Or ShortestDistance = 0) Then
               ShortestDistance = Distance
               ClosestShipFromSO = iShip
            End If
         End If
      Next iShip
   
   End With

End Function

Function DistanceFromStellarObject(ByVal FromShip As Integer, ByVal ToStellarObject As Integer) As Double

   With Ships(FromShip)
   
      DistanceFromStellarObject = BoundMin(Sqr((.x - StellarObjects(ToStellarObject).x) ^ 2 + (.y - StellarObjects(ToStellarObject).y) ^ 2) - ShipTypes(.ShipType).Size / 2 - StellarObjects(ToStellarObject).Size / 2, 0)
   
   End With

End Function

Function FuelToStellarObject(ByVal FromShip As Integer, ByVal ToStellarObject As Integer) As Double
Dim DistanceToDepart As Long
Dim DistanceFromSO As Long
Dim HyperspaceFuel As Long
Const JustToBeSafeConstant = 0.8 ' Ship isn't always at max speed

   With Ships(FromShip)
   
      If .System = StellarObjects(ToStellarObject).System Then
         ' Rational estimate
         DistanceFromSO = DistanceFromStellarObject(FromShip, ToStellarObject)
      Else
         ' Rational estimate
         DistanceToDepart = CartToMod(.x, .y)
         If DistanceToDepart < Systems(.System).HyperspaceDepartDistance Then
            DistanceToDepart = Systems(.System).HyperspaceDepartDistance - DistanceToDepart
         Else
            DistanceToDepart = 0
         End If
         
         ' Actual
         HyperspaceFuel = Sqr((Systems(StellarObjects(ToStellarObject).System).x - Systems(.System).x) ^ 2 + (Systems(StellarObjects(ToStellarObject).System).y - Systems(.System).x) ^ 2)
         
         ' Estimate
         DistanceFromSO = Systems(StellarObjects(ToStellarObject).System).HyperspaceArriveDistance
      End If
      
      FuelToStellarObject = HyperspaceFuel + (DistanceToDepart + DistanceFromSO) / (JustToBeSafeConstant * TopSpeed(.ShipType))
      
   End With

End Function

Function FuelToShip(ByVal FromShip As Integer, ByVal ToShip As Integer) As Double
Dim DistanceToDepart As Long
Dim DistFromShipInSys As Long
Dim HyperspaceFuel As Long
Const JustToBeSafeConstant = 0.8 ' Ship isn't always at max speed

   With Ships(FromShip)
   
      If .System = Ships(ToShip).System Then
         ' Rational estimate
         DistFromShipInSys = DistanceFromShip(FromShip, ToShip)
      Else
         ' Rational estimate
         DistanceToDepart = CartToMod(.x, .y)
         If DistanceToDepart < Systems(.System).HyperspaceDepartDistance Then
            DistanceToDepart = Systems(.System).HyperspaceDepartDistance - DistanceToDepart
         Else
            DistanceToDepart = 0
         End If
         
         ' Actual
         HyperspaceFuel = Sqr((Systems(Ships(ToShip).System).x - Systems(.System).x) ^ 2 + (Systems(Ships(ToShip).System).x - Systems(.System).x) ^ 2)
         
         ' Estimate
         DistFromShipInSys = Systems(Ships(ToShip).System).HyperspaceArriveDistance
      End If
      
      FuelToShip = HyperspaceFuel + (DistanceToDepart + DistFromShipInSys) / (JustToBeSafeConstant * TopSpeed(.ShipType))
      
   End With

End Function

Function ClosestStellarObjectByFuel(ByVal pRefShip As Integer, ByVal Government As Integer) As Integer
Dim FuelReq As Long
Dim LeastFuelReq As Long
Dim iSO As Integer

   ClosestStellarObjectByFuel = -1
   LeastFuelReq = -1
   
   For iSO = 0 To UBound(StellarObjects)
      If StellarObjects(iSO).Government = Government Then
         FuelReq = FuelToStellarObject(pRefShip, iSO)
         If (FuelReq < LeastFuelReq) Or LeastFuelReq = -1 Then
            LeastFuelReq = FuelReq
            ClosestStellarObjectByFuel = iSO
         End If
      End If
   Next iSO

End Function

Function ClosestShipByFuel(ByVal pRefShip As Integer, Optional ByVal pRelations As eRelations = eRelations.Neutral) As Integer
Dim FuelReq As Long
Dim LeastFuelReq As Long
Dim iShip As Integer

   ClosestShipByFuel = -1
   LeastFuelReq = -1
   
   For iShip = 0 To UBound(Ships)
      If ShipRelations(pRefShip, iShip) = pRelations Then
         FuelReq = FuelToShip(pRefShip, iShip)
         If (FuelReq < LeastFuelReq) Or LeastFuelReq = -1 Then
            LeastFuelReq = FuelReq
            ClosestShipByFuel = iShip
         End If
      End If
   Next iShip

End Function

Sub Reset(ByVal aShip As Integer)

   With Ships(aShip)
      
      .CurrentShipSelection = -1
      .CurrentStellarObjectSelection = -1
      
   End With

End Sub

 As Integer, ByVal ToStellarObject As Integer) As Double

   With Ships(FromShip)
   
      DistanceFromStellarObject = Sqr((.x - StellarObjects(ToStellarObject).x) ^ 2 + (.y - StellarObjects(ToStellarObject).y) ^ 2) - StellarObjects(ToStellarObject).Size / 2
      '  - ShipTypes(.ShipType).Size / 2 - StellarObjects(ToStellarObject).Size / 2
   End With

End Function

Function FuelToStellarObject(ByVal FromShip As Integer, ByVal ToStellarObject As Integer) As Double
Dim DistanceToDepart As Long
Dim DistanceFromSO As Long
Dim HyperspaceFuel As Long
Const JustToBeSafeConstant = 0.8 ' Ship isn't always at max speed

   With Ships(FromShip)
   
      If .system = StellarObjects(ToStellarObject).system Then
         ' Rational estimate
         DistanceFromSO = DistanceFromStellarObject(FromShip, ToStellarObject)
      Else
         ' Rational estimate
         DistanceToDepart = CartToMod(.x, .y)
         If DistanceToDepart < Systems(.system).HyperspaceDepartDistance Then
            DistanceToDepart = Systems(.system).HyperspaceDepartDistance - DistanceToDepart
         Else
            DistanceToDepart = 0
         End If
         
         ' Actual
         HyperspaceFuel = Sqr((Systems(StellarObjects(ToStellarObject).system).x - Systems(.system).x) ^ 2 + (Systems(StellarObjects(ToStellarObject).system).y - Systems(.system).x) ^ 2)
         
         ' Estimate
         DistanceFromSO = Systems(StellarObjects(ToStellarObject).system).HyperspaceArriveDistance
      End If
      
      FuelToStellarObject = HyperspaceFuel + (DistanceToDepart + DistanceFromSO) / (JustToBeSafeConstant * TopSpeed(.ShipType))
      
   End With

End Function

Function FuelToShip(ByVal FromShip As Integer, ByVal ToShip As Integer) As Double
Dim DistanceToDepart As Long
Dim DistFromShipInSys As Long
Dim HyperspaceFuel As Long
Const JustToBeSafeConstant = 0.8 ' Ship isn't always at max speed

   With Ships(FromShip)
   
      If .system = Ships(ToShip).system Then
         ' Rational estimate
         DistFromShipInSys = DistanceFromShipToShip(FromShip, ToShip)
      Else
         ' Rational estimate
         DistanceToDepart = CartToMod(.x, .y)
         If DistanceToDepart < Systems(.system).HyperspaceDepartDistance Then
            DistanceToDepart = Systems(.system).HyperspaceDepartDistance - DistanceToDepart
         Else
            DistanceToDepart = 0
         End If
         
         ' Actual
         HyperspaceFuel = Sqr((Systems(Ships(ToShip).system).x - Systems(.system).x) ^ 2 + (Systems(Ships(ToShip).system).x - Systems(.system).x) ^ 2)
         
         ' Estimate
         DistFromShipInSys = Systems(Ships(ToShip).system).HyperspaceArriveDistance
      End If
      
      FuelToShip = HyperspaceFuel + (DistanceToDepart + DistFromShipInSys) / (JustToBeSafeConstant * TopSpeed(.ShipType))
      
   End With

End Function

Function ClosestStellarObjectByFuel(ByVal pRefShip As Integer, Optional ByVal Government As Integer = -1, Optional LandableRequired As Boolean = True) As Integer
Dim FuelReq As Long
Dim LeastFuelReq As Long
Dim iSO As Integer

   ClosestStellarObjectByFuel = -1
   LeastFuelReq = -1
   
   For iSO = 0 To UBound(StellarObjects)
      If ((Government > -1 And StellarObjects(iSO).Government = Government) Or Government = -1) Then
         If ((StellarObjects(iSO).Landable And LandableRequired) Or Not LandableRequired) Then
            FuelReq = FuelToStellarObject(pRefShip, iSO)
            If (FuelReq < LeastFuelReq) Or LeastFuelReq = -1 Then
               LeastFuelReq = FuelReq
               ClosestStellarObjectByFuel = iSO
            End If
         End If
      End If
   Next iSO

End Function

Function ClosestSO(ByVal pRefShip As Integer, Optional ByVal toGovernment As Integer = -5, Optional toRelations As eRelations = -2, Optional LandableRequired As Boolean = True, Optional ByVal fromGovernment As Integer = -5) As Integer
Dim iClosestSO As Integer
Dim iClosestSODist As Long
Dim iNextSODist As Long
Dim iSO As Integer

' ToRelations can only be called when 2 government properties are filled

   iClosestSO = -1
   iClosestSODist = -1
   
   For iSO = 0 To UBound(StellarObjects)
      If (StellarObjects(iSO).system = Ships(pRefShip).system) Then
         ' -1 is no government
         ' If the planet is of the correct government ( -1 or > -1)
         If ((toGovernment > -2 And StellarObjects(iSO).Government = toGovernment) Or toGovernment = -5) Then
         
            ' if both governments are actual governments
            If StellarObjects(iSO).Government > -1 And fromGovernment > -1 Then
               ' consider relations
               If ((toRelations > -1 And GovRelations(fromGovernment, StellarObjects(iSO).Government) = toRelations) Or toRelations = -1) Then
                  If ((StellarObjects(iSO).Landable And LandableRequired) Or Not LandableRequired) Then
                     GoTo ConsiderPlanet
                  End If
               End If
            End If
         End If
      End If
NextiSO:
   Next iSO
   
   ClosestSO = iClosestSO
   
   Exit Function
ConsiderPlanet:

   iNextSODist = DistanceShipToSO(pRefShip, iSO)
   If iNextSODist < iClosestSODist Or iClosestSO = -1 Then
      iClosestSO = iSO
      iClosestSODist = iNextSODist
   End If
   
   GoTo NextiSO

End Function

Function ClosestShipByFuel(ByVal pRefShip As Integer, Optional ByVal pRelations As eRelations = eRelations.Neutral) As Integer
Dim FuelReq As Long
Dim LeastFuelReq As Long
Dim iShip As Integer

   ClosestShipByFuel = -1
   LeastFuelReq = -1
   
   For iShip = 0 To UBound(Ships)
      If ShipRelations(pRefShip, iShip) = pRelations Then
         FuelReq = FuelToShip(pRefShip, iShip)
         If (FuelReq < LeastFuelReq) Or LeastFuelReq = -1 Then
            LeastFuelReq = FuelReq
            ClosestShipByFuel = iShip
         End If
      End If
   Next iShip

End Function

Sub Reset(ByVal aShip As Integer)

   With Ships(aShip)
      
      .CurrentShipSelection = -1
      .CurrentStellarObjectSelection = -1
      
   End With

End Sub

