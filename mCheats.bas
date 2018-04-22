Attribute VB_Name = "mCheats"
'---------------------------------------------------------------------------------------
' Module    : mCheats
' DateTime  : 12/16/2004 21:15
' Author    : Shane Mulligan
' Purpose   : In game cheats. Not unlike the keyboard routine.
'---------------------------------------------------------------------------------------

Option Explicit

Sub Process()
Dim iIndex As Integer
   
   If Keys(vbKeyInsert) And Keys(192) Then
      
      With Ships(You.Ship)
         
         Select Case True
         
         Case Keys(vbKey1)
            .FuelLeft = ShipTypes(Ships(You.Ship).ShipType).MaxFuel
            
         Case Keys(vbKey2)
            NewShip 2, -1, 1, 0, 0, 0, 0, 0, 1, Privateer, You.Ship, 500
            mStars.InitialPhysics
            
         Case Keys(vbKey3)
            .Cloak = ShipTypes(Ships(You.Ship).ShipType).MaxCloak
            
         Case Keys(vbKey4)
            For iIndex = 0 To nShips
               If Not Ships(iIndex).Died Then Ships(iIndex).Hull = Int(Ships(iIndex).Hull / 1.1)
            Next iIndex
            
         Case Keys(vbKey5)
            For iIndex = 0 To nShips
               If Not iIndex = You.Ship Then
                  Ships(iIndex).ObjectiveType = "ChSh"
                  Ships(iIndex).ObjectiveIndex = You.Ship
               End If
            Next iIndex
         
         Case Keys(vbKey6)
            For iIndex = 0 To nShips
               If ShipRelations(iIndex, You.Ship) = Friendly Or ShipRelations(iIndex, You.Ship) = Member Or ShipRelations(iIndex, You.Ship) = Master Or ShipRelations(iIndex, You.Ship) = Self Then
                  Ships(iIndex).FuelLeft = ShipTypes(Ships(iIndex).ShipType).MaxFuel
               End If
            Next iIndex
         
         Case Keys(vbKey7)
            For iIndex = 0 To nShips
               If Not iIndex = You.Ship Then
                  Ships(iIndex).ObjectiveType = "FlSh"
                  Ships(iIndex).ObjectiveIndex = You.Ship
               End If
            Next iIndex
         
         Case Keys(vbKey8)
            If Ships(You.Ship).CurrentShipSelection <> -1 Then
               Ships(Ships(You.Ship).CurrentShipSelection).OwnorShip = You.Ship
               FormatIndividualShipRelations Ships(You.Ship).CurrentShipSelection
            End If
            
         Case Else
            Exit Sub
            
         End Select
         
      End With
      
      Beep
   End If

End Sub
         
         End Select
         
      End With
      
      Beep
   End If

End Sub
