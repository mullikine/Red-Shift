Attribute VB_Name = "mRelations"
Option Explicit

Public GovRelations() As eRelations
Public ShipRelations() As eRelations

Sub FormatGovRelations(ByVal eRelation As eRelations)
Dim GovX As Integer, GovY As Integer
   
   ReDim GovRelations(UBound(Governments), UBound(Governments)) As eRelations

   For GovX = 0 To UBound(Governments)
      For GovY = 0 To UBound(Governments)
         If GovX = GovY Then
            GovRelations(GovX, GovY) = Neutral
         Else
            GovRelations(GovX, GovY) = eRelation
         End If
      Next GovY
   Next GovX

End Sub

Sub FormatShipRelations()
Dim ShipX As Integer, ShipY As Integer
   
   ReDim ShipRelations(iMaxShips, iMaxShips) As eRelations
   
   For ShipX = 0 To nShips
      For ShipY = 0 To nShips
         Select Case True
         Case ShipX = ShipY
            ShipRelations(ShipX, ShipY) = Self
         Case Ships(ShipX).OwnorShip = ShipY
            ShipRelations(ShipX, ShipY) = Master
            ShipRelations(ShipY, ShipX) = Member
         Case Ships(ShipX).OwnorShip <> -1 ' Indirect allies
            If Ships(ShipX).OwnorShip = Ships(ShipY).OwnorShip Then
               ShipRelations(ShipX, ShipY) = Friendly
               ShipRelations(ShipY, ShipX) = Friendly
            End If
            If Ships(Ships(ShipX).OwnorShip).OwnorShip = ShipY Then
               ShipRelations(ShipX, ShipY) = Friendly
               ShipRelations(ShipY, ShipX) = Friendly
            End If
            If Ships(ShipY).OwnorShip <> -1 Then
               If Ships(Ships(ShipX).OwnorShip).OwnorShip <> -1 Then
                  If Ships(Ships(ShipX).OwnorShip).OwnorShip = Ships(Ships(ShipY).OwnorShip).OwnorShip Then
                     ShipRelations(ShipX, ShipY) = Friendly
                     ShipRelations(ShipY, ShipX) = Friendly
                  End If
                  If Ships(Ships(Ships(ShipX).OwnorShip).OwnorShip).OwnorShip = Ships(Ships(ShipY).OwnorShip).OwnorShip Then
                     ShipRelations(ShipX, ShipY) = Friendly
                     ShipRelations(ShipY, ShipX) = Friendly
                  End If
               End If
            End If
         Case Else
            ShipRelations(ShipX, ShipY) = GovRelations(Ships(ShipX).Government, Ships(ShipY).Government)
            ShipRelations(ShipY, ShipX) = GovRelations(Ships(ShipY).Government, Ships(ShipX).Government)
         End Select
      Next ShipY
   Next ShipX

End Sub

Sub FormatIndividualShipRelations(ByVal pRefShip As Integer)
Dim ShipY As Integer

   For ShipY = 0 To nShips
      Select Case True
      Case pRefShip = ShipY
         ShipRelations(pRefShip, ShipY) = Self
      Case Ships(pRefShip).OwnorShip = ShipY
         ShipRelations(pRefShip, ShipY) = Master
         ShipRelations(ShipY, pRefShip) = Member
      Case Ships(pRefShip).OwnorShip <> -1 ' Indirect allies
         If Ships(pRefShip).OwnorShip = Ships(ShipY).OwnorShip Then
            ShipRelations(pRefShip, ShipY) = Friendly
            ShipRelations(ShipY, pRefShip) = Friendly
         End If
         If Ships(Ships(pRefShip).OwnorShip).OwnorShip = ShipY Then
            ShipRelations(pRefShip, ShipY) = Friendly
            ShipRelations(ShipY, pRefShip) = Friendly
         End If
         If Ships(ShipY).OwnorShip <> -1 Then
            If Ships(Ships(pRefShip).OwnorShip).OwnorShip <> -1 Then
               If Ships(Ships(pRefShip).OwnorShip).OwnorShip = Ships(Ships(ShipY).OwnorShip).OwnorShip Then
                  ShipRelations(pRefShip, ShipY) = Friendly
                  ShipRelations(ShipY, pRefShip) = Friendly
               End If
               If Ships(Ships(Ships(pRefShip).OwnorShip).OwnorShip).OwnorShip = Ships(Ships(ShipY).OwnorShip).OwnorShip Then
                  ShipRelations(pRefShip, ShipY) = Friendly
                  ShipRelations(ShipY, pRefShip) = Friendly
               End If
            End If
         End If
      Case Else
         ShipRelations(pRefShip, ShipY) = GovRelations(Ships(pRefShip).Government, Ships(ShipY).Government)
         ShipRelations(ShipY, pRefShip) = GovRelations(Ships(ShipY).Government, Ships(pRefShip).Government)
      End Select
   Next ShipY

End Sub
