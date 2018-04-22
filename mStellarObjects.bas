Attribute VB_Name = "mStellarObjects"
'---------------------------------------------------------------------------------------
' Module    : mStellarObjects
' DateTime  : 12/16/2004 21:18
' Author    : Shane Mulligan
' Purpose   : Implements stellar objects
'---------------------------------------------------------------------------------------

Option Explicit

Public RefStellarObject As Integer


Sub DrawBodies()

   EnableBlendNormal
   
   For RefStellarObject = 0 To UBound(StellarObjects)
      With StellarObjects(RefStellarObject)
         
         If .System = Ships(You.Ship).System Then
            Select Case .Image
            Case -1
               ' draw nothing
            Case -2
               ' draw circle
               DrawCircle ZoomX(.x + SpaceOffset.x), ZoomY(-.y + SpaceOffset.y), SpaceZoom * (.Size / 2), White, Int(PI * (SpaceZoom * .Size) / 4)
            Case Else
               ' draw texture
               DrawTexture StellarObjectImages(.Image), srcRECTNorm, NewfRECT(.y - .Size / 2, .y + .Size / 2, .x - .Size / 2, .x + .Size / 2), .Bearing, True, White
            End Select
         End If
         
      End With
   Next RefStellarObject

End Sub


Sub DoPhysics()

   For RefStellarObject = 0 To UBound(StellarObjects)
      With StellarObjects(RefStellarObject)
      
          .Bearing = Mod360(.Bearing + .SpinSpeed)
      
      End With
   Next RefStellarObject

End Sub


Sub DrawRadars()
    
    For RefStellarObject = 0 To UBound(StellarObjects)
        With StellarObjects(RefStellarObject)
        
            If .System = Ships(You.Ship).System Then
               Select Case StellarObjects(RefStellarObject).Government
               Case -1
                  DrawToRadar mTextures.txrFlares(1), .x, .y, .Size, White
               Case Else
                  DrawToRadar mTextures.txrFlares(1), .x, .y, .Size, RelationColour(GovRelations(StellarObjects(RefStellarObject).Government, Ships(You.Ship).Government))
               End Select
            End If
            
        End With
    Next RefStellarObject

End Sub
lationColour(GovRelations(StellarObjects(RefStellarObject).Government, Ships(You.Ship).Government)), &H0, 0.5)
               End Select
            End If
            
        End With
    Next RefStellarObject

End Sub
