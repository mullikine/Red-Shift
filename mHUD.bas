Attribute VB_Name = "mHUD"
'---------------------------------------------------------------------------------------
' Module    : mHUD
' DateTime  : 12/16/2004 21:25
' Author    : Shane Mulligan
' Purpose   : Implements the HUD
'---------------------------------------------------------------------------------------

Option Explicit

Type tMessage
   message As String
   TimePlaced As Single
   TimeLast As Single
   Colour As Long
End Type

Dim Messages(3) As tMessage
Dim RefMessage As Integer


Sub DoDraw()

   If Not mStatusMenu.MenuUp Then
      DrawMisc
   End If
   
   DrawMessages

End Sub


Private Sub DrawMisc()

   With Ships(You.Ship)
      
      EnableBlendNormal
      
      If .CurrentStellarObject = -1 Then
         DrawText Systems(.System).Name, 1, 1, 12, 12, LightRed
      Else
         DrawText Systems(.System).Name & ", " & StellarObjects(.CurrentStellarObject).Name & " (" & Governments(StellarObjects(.CurrentStellarObject).Government).Name & ")", 1, 1, 12, 12, LightRed
      End If
      
      If Ships(You.Ship).FightersHot Then
         DrawText "Fighters hot", 1, 52, 12, 12, Red
      End If
      DrawText "Next tab relations: " & RelationsToString(SelectRelations), 1, 66, 12, 12, White
      DrawText "Combat rating: " & GetCombatRating(Ships(You.Ship).Kills) & ", Kill count: " & Ships(You.Ship).Kills, 1, 80, 12, 12, White
      DrawText "Max speed: " & Int(TopSpeed(.ShipType)), 1, 94, 12, 12, White
      EnableBlendColour
      DrawCourse You.Ship, 250, vbYellow, vbRed
      EnableBlendNormal
      
      
      DrawText "FPS: " & mMonitor.LastFPS & ",   Procs per frame: " & mGame.ProcsPerFrame & ",   ProcCount: " & mMonitor.ProcCount & ",   nProjectiles: " & nProjectiles & ",  You.Ship: " & You.Ship & ",  nShips: " & nShips & ", nShipRel: " & UBound(ShipRelations, 1), 1, 718, 8, 8, White
      
      ' Ammo remaining
      DrawText "Ammo ", 1, 152 + 50 * 6, 13, 13, White
      For i = 0 To UBound(Split(.Guns, ","))
         If Split(.Guns, ",")(i) <> -1 Then
            If Guns(Split(.Guns, ",")(i)).GunType <> -1 Then
               If Len(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ClassName) Then
                  EnableBlendOne
                  DrawXYWH ViewImages(0), NewXYWH(1, 152 + 50 * 6 + 13 * (i + 1), 12 * Len(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ClassName), 12), Blue, , True
                  DrawText GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ClassName & "(" & ProjectileTypes(GunTypes(Guns(Split(.Guns, ",")(i)).GunType).ProjectileType).ClassName & " × " & Guns(Split(.Guns, ",")(i)).AmmoRemaining & ")", 1, 152 + 50 * 6 + 13 * (i + 1), 12, 12, White
               End If
            End If
         End If
      Next i
      
   End With

End Sub

Sub DrawMessages()
Dim ShiftDown As Boolean

   ShiftDown = False
   For RefMessage = 0 To 3
      With Messages(RefMessage)
         
         If .TimePlaced + .TimeLast > Timer And .message <> vbNullString Then
            DrawText .message, 10, 768 - 14 * (RefMessage + 1), 12, 12, .Colour, 1
            If ShiftDown Then
               Messages(RefMessage - 1) = Messages(RefMessage)
               Messages(RefMessage).message = vbNullString
            Else
               ShiftDown = False
            End If
         Else
            .message = vbNullString
            ShiftDown = True
         End If
         
      End With
   Next RefMessage

End Sub

Sub DisplayMessage(ByVal message As String, Optional ByVal Colour As Long = White, Optional ByVal TimeLast As Single = 3)

   For RefMessage = 0 To 3
      With Messages(RefMessage)
      
         If .message = vbNullString Then
            .message = message
            .Colour = Colour
            .TimePlaced = Timer
            .TimeLast = TimeLast
            Exit For
         Else
            If RefMessage = 3 Then
               Messages(0).message = message
               Messages(0).Colour = Colour
               Messages(0).TimePlaced = Timer
               Messages(0).TimeLast = TimeLast
            End If
         End If
      
      End With
   Next RefMessage
   mSounds.Play sndAlert, 0

End Sub

Sub DrawCourse(ByVal pShip As Integer, ByVal MaxPredict As Single, Optional ByVal Colour1 As Long = Yellow, Optional ByVal Colour2 As Long = Red)
Dim Pos As Pol2
Dim Vel As Pol2
Dim TempVerts() As TLVERTEX
Dim iPrediction As Integer

   Pos = Rect2ToPol2(NewRect2(Ships(pShip).x, Ships(pShip).y))
   Vel = NewPol2(Ships(pShip).maMod, Ships(pShip).maArg)
   
   ReDim TempVerts(MaxPredict) As TLVERTEX
   
   For iPrediction = 0 To MaxPredict
   
      With TempVerts(iPrediction)
         
         .x = ZoomX(Pol2ToRect2(Pos).x + SpaceOffset.x)
         .y = ZoomY(-Pol2ToRect2(Pos).y + SpaceOffset.y)
         .color = Blend(Colour2, Colour1, iPrediction / MaxPredict)
         .rhw = 1
         .specular = 0
         .tu = 0
         .tv = 0
         
      End With
      
      Pos = Pol2Pol2Add(Pos, Vel)
      Vel.M = Vel.M * ShipTypes(Ships(pShip).ShipType).FrictionRatio ^ Vel.M
      Vel = Pol2Pol2Add(Vel, NetGravityAt(Ships(pShip).System, Pol2ToRect2(Pos)))
   
   Next iPrediction
   
   D3DDevice.SetTexture 0, ViewImages(0)
   D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, MaxPredict, TempVerts(0), Len(TempVerts(0))

End Sub
