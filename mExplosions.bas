Attribute VB_Name = "mExplosions"
'---------------------------------------------------------------------------------------
' Module    : mExplosions
' DateTime  : 5/2/2005 21:02
' Author    : Shane Mulligan
' Purpose   : Handles explosions and smoke
'---------------------------------------------------------------------------------------

Option Explicit

Private Type ExplosionType
    Alive As Boolean
    InitialTimeLeft As Integer
    TimeLeft As Integer
    x As Single
    y As Single
    InitialSize As Single
    Size As Single
    DeltaSize As Single
    System As Integer
    ExplosionType As eExplosionType
    Volume As Single
    Colour As Long
End Type

Public Explosions() As ExplosionType

Public ExplsnImages(1) As Direct3DTexture8

Dim RefExplosion As Integer

Sub Draw()
EnableBlendOne

   For RefExplosion = 0 To UBound(Explosions)
      With Explosions(RefExplosion)
      
         If .Alive And .System = Ships(You.Ship).System Then
            DrawTexture ExplsnImages(.ExplosionType), NewfRECT(1, 0, 1, 0), NewfRECT(.y - .Size, .y + .Size, .x - .Size, .x + .Size), Int(Rnd * 360) + 1, True, .Colour
         End If
         
      End With
   Next RefExplosion
    
End Sub

Sub DoPhysics()

   For RefExplosion = 0 To UBound(Explosions)
      With Explosions(RefExplosion)
      
         If .Alive Then
            .Size = .Size + .DeltaSize
            .TimeLeft = .TimeLeft - 1
            .Alive = .TimeLeft >= 0 And .Size > 0
            .Colour = NewGrade((.TimeLeft / .InitialTimeLeft) * 255, 255)
         End If
       
      End With
   Next RefExplosion

End Sub

Sub MakeExplosion(ByVal x As Single, ByVal y As Single, ByVal System As Integer, ByVal StartSize As Single, ByVal DeltaSize As Single, ByVal time As Integer, ByVal ExplosionType As eExplosionType, ByVal Volume As Single, Optional ByVal Colour As Long = &HFFFFFFFF)
Dim Selection As Integer

   Selection = -1
   For RefExplosion = 0 To UBound(Explosions)
      If Explosions(RefExplosion).Alive = False Then
         Selection = RefExplosion
         Exit For
      End If
   Next RefExplosion
   
   If Selection = -1 Then Exit Sub ' Dont bother
   
   With Explosions(Selection)
      .System = System
      .x = x
      .y = y
      .InitialSize = StartSize
      .Size = StartSize
      .DeltaSize = DeltaSize
      .InitialTimeLeft = time
      .TimeLeft = time
      .ExplosionType = ExplosionType
      .Volume = Volume
      .Colour = Colour
      .Alive = True
   End With
   
   If ExplosionType = CatalystEx Then
      If (mSounds.MinVolume * (DistanceFromExToShip(RefExplosion, You.Ship) / mSounds.MaxRange)) / Volume > mSounds.MinVolume And System = Ships(You.Ship).System Then
         If Volume <> 0 Then
            mSounds.Play sndExplosion, (mSounds.MinVolume * (DistanceFromExToShip(RefExplosion, You.Ship) / mSounds.MaxRange)) / Volume, x - Ships(You.Ship).x
         End If
      End If
   End If

End Sub

Sub ChangeMax(ByVal nMax As Integer)

    ReDim Explosions(nMax - 1)

End Sub

Public Sub CleanUp()

    Erase ExplsnImages
    Erase Explosions

End Sub

nds.HearingDistance) / 2 + 0.5, 0
      End If
   End If

End Sub

Sub ChangeMax(ByVal nMax As Integer)

    ReDim Preserve Explosions(nMax - 1)

End Sub

Public Sub CleanUp()

    Erase ExplsnImages
    Erase Explosions

End Sub

