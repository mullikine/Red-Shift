Attribute VB_Name = "mGuns"
Option Explicit

Public Guns() As tGun
Public nGuns As Integer

Sub Init()

   nGuns = -1

End Sub

Function NewGun(ByVal GunType As Integer, Optional ByVal AmmoRemaining As Integer = -2, Optional ByVal pRefGunOverwrite As Integer = -1) As Integer
   
   If GunType = -1 Then
      NewGun = -1
      Exit Function
   End If
   
   If pRefGunOverwrite = -1 Then
      nGuns = nGuns + 1
      ReDim Preserve Guns(nGuns) As tGun
      pRefGunOverwrite = nGuns
   End If
   
   With Guns(pRefGunOverwrite)
   
      .GunType = GunType
      If AmmoRemaining = -2 Then
         .AmmoRemaining = GunTypes(GunType).MaxAmmo
      Else
         .AmmoRemaining = AmmoRemaining
      End If

   End With
   
   ' Return pRefGunOverwrite
   NewGun = pRefGunOverwrite

End Function
    GoTo GotRef
   End If
   
   For iChooser = 0 To nGuns
      If Guns(iChooser).Exists = False Then GoTo GotRef
   Next iChooser
   If iChooser = nGuns Then MsgBox "iChooser = nguns"
NewGunSlot:
   nGuns = nGuns + 1
   ReDim Preserve Guns(nGuns) As tGun
GotRef:
   NewGun = iChooser
   
   With Guns(NewGun)
   
      .GunType = GunType
      If AmmoRemaining = -2 Then
         .AmmoRemaining = GunTypes(GunType).MaxAmmo
      Else
         .AmmoRemaining = AmmoRemaining
      End If
      .Exists = True

   End With

End Function
