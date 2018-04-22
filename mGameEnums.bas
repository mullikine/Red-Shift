Attribute VB_Name = "mGameEnums"
'---------------------------------------------------------------------------------------
' Module    : mGameEnums
' DateTime  : 12/16/2004 21:11
' Author    : Shane Mulligan
' Purpose   : Enums for wallpapers, relations, combat ratings, stellar object kinds...
'---------------------------------------------------------------------------------------

Option Explicit

Enum eWallpaper
    tScreen1 = 0
    tScreen2 = 1
    tScreen3 = 2
    tScreen4 = 3
    tScreen5 = 4
End Enum

Public Enum eRelations
   Neutral = 0
   Hostile = 1
   Friendly = 2
   Member = 3
   Self = 4
   Forbiddon = 5
   Master = 6
End Enum

Public Enum eCombatRating
   NoAbility = 0
   LittleAbility = 1
   FairAbility = 200
   AverageAbility = 400
   GoodAbility = 800
   Competent = 1600
   VeryCompetent = 3200
   WorthyofNote = 6400
   Dangerous = 12800
   Deadly = 25600
   Frightening = 51200
End Enum

Public Enum eSOKind
   Planet = 0
   Moon = 1
   Station = 2
   Quasar = 3
End Enum

Public Enum eWeaponClass
   Energy = 0
   Projectile = 1
   Beam = 2
   Rocket = 3
End Enum

Public Enum eCareer
   Pirate = 0
   Trader = 1
   Millitary = 2
   Privateer = 3
   Escort = 4
End Enum

Public Enum ePropulsionMethod
   Antigravity = 0
   Catalyst = 1
   Plasma = 2
End Enum

Public Enum eAlign
   Top_Align = 0
   Bottom_Align = 1
   Left_Align = 2
   Right_Align = 3
   TopLeft_Align = 4
   TopRight_Align = 5
   BottomLeft_Align = 6
   BottomRight_Align = 7
   Center_Align = 8
End Enum

Enum eExplosionType
   CatalystEx = 0
   SmokeEx = 1
End Enum

Public Enum SoundName

   sndJumpArive = 0
   sndJumpLeave = 1
   sndHostileJumpArive = 2
   sndHostileJumpLeave = 3
   sndSelect = 4
   sndAlert = 5
   sndAfterburner = 6
   sndPing = 7
   sndProton = 8
   sndLaser = 9
   sndNeutron = 10
   sndQuasar = 11
   sndExplosion = 12
   sndPhaser = 13

End Enum

Public Enum eHyperspaceToReturn
   Not_Enough_Fuel = 0
   Not_Far_Enough_From_Center = 1
   Entering_Hyperspace = 2
   Already_Entering_Hyperspace = 3
End Enum

Function StringToeWeaponClass(ByVal sString As String) As eWeaponClass

   Select Case sString
   Case "Energy"
      StringToeWeaponClass = Energy
   Case "Projectile"
      StringToeWeaponClass = Projectile
   Case "Beam"
      StringToeWeaponClass = Beam
   Case "Rocket"
      StringToeWeaponClass = Rocket
   End Select

End Function

Function RelationsToString(ByVal pRelations As eRelations) As String

   Select Case pRelations
   Case 0
      RelationsToString = "Neutral"
   Case 1
      RelationsToString = "Hostile"
   Case 2
      RelationsToString = "Friendly"
   Case 3
      RelationsToString = "Fleet Member"
   Case 4
      RelationsToString = "Yourself"
   Case 5
      RelationsToString = "Forbiddon"
   Case 6
      RelationsToString = "Fleet Master"
   End Select

End Function
