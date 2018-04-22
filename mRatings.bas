Attribute VB_Name = "mRatings"
'---------------------------------------------------------------------------------------
' Module    : mRatings
' DateTime  : 12/16/2004 21:20
' Author    : Shane Mulligan
' Purpose   : Contains a function for determining a sim's combat rating
'---------------------------------------------------------------------------------------

Option Explicit


Function GetCombatRating(ByVal Kills As Long) As String

   Select Case Kills
   Case Is >= Frightening
      GetCombatRating = "Frightening"
   Case Is >= Deadly
      GetCombatRating = "Deadly"
   Case Is >= Dangerous
      GetCombatRating = "Dangerous"
   Case Is >= WorthyofNote
      GetCombatRating = "Worthy of Note"
   Case Is >= VeryCompetent
      GetCombatRating = "Very Competent"
   Case Is >= Competent
      GetCombatRating = "Competent"
   Case Is >= GoodAbility
      GetCombatRating = "Good Ability"
   Case Is >= AverageAbility
      GetCombatRating = "Average Ability"
   Case Is >= FairAbility
      GetCombatRating = "Fair Ability"
   Case Is >= LittleAbility
      GetCombatRating = "Little Ability"
   Case NoAbility
      GetCombatRating = "No Ability"
   End Select

End Function
