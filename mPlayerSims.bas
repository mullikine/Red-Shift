Attribute VB_Name = "mPlayerSims"
'---------------------------------------------------------------------------------------
' Module    : mPlayerSims
' DateTime  : 12/16/2004 21:21
' Author    : Shane Mulligan
' Purpose   : Implements the loading of games or players
'---------------------------------------------------------------------------------------

Option Explicit

Public PlayerSims() As tProfile
Public nPlayerSims As Integer

'Public SaveProfiles() As tProfile
'Public SavePlayerProfilesLen As Integer


Function LoadPlayerSims() As Boolean
' Temporary function

   nPlayerSims = -1
   For i = 0 To UBound(ShipTypes)
      AddPlayerSim "eg. " & ShipTypes(i).ClassName, i
   Next i

End Function

'Function LoadPlayerSims() As Boolean
'
'On Error GoTo resetdatafile
'
'   Erase PlayerSims
'   nPlayerSims = -1
'   Open App.Path & "\Data\data.dat" For Binary As #1
'      Get #1, 1, SavePlayerProfilesLen
'   Close #1
'   Open App.Path & "\Data\data.dat" For Binary As #1 Len = 2 + SavePlayerProfilesLen
'      Get #1, 2, SaveProfiles()
'      nPlayerSims = UBound(SaveProfiles)
'      For i = 0 To nPlayerSims
'         AddPlayerSim SaveProfiles(i).Name, tProfileTocSim(SaveProfiles(i)).Ship
'      Next
'   Close #1
'
'   LoadPlayerSims = True
'
'   Exit Function
   
'resetdatafile:
'
'   Close #1
'
'   OverwriteTextFile App.Path & "\Data\data.dat", vbNullString
'
'   SavePlayerSims
'
'End Function


Function SavePlayerSims() As Boolean
' Temporary function

   'nPlayerSims = 0: ReDim PlayerSims(nPlayerSims)

End Function

'Sub SavePlayerSims()
'
'   ReDim SaveProfiles(nPlayerSims)
'   For i = 0 To nPlayerSims
'      SaveProfiles(i) = cSimTotProfile(PlayerSims(i))
'   Next
'
'   Open App.Path & "\Data\data.dat" For Binary As #1
'      Put #1, 1, CInt(Len(SaveProfiles))
'      Put #1, 2, SaveProfiles()
'   Close #1
'
'End Sub


Function DeletePlayerSim(ByVal PlayerSimname As String) As Boolean

On Error GoTo ErrHandler
   
   For i = 0 To nPlayerSims
      If PlayerSims(i).Name = PlayerSimname Then GoTo NextStep
   Next i
   
   Exit Function
   
NextStep:

   PlayerSims(i) = PlayerSims(nPlayerSims)
   
   ReDim Preserve PlayerSims(nPlayerSims - 1): nPlayerSims = nPlayerSims - 1

   DeletePlayerSim = True: Exit Function
   
ErrHandler:

End Function

Sub AddPlayerSim(ByVal PlayerSimname As String, ByVal ShipType As Integer)

   ReDim Preserve PlayerSims(nPlayerSims + 1): nPlayerSims = nPlayerSims + 1
   
   With PlayerSims(nPlayerSims)

      .Name = PlayerSimname
      .Kills = 0
      .Credit = 0
      .ShipType = ShipType
      .System = 0
      .x = 0
      .y = 0

   End With

End Sub

Function PlayerSimByName(ByVal PlayerSimname As String) As Integer

   PlayerSimByName = -1
   For j = 0 To nPlayerSims
      If PlayerSims(j).Name = PlayerSimname Then
         PlayerSimByName = j
      End If
   Next j

End Function
)
      .x = x
      .y = y

   End With

End Sub

Function PlayerSimByName(ByVal PlayerSimname As String) As Integer

   PlayerSimByName = -1
   For j = 0 To nPlayerSims
      If PlayerSims(j).Name = PlayerSimname Then
         PlayerSimByName = j
      End If
   Next j

End Function
