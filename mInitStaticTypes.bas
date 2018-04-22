Attribute VB_Name = "mInitStaticTypes"
Option Explicit

Public Galaxies() As tGalaxy
Public Systems() As tSystem
Public StellarObjects() As tStellarObject
Public Governments() As tGovernment
Public Species() As tSpecies
Public Sims() As tProfile

Sub Init()

   InitGalaxies
   InitSystems
   InitStellarObjects
   InitGovernments
   InitSpecies

End Sub

Private Sub InitGalaxies()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Galaxies\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim Galaxies(aInput(1))
   
   For i = 0 To UBound(Galaxies)
      With Galaxies(i)
      
         Open App.Path & "\Data\Classes\Galaxies\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Name = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .x = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .y = Val(aInput(1))
         Close #1
      
      End With
   Next i

End Sub

Private Sub InitSystems()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Systems\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim Systems(aInput(1))
   
   For i = 0 To UBound(Systems)
      With Systems(i)
      
         Open App.Path & "\Data\Classes\Systems\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Name = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .x = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .y = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Galaxy = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HyperspaceArriveDistance = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .HyperspaceDepartDistance = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Government = Val(aInput(1))
         Close #1
      
      End With
   Next i
   
End Sub

Private Sub InitStellarObjects()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Stellar Objects\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim StellarObjects(aInput(1))
   
   For i = 0 To UBound(StellarObjects)
      With StellarObjects(i)
      
         Open App.Path & "\Data\Classes\Stellar Objects\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Name = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .Kind = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .System = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .x = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .y = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Size = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Image = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .SpinSpeed = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .GravitationalFieldStrength = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .GravGeometricRatio = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .MaxGravityAcceleration = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .DescriptionIndex = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Government = Val(aInput(1))
            aInput = Tokenize(ReadStr(1), " ")
            .Landable = CBool(aInput(1))
            
            ' non fs values
            .Bearing = Rnd * 360
         Close #1
      
      End With
   Next i

End Sub

Private Sub InitGovernments()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Governments\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim Governments(aInput(1))
   ReDim GovRelations(aInput(1), aInput(1))
   
   For i = 0 To UBound(Governments)
      With Governments(i)
      
         Open App.Path & "\Data\Classes\Governments\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Name = Replace(aInput(1), "~", " ")
            aInput = Tokenize(ReadStr(1), " ")
            .Species = Val(aInput(1))
            For j = 0 To UBound(Governments)
               aInput = Tokenize(ReadStr(1), " ")
               GovRelations(i, j) = aInput(1)
            Next j
         Close #1
      
      End With
   Next i

End Sub

Private Sub InitSpecies()
Dim aInput() As String

   Open App.Path & "\Data\Classes\Species\Specs.txt" For Input As #1
      aInput = Tokenize(ReadStr(1), " ")
   Close #1
   
   ReDim Species(aInput(1))
   
   For i = 0 To UBound(Species)
      With Species(i)
      
         Open App.Path & "\Data\Classes\Species\" & i & ".txt" For Input As #1
            aInput = Tokenize(ReadStr(1), " ")
            .Name = Replace(aInput(1), "~", " ")
         Close #1
      
      End With
   Next i

End Sub
