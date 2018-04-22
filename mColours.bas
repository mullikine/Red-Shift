Attribute VB_Name = "mColours"
'---------------------------------------------------------------------------------------
' Module    : mColours
' DateTime  : 12/16/2004 21:14
' Author    : Shane Mulligan
' Purpose   : Handles colour operations and gets colours from relations etc...
'---------------------------------------------------------------------------------------

Option Explicit

Public Const _
            White = &HFFFFFFFF, _
            LightRed = &HFFFF7777, _
            Red = &HFF992222, _
            LightYellow = &HFFFFFF77, _
            Yellow = &HFFFFFF22, _
            LightGreen = &HFF77FF77, _
            Green = &HFF22FF22, _
            LightBlue = &HFF7777FF, _
            Blue = &HFF222299, _
            Purple = &HFF992299, _
            Cyan = &HFF10FFFF, _
            Grey = &HFFA6A6A6, _
            NavyGrey = &HFF6A5A4A, _
            DarkGrey = &HFF404040

Public Type RGBA
   Red As Byte
   Green As Byte
   Blue As Byte
   alpha As Byte
   Reserved As Byte
End Type

Function RelationColour(ByVal pRelation As eRelations) As Long

   Select Case pRelation
   Case Self
      RelationColour = Blue
   Case Hostile
      RelationColour = Red
   Case Neutral
      RelationColour = Cyan
   Case Friendly
      RelationColour = Green
   Case Member
      RelationColour = Blue
   Case Forbiddon
      RelationColour = Yellow
   Case Master
      RelationColour = Purple
   End Select

End Function


Function NewColour(ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal alpha As Integer) As Long

   NewColour = D3DColorRGBA(Red, Green, Blue, alpha)

End Function


Function NewGrade(ByVal Strength As Integer, Optional alpha As Integer = 255) As Long

   NewGrade = D3DColorRGBA(Strength, Strength, Strength, alpha)

End Function

Function AddRGBA(ByRef RGBA1 As RGBA, ByRef RGBA2 As RGBA) As RGBA

   On Error Resume Next
   With AddRGBA
   
      .Red = RGBA1.Red + RGBA2.Red
      .Green = RGBA1.Green + RGBA2.Green
      .Blue = RGBA1.Blue + RGBA2.Blue
      .alpha = RGBA1.alpha + RGBA2.alpha
      .Reserved = RGBA1.Reserved + RGBA2.Reserved
   
   End With

End Function

Function SubRGBA(ByRef RGBA1 As RGBA, ByRef RGBA2 As RGBA) As RGBA

   On Error Resume Next
   With SubRGBA
   
      .Red = RGBA1.Red - RGBA2.Red
      .Green = RGBA1.Green - RGBA2.Green
      .Blue = RGBA1.Blue - RGBA2.Blue
      .alpha = RGBA1.alpha - RGBA2.alpha
      .Reserved = RGBA1.Reserved - RGBA2.Reserved
   
   End With

End Function

Function NewRGBA(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal alpha As Byte, Optional Reserved As Byte = 255) As RGBA

   With NewRGBA
   
      .Red = Red
      .Green = Green
      .Blue = Blue
      .alpha = alpha
      .Reserved = Reserved
      
   End With

End Function

Function LongToRGBA(ByVal Colour As Long, Optional ByVal Reserved As Byte = 0) As RGBA

   With LongToRGBA
   
      .Red = Colour Mod &H100
      Colour = Colour \ &H100
      .Green = Colour Mod &H100
      Colour = Colour \ &H100
      .Blue = Colour Mod &H100
      Colour = Colour \ &H100
      .alpha = Colour Mod &H100
      .Reserved = Reserved
      
   End With

End Function

Function RGBAToLong(ByRef Colour As RGBA) As Long

   With Colour
   
      RGBAToLong = D3DColorRGBA(.Red, .Green, .Blue, .alpha)
      
   End With

End Function

Function Blend(ByVal Colour1 As Long, ByVal Colour2 As Long, ByVal Ratio As Single) As Long
Dim Col(1) As RGBA

   Col(0) = LongToRGBA(Colour1)
   Col(1) = LongToRGBA(Colour2)
   Blend = RGBAToLong(NewRGBA(CByte(Col(0).Red * Ratio + Col(1).Red * (1 - Ratio)), CByte(Col(0).Green * Ratio + Col(1).Green * (1 - Ratio)), CByte(Col(0).Blue * Ratio + Col(1).Blue * (1 - Ratio)), CByte(Col(0).alpha * Ratio + Col(1).alpha * (1 - Ratio)), CByte(Col(0).Reserved * Ratio + Col(1).Reserved * (1 - Ratio))))

End Function
o + Col(1).alpha * (1 - Ratio)), CByte(Col(0).Reserved * Ratio + Col(1).Reserved * (1 - Ratio))))

End Function
