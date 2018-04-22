Attribute VB_Name = "mText"
'---------------------------------------------------------------------------------------
' Module    : mText
' DateTime  : 12/16/2004 21:17
' Author    : Shane Mulligan
' Purpose   : Handles font blitting and custom fonts
'---------------------------------------------------------------------------------------

Option Explicit

'> Font
Public fntTex(1) As Direct3DTexture8
Private VertChar(3) As TLVERTEX

Sub DrawText(ByVal strText As String, ByVal startX As Single, ByVal StartY As Single, ByVal Height As Integer, ByVal Width As Integer, Optional ByVal Colour As Long = vbWhite, Optional ByVal iFont As Integer = 1)
   
   Dim i As Integer ' Loop variable
   Dim CharX As Integer, CharY As Integer ' Grid coordinates for our character 0-15 and 0-7
   Dim Char As String ' The current Character in the string
   Dim LinearEntry As Integer 'Without going into 2D entries, just work it out if it were a line
   Const tu As Single = (1 / 16)
   Const tv As Single = (1 / 16)
   
   startX = startX - Width
   
   If Len(strText) = 0 Then Exit Sub ' If there is no text dont try to render it....
   
   ChangeTextureFactor Colour
   EnableBlendNormal
   
   For i = 1 To Len(strText) ' Loop through each character
    
      ' 1. Choose the Texture Coordinates
      'To do this we just need to isolate which entry in the texture we
      'need to use - the Vertex creation code sorts out the ACTUAL texture coordinates
      Char = Mid$(strText, i, 1) ' Get the current character
      LinearEntry = Asc(Char) - 32
      CharY = Int((LinearEntry) / 16)
      CharX = LinearEntry - CharY * 16
              
          
      ' 2. Generate the Vertices
      VertChar(0) = NewTLVERTEX(startX + (Width * i), StartY, 0, 1, Colour, 0, tu * CharX + 1 / 256, tv * CharY + 1 / 256)
      VertChar(1) = NewTLVERTEX(startX + (Width * i) + Width, StartY, 0, 1, Colour, 0, (tu * CharX) + tu + 1 / 256, tv * CharY + 1 / 256)
      VertChar(2) = NewTLVERTEX(startX + (Width * i), StartY + Height, 0, 1, Colour, 0, tu * CharX + 1 / 256, (tv * CharY) + tv + 1 / 256)
      VertChar(3) = NewTLVERTEX(startX + (Width * i) + Width, StartY + Height, 0, 1, Colour, 0, (tu * CharX) + tu + 1 / 256, (tv * CharY) + tv + 1 / 256)
      
      ' 3. Render the vertices
      D3DDevice.SetTexture 0, fntTex(iFont) ' Set the device to use our custom font as a texture
      D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertChar(0), Len(VertChar(0))
   Next i

End Sub

Function Italics(ByVal NonItalicString) As String

   For i = 1 To Len(NonItalicString)
      Italics = Italics & Chr(Asc(Mid(NonItalicString, i, 1)) + 128)
   Next i
   
End Function

Function UnItalics(ByVal ItalicString) As String

   For i = 1 To Len(ItalicString)
      UnItalics = UnItalics & Chr(Asc(Mid(ItalicString, i, 1)) - 128)
   Next i
   
End Function


Sub CleanUp()

    Erase fntTex

End Sub
