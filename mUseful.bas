Attribute VB_Name = "mUseful"
Function Pick(ByRef arOptions() As Variant) As Variant

   If UBound(arOptions) > -1 Then
      Pick = arOptions(Int(Rnd * UBound(arOptions)))
   Else
      Pick = False
   End If

End Function

Function Rect2InXYWH(ByRef Point As Rect2, ByRef Quad As XYWH, Optional ByVal bTouching As Boolean = False) As Boolean

   If bTouching Then
      Rect2InXYWH = (Point.x >= Quad.x And Point.x <= (Quad.x + Quad.Width)) And (Point.y >= Quad.y And Point.y <= (Quad.y + Quad.Height))
   Else
      Rect2InXYWH = (Point.x > Quad.x And Point.x < (Quad.x + Quad.Width)) And (Point.y > Quad.y And Point.y < (Quad.y + Quad.Height))
   End If

End Function
