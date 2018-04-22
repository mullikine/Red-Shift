Attribute VB_Name = "mMath"
Option Explicit

Public Type Rect2
  x As Single
  y As Single
End Type

Public Type Pol2
  M As Single
  A As Single
End Type

Public Const PI As Double = 3.14159265358979
Public Const e As Double = 32.718281828459
'Public Const lInfinity = 2147483647, dInfinity = 1.79769313486231E+308, fInfinity = 3.402823E+38!

Function NewPoint(ByVal x As Single, ByVal y As Single) As Point

   With NewPoint
   
      .x = x
      .y = y
      
   End With

End Function

Function ReachP(ByVal V_i As Double, ByVal V_f As Double, ByVal r As Single) As Double

   ReachP = Log(V_f / V_i) / Log(r)

End Function

Function ReachD(ByVal V_i As Double, ByVal r As Single, ByVal P_i As Double, ByVal P_f As Double) As Double

   ReachD = -V_i * (r ^ P_i - r ^ P_f) / ln(r)

End Function

Function ln(ByVal Number As Double) As Double

   ln = Log(Number) / Log(e)

End Function

Function RotatePoint(ByVal x As Single, ByVal y As Single, ByVal RotAngle As Single) As Point
Dim Angle As Single
Dim Mag As Single

   Angle = CartToArg(x, y)
   Mag = CartToMod(x, y)
   RotatePoint.x = PolToX(Mag, Angle + RotAngle) 'RotateX(RotAngle, X, Y)
   RotatePoint.y = PolToY(Mag, Angle + RotAngle) 'RotateY(RotAngle, X, Y)

End Function

Function RotateX(Angle As Single, x As Single, y As Single) As Single

    'x = xCosq - ySinq
    RotateX = x * Cos(ToRadians(Angle)) - y * Sin(ToRadians(Angle))

End Function

Function RotateY(Angle As Single, x As Single, y As Single) As Single

    'y = xSinq + yCosq
    RotateY = x * Sin(ToRadians(Angle)) + y * Cos(ToRadians(Angle))

End Function

Sub MakeIdentity(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 4
        For j = 1 To 4
            If i = j Then
                M(i, j) = 1
            Else
                M(i, j) = 0
            End If
        Next j
    Next i
End Sub

Sub MatrixMatrixMult(r() As Single, A() As Single, B() As Single)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Value As Single

    For i = 1 To 4
        For j = 1 To 4
            Value = 0#
            For k = 1 To 4
                Value = Value + A(i, k) * B(k, j)
            Next k
            r(i, j) = Value
        Next j
    Next i
End Sub

Sub ShowMatrix(M() As Single)
Dim i As Integer
Dim j As Integer

    For i = 1 To 4
        For j = 1 To 4
            Debug.Print M(i, j);
        Next j
        Debug.Print
    Next i

End Sub

Sub VectorMatrixMult(ByRef Rpt() As Single, ByRef Ppt() As Single, ByRef A() As Single)
Dim i As Integer
Dim j As Integer
Dim Value As Single

    For i = 1 To 4
        Value = 0#
        For j = 1 To 4
            Value = Value + Ppt(j) * A(j, i)
        Next j
        Rpt(i) = Value
    Next i
    
    ' Renormalize the point.
    ' Note that value still holds Rpt(4).
    Rpt(1) = Rpt(1) / Value
    Rpt(2) = Rpt(2) / Value
    Rpt(3) = Rpt(3) / Value
    Rpt(4) = 1#
End Sub

Public Function Atan(ByVal x As Single, ByVal y As Single) As Single

   If x = 0 Then
      If y >= 0 Then Atan = 0 Else Atan = 180
   Else
      Atan = ToDegrees(Atn(y / x))
      If x > 0 Then
         If y >= 0 Then Atan = 90 - Atan
         If y < 0 Then Atan = -Atan + 90
      Else
         If y >= 0 Then Atan = -Atan + 270
         If y < 0 Then Atan = 90 - Atan + 180
      End If
   End If

End Function

Function Sign(ByVal Value As Single) As Single

   If Value <> 0 Then
      Sign = Value / Abs(Value)
   End If

End Function

Function InvertY(ByVal Arg As Single) As Single
Dim v As Rect2
   
   v = Pol2ToRect2(NewPol2(1, Arg))
   v.y = -v.y
   InvertY = Rect2ToPol2(v).A

End Function

Function Rect2Rect2Add(ByRef v As Rect2, ByRef w As Rect2) As Rect2

   With Rect2Rect2Add
   
      .x = v.x + w.x
      .y = v.y + w.y
      
   End With

End Function

Function Pol2Pol2Add(ByRef v As Pol2, ByRef w As Pol2) As Pol2

   Pol2Pol2Add = Rect2ToPol2(Rect2Rect2Add(Pol2ToRect2(v), Pol2ToRect2(w)))

End Function

Function PolarVectorVectorAddToMod(ByVal Mod1 As Single, ByVal Arg1 As Single, ByVal Mod2 As Single, ByVal Arg2 As Single) As Single

   PolarVectorVectorAddToMod = CartToMod(PolToX(Mod1, Arg1) + PolToX(Mod2, Arg2), PolToY(Mod1, Arg1) + PolToY(Mod2, Arg2))

End Function

Function PolarVectorVectorAddToArg(ByVal Mod1 As Single, ByVal Arg1 As Single, ByVal Mod2 As Single, ByVal Arg2 As Single) As Single

   PolarVectorVectorAddToArg = CartToArg(PolToX(Mod1, Arg1) + PolToX(Mod2, Arg2), PolToY(Mod1, Arg1) + PolToY(Mod2, Arg2))

End Function

Function PolToX(ByVal pMod As Single, ByVal Arg As Single) As Single

   PolToX = pMod * Sin(ToRadians(Arg))

End Function


Function PolToY(ByVal pMod As Single, ByVal Arg As Single) As Single

   PolToY = pMod * Cos(ToRadians(Arg))

End Function


Function CartToMod(ByVal x As Single, ByVal y As Single) As Single

   CartToMod = Sqr(x ^ 2 + y ^ 2)

End Function

Function CartToArg(ByVal x As Single, ByVal y As Single) As Single

   CartToArg = Atan(x, y)

End Function

Function Mod360(ByVal Degs As Single) As Single

   While Degs > 360: Degs = Degs - 360: Wend
   While Degs < 0: Degs = Degs + 360: Wend
   Mod360 = Degs

End Function

Function ToDegrees(ByVal sngAngle As Single) As Single

   ToDegrees = sngAngle * 180 / PI

End Function

Function ToRadians(ByVal sngAngle As Single) As Single

   ToRadians = sngAngle * PI / 180

End Function

Function BiggerNumber(ByVal Num1 As Single, ByVal Num2 As Single) As Single

   If Num1 >= Num2 Then BiggerNumber = Num1 Else BiggerNumber = Num2

End Function

Function SmallerNumber(ByVal Num1 As Single, ByVal Num2 As Single) As Single

   If Num1 <= Num2 Then SmallerNumber = Num1 Else SmallerNumber = Num2

End Function

Function DifferenceBetweenAngles(ByVal Angle1 As Single, ByVal Angle2 As Single) As Single
Dim v As Pol2
Dim w As Pol2

   v = NewPol2(1, Mod360(Angle1))
   w = NewPol2(1, Mod360(Angle2))
   
   DifferenceBetweenAngles = AngleBetweenRect2Rect2(Pol2ToRect2(v), Pol2ToRect2(w))
   
   If Int(Mod360(v.A - DifferenceBetweenAngles)) = Int(w.A) Then
      DifferenceBetweenAngles = -DifferenceBetweenAngles
   End If

End Function

Function PositivePart(ByVal Value As Integer) As Integer

   If Value > 0 Then PositivePart = Value

End Function

Function NegativePart(ByVal Value As Integer) As Integer

   If Value < 0 Then NegativePart = Value

End Function

Function AngleBetweenRect2Rect2(ByRef v As Rect2, ByRef w As Rect2) As Single

   AngleBetweenRect2Rect2 = ToDegrees(ACos(Bound(DotProductRect2Rect2(v, w) / (Rect2Mod(v) * Rect2Mod(w)), 1, -1)))

End Function

Function Bound(ByVal Value As Single, ByVal Max As Single, ByVal Min As Single) As Single

   Select Case Value
   Case Is > Max
      Bound = Max
   Case Is < Min
      Bound = Min
   Case Else
      Bound = Value
   End Select

End Function

Function BoundMax(ByVal Value As Single, ByVal Max As Single) As Single

   If Value > Max Then
      BoundMax = Max
   Else
      BoundMax = Value
   End If

End Function

Function BoundMin(ByVal Value As Single, ByVal Min As Single) As Single

   If Value < Min Then
      BoundMin = Min
   Else
      BoundMin = Value
   End If

End Function

Function NewRect2(ByVal x As Single, ByVal y As Single) As Rect2

   NewRect2.x = x
   NewRect2.y = y

End Function

Function NewPol2(ByVal M As Single, ByVal A As Single) As Pol2

   NewPol2.M = M
   NewPol2.A = A

End Function

Function Rect2ToPol2(ByRef v As Rect2) As Pol2

   Rect2ToPol2.M = CartToMod(v.x, v.y)
   Rect2ToPol2.A = CartToArg(v.x, v.y)

End Function

Function Pol2ToRect2(ByRef v As Pol2) As Rect2

   Pol2ToRect2.x = PolToX(v.M, v.A)
   Pol2ToRect2.y = PolToY(v.M, v.A)

End Function

Function Rect2Mod(ByRef v As Rect2) As Double

   Rect2Mod = CartToMod(v.x, v.y)

End Function

Function Rect2Arg(ByRef v As Rect2) As Double

   Rect2Arg = CartToArg(v.x, v.y)

End Function

Function DotProductRect2Rect2(ByRef v1 As Rect2, ByRef v2 As Rect2) As Double

   DotProductRect2Rect2 = v1.x * v2.x + v1.y * v2.y

End Function

' arc sine
' error if value is outside the range [-1,1]

Function ASin(Value As Double) As Double
    If Abs(Value) <> 1 Then
        ASin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ASin = 1.5707963267949 * Sgn(Value)
    End If
End Function

' arc cosine
' error if NUMBER is outside the range [-1,1]

Function ACos(ByVal Number As Double) As Double
    If Abs(Number) <> 1 Then
        ACos = 1.5707963267949 - Atn(Number / Sqr(1 - Number * Number))
    ElseIf Number = -1 Then
        ACos = 3.14159265358979
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function

' arc cotangent
' error if NUMBER is zero

Function ACot(Value As Double) As Double
    ACot = Atn(1 / Value)
End Function

' arc secant
' error if value is inside the range [-1,1]

Function ASec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ASec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ASec = 3.14159265358979 * Sgn(Value)
    End If
End Function

' arc cosecant
' error if value is inside the range [-1,1]

Function ACsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function
arc cosecant
' error if value is inside the range [-1,1]

Function ACsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ACsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ACsc = 1.5707963267949 * Sgn(Value)
    End If
End Function
