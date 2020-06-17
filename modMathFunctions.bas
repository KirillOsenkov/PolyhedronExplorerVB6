Attribute VB_Name = "modMathFunctions"
Option Explicit

Public Const PI As Double = 3.14159265358979
Public Const PI2 = 6.28318530717958
Public Const PIDiv2 As Double = 1.5707963267949
Public Const E As Double = 2.71828182845905
Public Const Ln2 As Double = 0.693147180559945
Public Const Ln10 As Double = 2.30258509299405
Public Const Infinity As Double = 10000000
Public Const Epsilon As Double = 0.0001
Public Const Sqr2 As Double = 1.4142135623731
Public Const Sqr3 As Double = 1.73205080756888
Public Const Sqr5 As Double = 2.23606797749979
Public Const GoldenSection As Double = 0.618033988749895
Public Const GoldenRatio As Double = 1.61803398874989
Public Const DegreeSign As String = "°"

Public Const ToRadians = PI / 180
Public Const ToDegrees = 180 / PI
Public Const ToRad = PI / 180
Public Const ToDeg = 180 / PI
Public Const Rad = PI / 180
Public Const Deg = 180 / PI

'======================================================

Public Function Factorial(ByVal Num As Long) As Double
Dim Product As Double, Z As Long
If Num < 1 Or Num > 100 Then Factorial = 1: Exit Function
Product = 1
For Z = 1 To Num: Product = Product * Z: Next
Factorial = Product
End Function

Public Function Combin(ByVal longN As Long, ByVal longK As Long) As Double
If longN < 1 Or longN > 100 Or longK < 0 Or longK > longN Then Exit Function
Combin = Factorial(longN) / (Factorial(longK) * Factorial(longN - longK))
End Function

'Trigonometric functions
Public Function Ctg(ByVal X As Double) As Double
Ctg = Cos(X) / Sin(X)
End Function

Public Function Sec(ByVal X As Double) As Double
Sec = 1 / Cos(X)
End Function

Public Function Cosec(ByVal X As Double) As Double
Cosec = 1 / Sin(X)
End Function

'Inverse trigonometric functions
Public Function Arcsin(ByVal X As Double) As Double
If X < -1 Or X > 1 Then Exit Function
If Abs(X) = 1 Then
    Arcsin = PI / 2 * Sgn(X)
Else
    Arcsin = Atn(X / Sqr(1 - X * X))
End If
End Function

Public Function Arccos(ByVal X As Double) As Double
If X < -1 Or X > 1 Then Exit Function
If X = 0 Then
    Arccos = PI / 2
Else
    Arccos = PI / 2 - Arcsin(X)
End If
End Function

Public Function Arcctg(ByVal X As Double) As Double
Arcctg = PI / 2 - Atn(X)
End Function

Public Function ArcSec(ByVal X As Double) As Double
ArcSec = Atn(X / Sqr(1 - X * X)) + (Sgn(X) - 1) * PI / 2
End Function

Public Function ArcCsc(ByVal X As Double) As Double
ArcCsc = Atn(1 / Sqr(1 - X * X)) + (Sgn(X) - 1) * PI / 2
End Function

'Hyperbolic functions
Public Function SinH(ByVal X As Double) As Double
SinH = (Exp(X) - Exp(-X)) / 2
End Function

Public Function CosH(ByVal X As Double) As Double
CosH = (Exp(X) + Exp(-X)) / 2
End Function

Public Function TanH(ByVal X As Double) As Double
TanH = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

Public Function CotH(ByVal X As Double) As Double
CotH = Exp(-X) / (Exp(X) - Exp(-X)) * 2 + 1
End Function

Public Function SecH(ByVal X As Double) As Double
SecH = 2 / (Exp(X) + Exp(-X))
End Function

Public Function CscH(ByVal X As Double) As Double
CscH = 2 / (Exp(X) - Exp(-X))
End Function

'Area functions hyperbolic
Public Function ArcsinH(ByVal X As Double) As Double
ArcsinH = Log(X + Sqr(X * X + 1))
End Function

Public Function ArccosH(ByVal X As Double) As Double
If X >= 1 Then ArccosH = Log(X + Sqr(X * X - 1))
End Function

Public Function ArctanH(ByVal X As Double) As Double
If X <> 1 And Sgn(X + 1) = Sgn(1 - X) Then ArctanH = Log((1 + X) / (1 - X)) / 2
End Function

Public Function ArccotH(ByVal X As Double) As Double
If Sgn(X + 1) = Sgn(X - 1) Then ArccotH = Log((X + 1) / (X - 1)) / 2
End Function

Public Function ArcsecH(ByVal X As Double) As Double
If X <> 0 And Abs(X) <= 1 Then ArcsecH = Log((Sqr(1 - X * X) + 1) / X)
End Function

Public Function ArccscH(ByVal X As Double) As Double
If X <> 0 And Sgn(X) * Sqr(X * X + 1) > -1 Then ArccscH = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function
'/Areafunctions hyperbolic

Public Function Random(ByVal LB As Double, UB As Double) As Double
Random = Rnd * (UB - LB) + LB
End Function

Public Function Minimum(ParamArray Num() As Variant) As Variant
Dim M As Double, Z As Long
M = Infinity
For Z = 0 To UBound(Num)
    If Num(Z) < M Then M = Num(Z)
Next Z
Minimum = M
End Function

Public Function Maximum(ParamArray Num() As Variant) As Variant
Dim M As Double, Z As Long
M = -Infinity
For Z = 0 To UBound(Num)
    If Num(Z) > M Then M = Num(Z)
Next Z
Maximum = M
End Function

Public Function Infimum(ParamArray Num() As Variant) As Variant
Dim M As Double, Z As Long
M = 1073741824
For Z = 0 To UBound(Num)
    If Num(Z) < M Then M = Num(Z)
Next Z
Infimum = M
End Function

Public Function Supremum(ParamArray Num() As Variant) As Variant
Dim M As Double, Z As Long
M = -1073741824
For Z = 0 To UBound(Num)
    If Num(Z) > M Then M = Num(Z)
Next Z
Supremum = M
End Function

Public Function Distance(ByVal X1 As Double, ByVal Y1 As Double, ByVal Z1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal Z2 As Double) As Double
Distance = Sqr((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1) + (Z2 - Z1) * (Z2 - Z1))
End Function

Public Function Angle(ByVal X1 As Double, ByVal Y1 As Double, ByVal Z1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal Z2 As Double) As Double
Angle = Arccos((X1 * X2 + Y1 * Y2 + Z1 * Z2) / Sqr((X1 * X1 + Y1 * Y1 + Z1 * Z1) * (X2 * X2 + Y2 * Y2 + Z2 * Z2)))
End Function
