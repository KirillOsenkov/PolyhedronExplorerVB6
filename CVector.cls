VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=====================================================
' Models a three-dimensional vector
'=====================================================

Public X As Double ' Vector components
Public Y As Double
Public Z As Double

Public Sub Add(V As CVector)
X = X + V.X
Y = Y + V.Y
Z = Z + V.Z
End Sub

Public Sub Subtract(V As CVector)
X = X - V.X
Y = Y - V.Y
Z = Z - V.Z
End Sub

Public Sub InitWith(V As CVector)
X = V.X
Y = V.Y
Z = V.Z
End Sub

Public Sub Normalize()
Dim D As Double

If X = 0 And Y = 0 And Z = 0 Then Exit Sub

D = 1 / Norm
X = X * D
Y = Y * D
Z = Z * D
End Sub

Public Sub ScaleVector(ByVal Scalar As Double)
X = X * Scalar
Y = Y * Scalar
Z = Z * Scalar
End Sub

Public Function NormalizedVector() As CVector
Dim D As Double
Set NormalizedVector = New CVector

If X = 0 And Y = 0 And Z = 0 Then
    NormalizedVector.InitWith Me
    Exit Function
End If

D = 1 / Norm
NormalizedVector.X = X * D
NormalizedVector.Y = Y * D
NormalizedVector.Z = Z * D
End Function

Public Sub Clear()
X = 0
Y = 0
Z = 0
End Sub

Public Function Norm() As Double
Norm = Sqr(X * X + Y * Y + Z * Z)
End Function
