Attribute VB_Name = "modCVector"
Option Explicit

'==============================================
'==============================================
' CVector operations
'==============================================
'==============================================

Public Function CScalarProduct(V1 As CVector, V2 As CVector) As Double
CScalarProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function

Public Function CMixedProduct(V1 As CVector, V2 As CVector, V3 As CVector) As Double
CMixedProduct = V1.X * (V2.Y * V3.Z - V3.Y * V2.Z) - V1.Y * (V2.X * V3.Z - V3.X * V2.Z) + V1.Z * (V2.X * V3.Y - V3.X * V2.Y)
End Function

Public Function CVectorsComplanar(V1 As CVector, V2 As CVector, V3 As CVector) As Boolean
CVectorsComplanar = Abs(CMixedProduct(V1, V2, V3)) < Epsilon
End Function

Public Function CVectorProduct(V1 As CVector, V2 As CVector) As CVector
Set CVectorProduct = New CVector
CVectorProduct.X = V1.Y * V2.Z - V2.Y * V1.Z
CVectorProduct.Y = V2.X * V1.Z - V1.X * V2.Z
CVectorProduct.Z = V1.X * V2.Y - V2.X * V1.Y
End Function

Public Function AngleTwoCVectors(V1 As CVector, V2 As CVector) As Double
AngleTwoCVectors = Arccos(CScalarProduct(V1, V2) / (V1.Norm * V2.Norm))
End Function

Public Function DistanceTwoCVectors(V1 As CVector, V2 As CVector) As Double
DistanceTwoCVectors = Sqr((V1.X - V2.X) * (V1.X - V2.X) + (V1.Y - V2.Y) * (V1.Y - V2.Y) + (V1.Z - V2.Z) * (V1.Z - V2.Z))
End Function

Public Function CGrahamSchmidtOrthogonalize(V1 As CVector, V2 As CVector) As CVector
Set CGrahamSchmidtOrthogonalize = New CVector
CGrahamSchmidtOrthogonalize.InitWith V1
CGrahamSchmidtOrthogonalize.ScaleVector -CScalarProduct(V1, V2)
CGrahamSchmidtOrthogonalize.Add V2
End Function

Public Function CVectorDifference(V1 As CVector, V2 As CVector) As CVector
Set CVectorDifference = New CVector
CVectorDifference.X = V1.X - V2.X
CVectorDifference.Y = V1.Y - V2.Y
CVectorDifference.Z = V1.Z - V2.Z
End Function

Public Function CVectorLinearCombination(D1 As Double, V1 As CVector, D2 As Double, V2 As CVector) As CVector
Set CVectorLinearCombination = New CVector
CVectorLinearCombination.X = D1 * V1.X + D2 * V2.X
CVectorLinearCombination.Y = D1 * V1.Y + D2 * V2.Y
CVectorLinearCombination.Z = D1 * V1.Z + D2 * V2.Z
End Function

Public Function CNullVector() As CVector
Set CNullVector = New CVector
End Function

Public Function SegmentsIntersect(V1 As CVector, V2 As CVector, V3 As CVector, V4 As CVector) As Boolean
If Not CVectorsComplanar(CVectorDifference(V2, V1), CVectorDifference(V3, V1), CVectorDifference(V4, V1)) Then SegmentsIntersect = False: Exit Function
If CScalarProduct(CVectorProduct(CVectorDifference(V2, V1), CVectorDifference(V3, V1)), CVectorProduct(CVectorDifference(V2, V1), CVectorDifference(V3, V1))) > 0 Then SegmentsIntersect = False: Exit Function
If CScalarProduct(CVectorProduct(CVectorDifference(V4, V3), CVectorDifference(V1, V3)), CVectorProduct(CVectorDifference(V4, V3), CVectorDifference(V2, V3))) > 0 Then SegmentsIntersect = False: Exit Function
SegmentsIntersect = True
End Function
