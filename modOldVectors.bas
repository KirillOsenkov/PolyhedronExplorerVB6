Attribute VB_Name = "modOldVectors"
Option Explicit

'Public Type Vector
'    X As Double
'    Y As Double
'    Z As Double
'End Type
'
'Public Type Matrix
'    V1 As Vector
'    V2 As Vector
'    V3 As Vector
'End Type

'==============================================
'==============================================
' Vector operations
'==============================================
'==============================================

'Public Function MakeVector(V1 As CVector) As Vector
'MakeVector.X = V1.X
'MakeVector.Y = V1.Y
'MakeVector.Z = V1.Z
'End Function

'Public Function Vector(V1 As CVector) As Vector
'Vector.X = V1.X
'Vector.Y = V1.Y
'Vector.Z = V1.Z
'End Function

'Public Function VectorsComplanar(V1 As Vector, V2 As Vector, V3 As Vector) As Boolean
'VectorsComplanar = Abs(MixedProduct(V1, V2, V3)) < Epsilon
'End Function

'Public Function MixedProduct(V1 As Vector, V2 As Vector, V3 As Vector) As Double
'MixedProduct = V1.X * (V2.Y * V3.Z - V3.Y * V2.Z) - V1.Y * (V2.X * V3.Z - V3.X * V2.Z) + V1.Z * (V2.X * V3.Y - V3.X * V2.Y)
'End Function
'
'
'Public Function ScalarProduct(V1 As Vector, V2 As Vector) As Double
'ScalarProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
'End Function
'
'
'Public Function VectorProduct(V1 As Vector, V2 As Vector) As Vector
'VectorProduct.X = V1.Y * V2.Z - V2.Y * V1.Z
'VectorProduct.Y = V2.X * V1.Z - V1.X * V2.Z
'VectorProduct.Z = V1.X * V2.Y - V2.X * V1.Y
'End Function

'Public Function DistanceTwoVectors(V1 As Vector, V2 As Vector) As Double
'DistanceTwoVectors = Sqr((V1.X - V2.X) ^ 2 + (V1.Y - V2.Y) ^ 2 + (V1.Z - V2.Z) ^ 2)
'End Function

'Public Function AngleTwoVectors(V1 As Vector, V2 As Vector) As Double
'AngleTwoVectors = Arccos(ScalarProduct(V1, V2) / (Norm(V1) * Norm(V2)))
'End Function

'Public Function ProjectionOnVector(V1 As Vector, V2 As Vector) As Vector
''ProjectionOnVector = ScalarProduct(V1, V2) / (Norm(V1) ^ 2)
'Dim D As Double
'D = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
'D = D / (V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
'ProjectionOnVector.X = V2.X * D
'ProjectionOnVector.Y = V2.Y * D
'ProjectionOnVector.Z = V2.Z * D
'End Function

'Public Function GrahamSchmidtOrthogonalize(V1 As Vector, V2 As Vector) As Vector
'GrahamSchmidtOrthogonalize = VectorSum(ScaleVector(V1, -ScalarProduct(V1, V2)), V2)
'End Function

'Public Function ProjectionOnPlane(V As Vector, V1 As Vector, V2 As Vector) As Vector
'Dim N As Vector, D As Double
'N = NormalToPlane(V1, V2)
'ProjectionOnPlane = VectorSum(V, ScaleVector(N, -ScalarProduct(N, V)))
'End Function

'Public Function DistanceFromPointToPlane(V As Vector, V1 As Vector, V2 As Vector) As Double
'Dim N As Vector, D As Double
'N = NormalToPlane(V1, V2)
'D = Abs(ScalarProduct(N, V))
'End Function

'Public Function NormalToPlane(V1 As Vector, V2 As Vector) As Vector
'NormalToPlane = Normalize(VectorProduct(V1, V2))
'End Function

'Public Function VectorSum(V1 As Vector, V2 As Vector) As Vector
'VectorSum.X = V1.X + V2.X
'VectorSum.Y = V1.Y + V2.Y
'VectorSum.Z = V1.Z + V2.Z
'End Function

'Public Function VectorDifference(V1 As Vector, V2 As Vector) As Vector
'VectorDifference.X = V1.X - V2.X
'VectorDifference.Y = V1.Y - V2.Y
'VectorDifference.Z = V1.Z - V2.Z
'End Function

'Public Function VectorLinearCombination(D1 As Double, V1 As Vector, D2 As Double, V2 As Vector) As Vector
'VectorLinearCombination.X = D1 * V1.X + D2 * V2.X
'VectorLinearCombination.Y = D1 * V1.Y + D2 * V2.Y
'VectorLinearCombination.Z = D1 * V1.Z + D2 * V2.Z
'End Function

'Public Function ScaleVector(V As Vector, ByVal Scalar As Double) As Vector
'ScaleVector.X = V.X * Scalar
'ScaleVector.Y = V.Y * Scalar
'ScaleVector.Z = V.Z * Scalar
'End Function
'
'Public Function Norm(V1 As Vector) As Double
'Norm = Sqr(V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
'End Function
'
'Public Function Normalize(V1 As Vector) As Vector
'Dim D As Double
'If V1.X = 0 And V1.Y = 0 And V1.Z = 0 Then
'    Normalize = V1
'    Exit Function
'End If
'D = 1 / Sqr(V1.X * V1.X + V1.Y * V1.Y + V1.Z * V1.Z)
'Normalize.X = V1.X * D
'Normalize.Y = V1.Y * D
'Normalize.Z = V1.Z * D
'End Function
'
'Public Function NullVector() As Vector
'End Function

'Public Function OrtX() As Vector
'OrtX.X = 1
'End Function
'
'Public Function OrtY() As Vector
'OrtY.Y = 1
'End Function
'
'Public Function OrtZ() As Vector
'OrtZ.Z = 1
'End Function

'Public Function GetVector(VertexArray As CVertices, ByVal Vertex1 As Long, ByVal Vertex2 As Long) As Vector
'With VertexArray
'    GetVector.X = .Item(Vertex2).X - .Item(Vertex1).X
'    GetVector.Y = .Item(Vertex2).Y - .Item(Vertex1).Y
'    GetVector.Z = .Item(Vertex2).Z - .Item(Vertex1).Z
'End With
'End Function

'Public Function DirectingVector(VertexArray As CVertices, ByVal Vertex1 As Long) As Vector
'With VertexArray.Item(Vertex1)
'    DirectingVector.X = .X
'    DirectingVector.Y = .Y
'    DirectingVector.Z = .Z
'End With
'End Function

'Public Function EdgeMiddlePoint(VertexArray As CVertices, ByVal V1 As Long, ByVal V2 As Long) As Vector
'EdgeMiddlePoint.X = (VertexArray(V1).X + VertexArray(V2).X) / 2
'EdgeMiddlePoint.Y = (VertexArray(V1).Y + VertexArray(V2).Y) / 2
'EdgeMiddlePoint.Z = (VertexArray(V1).Z + VertexArray(V2).Z) / 2
'End Function
'
'Public Function EdgeSubdivision(VertexArray As CVertices, ByVal V1 As Long, ByVal V2 As Long, ByVal Ratio As Double) As Vector
'EdgeSubdivision.X = VertexArray(V1).X + (VertexArray(V2).X - VertexArray(V1).X) * Ratio
'EdgeSubdivision.Y = VertexArray(V1).Y + (VertexArray(V2).Y - VertexArray(V1).Y) * Ratio
'EdgeSubdivision.Z = VertexArray(V1).Z + (VertexArray(V2).Z - VertexArray(V1).Z) * Ratio
'End Function

'Public Function MultiplyVectorByMatrix(V As Vector, M As Matrix) As Vector
'MultiplyVectorByMatrix.X = V.X * M.V1.X + V.Y * M.V2.X + V.Z * M.V3.X
'MultiplyVectorByMatrix.Y = V.X * M.V1.Y + V.Y * M.V2.Y + V.Z * M.V3.Y
'MultiplyVectorByMatrix.Z = V.X * M.V1.Z + V.Y * M.V2.Z + V.Z * M.V3.Z
'End Function
'
'Public Function MultiplyMatrixByVector(M As Matrix, V As Vector) As Vector
'MultiplyMatrixByVector.X = ScalarProduct(V, M.V1)
'MultiplyMatrixByVector.Y = ScalarProduct(V, M.V2)
'MultiplyMatrixByVector.Z = ScalarProduct(V, M.V3)
'End Function
'
'Public Function MultiplyMatrices(M As Matrix, N As Matrix) As Matrix
'MultiplyMatrices.V1.X = M.V1.X * N.V1.X + M.V1.Y * N.V2.X + M.V1.Z * N.V3.X
'MultiplyMatrices.V1.Y = M.V1.X * N.V1.Y + M.V1.Y * N.V2.Y + M.V1.Z * N.V3.Y
'MultiplyMatrices.V1.Z = M.V1.X * N.V1.Z + M.V1.Y * N.V2.Z + M.V1.Z * N.V3.Z
'MultiplyMatrices.V2.X = M.V2.X * N.V1.X + M.V2.Y * N.V2.X + M.V2.Z * N.V3.X
'MultiplyMatrices.V2.Y = M.V2.X * N.V1.Y + M.V2.Y * N.V2.Y + M.V2.Z * N.V3.Y
'MultiplyMatrices.V2.Z = M.V2.X * N.V1.Z + M.V2.Y * N.V2.Z + M.V2.Z * N.V3.Z
'MultiplyMatrices.V3.X = M.V3.X * N.V1.X + M.V3.Y * N.V2.X + M.V3.Z * N.V3.X
'MultiplyMatrices.V3.Y = M.V3.X * N.V1.Y + M.V3.Y * N.V2.Y + M.V3.Z * N.V3.Y
'MultiplyMatrices.V3.Z = M.V3.X * N.V1.Z + M.V3.Y * N.V2.Z + M.V3.Z * N.V3.Z
'End Function
'
'Public Function MultiplyMatrixByScalar(M As Matrix, A As Double) As Matrix
'MultiplyMatrixByScalar.V1.X = M.V1.X * A
'MultiplyMatrixByScalar.V1.Y = M.V1.Y * A
'MultiplyMatrixByScalar.V1.Z = M.V1.Z * A
'MultiplyMatrixByScalar.V2.X = M.V2.X * A
'MultiplyMatrixByScalar.V2.Y = M.V2.Y * A
'MultiplyMatrixByScalar.V2.Z = M.V2.Z * A
'MultiplyMatrixByScalar.V3.X = M.V3.X * A
'MultiplyMatrixByScalar.V3.Y = M.V3.Y * A
'MultiplyMatrixByScalar.V3.Z = M.V3.Z * A
'End Function
'
'Public Function Determinant(M As Matrix) As Double
'Determinant = M.V1.X * (M.V2.Y * M.V3.Z - M.V3.Y * M.V2.Z) - M.V1.Y * (M.V2.X * M.V3.Z - M.V3.X * M.V2.Z) + M.V1.Z * (M.V2.X * M.V3.Y - M.V3.X * M.V2.Y)
'End Function
'
'Public Function NormalizeMatrix(M As Matrix) As Matrix
'Dim D As Double
'D = Determinant(M)
'If D <> 0 Then NormalizeMatrix = MultiplyMatrixByScalar(M, 1 / D) Else NormalizeMatrix = M
'End Function
'
'Public Function Identity() As Matrix
'Identity.V1.X = 1
'Identity.V2.Y = 1
'Identity.V3.Z = 1
'End Function
'
'Public Function GetRotationMatrix() As Matrix
'Dim V1 As Vector, V2 As Vector, V3 As Vector
'Dim Z As Long, T As Long
'
'T = timeGetTime
'
'V1.X = 2
'V1.Y = 3
'V1.Z = 0.5
'V2.X = 1
'V2.Y = -3.7
'V2.Z = PI
'V3.X = E
'V3.Y = 4
'V3.Z = 5
'
'For Z = 1 To 1000
'    V3 = VectorProduct(V1, V2)
'Next
'
'MsgBox timeGetTime - T
'End Function

'Public Function CenterOfGravity(F As CFacet) As Vector
'Dim Z As Long
'Dim V As Vector, S As Vector
'
'For Z = 1 To F.Vertices.Count
'    V = DirectingVector(F.VertexArray, F.Vertices(Z))
'    S = VectorSum(V, S)
'Next
'CenterOfGravity = ScaleVector(S, 1 / F.Vertices.Count)
'End Function
'
'Public Function PolyhedronCenterOfGravity(VertexArray As CVertices) As Vector
'Dim V As Vector, Z As Long
'
'For Z = 1 To VertexArray.Count
'    V.X = V.X + VertexArray(Z).X
'    V.Y = V.Y + VertexArray(Z).Y
'    V.Z = V.Z + VertexArray(Z).Z
'Next
'V.X = V.X / VertexArray.Count
'V.Y = V.Y / VertexArray.Count
'V.Z = V.Z / VertexArray.Count
'PolyhedronCenterOfGravity = V
'End Function

'Public Sub TransformPolyhedron()
'
'End Sub

