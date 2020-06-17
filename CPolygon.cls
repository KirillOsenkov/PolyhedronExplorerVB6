VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Describes a flat simple polygon in three dimensional space."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Represents a polygon in a three-dimensional space
'that connects vertices of the parent polyhedron.
'Given by the enumeration of indices of vertices in the global vertex array;
'the order of enumeration specifies orientation of the face(t)
Option Explicit

Private colVertices As Collection 'created and killed INSIDE this class
Public m_CenterOfGravity As New CVector
Public VertexArray As CVertices 'created and killed OUTSIDE this class
Public Parent As CFacet

Private Sub Class_Initialize()
Set colVertices = New Collection 'Creating
End Sub

Private Sub Class_Terminate()
Set colVertices = Nothing 'Killing
End Sub

Public Property Get Planar() As Boolean
'=================================================
'check that the polygon lies in plane
'=================================================
Dim Z As Long
Dim V1 As CVector, V2 As CVector, V As CVector

If colVertices.Count > 3 Then
    Set V1 = VertexArray.JoiningVector(Vertices(1), Vertices(2))
    Set V2 = VertexArray.JoiningVector(Vertices(1), Vertices(3))
    
    For Z = 4 To colVertices.Count
        Set V = VertexArray.JoiningVector(Vertices(1), Vertices(Z))
        If Not CVectorsComplanar(V1, V2, V) Then
            AuxErrorString = GetString(ResMsg_ContainsNonPlanarFaces)
            AuxErrorString = Replace(AuxErrorString, "%2", colVertices(1))
            AuxErrorString = Replace(AuxErrorString, "%3", colVertices(2))
            AuxErrorString = Replace(AuxErrorString, "%4", colVertices(3))
            AuxErrorString = Replace(AuxErrorString, "%5", colVertices(Z))
            Planar = False
            Exit Property
        End If
    Next
    Planar = True
Else
    Planar = True
End If
'Dim Z As Long
'Dim V1 As Vector, V2 As Vector, V As Vector
'
'If colVertices.Count > 3 Then
'    V1 = JoiningVector(VertexArray, colVertices(1), colVertices(2))
'    V2 = JoiningVector(VertexArray, colVertices(1), colVertices(3))
'    For Z = 4 To colVertices.Count
'        V = JoiningVector(VertexArray, colVertices(1), colVertices(Z))
'        If Not VectorsComplanar(V1, V2, V) Then
'            AuxErrorString = GetString(ResMsg_ContainsNonPlanarFaces)
'            AuxErrorString = Replace(AuxErrorString, "%2", colVertices(1))
'            AuxErrorString = Replace(AuxErrorString, "%3", colVertices(2))
'            AuxErrorString = Replace(AuxErrorString, "%4", colVertices(3))
'            AuxErrorString = Replace(AuxErrorString, "%5", colVertices(Z))
'            Planar = False
'            Exit Property
'        End If
'    Next
'    Planar = True
'Else
'    Planar = True
'End If
End Property

'Public Sub Planarize()
''force projection of all vertices to a single plane;
''not recommended because it step-by-step destroys topological structure
''of a parent polyhedron
'Dim Z As Long
'Dim V1 As Vector, V2 As Vector, V As Vector, C As Vector
'
'C = Me.Parent.GetCenterOfGravity
'
'If colVertices.Count > 3 Then 'a triangle is always planar :-)
'    V1 = VectorDifference(DirectingVector(VertexArray, colVertices(1)), C)
'    V2 = VectorDifference(DirectingVector(VertexArray, colVertices(2)), C)
'    For Z = 3 To colVertices.Count
'        V = VectorDifference(DirectingVector(VertexArray, colVertices(Z)), C)
'        If Not VectorsComplanar(V1, V2, V) Then
'            V = ProjectionOnPlane(V, V1, V2)
'            VertexArray(colVertices(Z)).Move C.X + V.X, C.Y + V.Y, C.Z + V.Z
'        End If
'    Next
'End If
'
''If colVertices.Count > 3 Then 'a triangle is always planar :-)
''    V1 = JoiningVector(VertexArray, colVertices(1), colVertices(2))
''    V2 = JoiningVector(VertexArray, colVertices(1), colVertices(3))
''    For Z = 4 To colVertices.Count
''        V = JoiningVector(VertexArray, colVertices(1), colVertices(Z))
''        If Not VectorsComplanar(V1, V2, V) Then
''            V = ProjectionOnPlane(V, V1, V2)
''            VertexArray(colVertices(Z)).Move VertexArray(colVertices(1)).X + V.X, VertexArray(colVertices(1)).Y + V.Y, VertexArray(colVertices(1)).Z + V.Z
''        End If
''    Next
''End If
'End Sub

'Public Property Get Consistent() As Boolean
'Dim Z As Long, Q As Long, W As Long, VerticesCount As Long
'
'VerticesCount = colVertices.Count
'If VerticesCount < 3 Then Consistent = False: Exit Property
'
'For Z = 1 To VerticesCount
'    Q = colVertices(Z)
'    If Q < 1 Or Q > VertexArray.Count Then
'        AuxErrorString = Replace(GetString(ResMsg_UnexistingVertex), "%1", Q)
'        Consistent = False
'        Exit Property
'    End If
'    If Z < VerticesCount Then
'        For W = Z + 1 To VerticesCount
'            If colVertices(W) = Q Then
'                AuxErrorString = Replace(Replace(GetString(ResMsg_VertexBelongsTo3SidesOfAFace), "%1", Q), "%2", Me.Index)
'                Consistent = False
'                Exit Property
'            End If
'        Next
'    End If
'Next
'
'Consistent = True
'End Property

Public Sub ChangeOrientation()
'does what it says
'by changing the order in which vertices are listed
Dim C As Collection, Z As Long, Q As Long

Set C = New Collection
Q = colVertices.Count
For Z = 1 To Q
    C.Add colVertices(Q + 1 - Z)
Next

'Set Vertices = Nothing
Set Vertices = C
Set C = Nothing
End Sub

Public Property Get Vertices() As Collection
Set Vertices = colVertices
End Property

Public Property Set Vertices(vNewValue As Collection)
Set colVertices = vNewValue
Set Me.Parent.Vertices = vNewValue
End Property

Public Property Get PyramidVolume() As Double
'========================================================
' Calculates oriented pyramid volume formed by a polygon and the origin
'========================================================
Dim Z As Long, Q As Long
Dim Sum As Double
Dim tV1 As CVector
Dim tV2 As CVector
Dim C As CVector

Sum = 0
Q = colVertices.Count

Set tV1 = VertexArray(colVertices(1)).DirectingVector
Set tV2 = VertexArray(colVertices(2)).DirectingVector
Set C = Me.Parent.GetCenterOfGravity

For Z = 1 To Q
    Sum = Sum + CMixedProduct(tV1, tV2, C)
    Set tV1 = Nothing
    Set tV1 = tV2
    Set tV2 = Nothing
    If Z < Q Then Set tV2 = VertexArray(colVertices(IIf(Z < Q - 1, Z + 2, 1))).DirectingVector
Next

Set tV1 = Nothing
Set tV2 = Nothing

PyramidVolume = Sum / 6
End Property

Public Property Get AreaOriented() As Double
'========================================================
' Retrieves oriented area of the polygon
'========================================================
Dim Z As Long, Q As Long
Dim Sum As Double
Dim tV1 As CVector
Dim tV2 As CVector
Dim tV3 As CVector
Dim tV4 As CVector
Dim C As CVector

Sum = 0
Q = colVertices.Count

Set C = Me.GetCenterOfGravity
Set tV1 = VertexArray(colVertices(1)).DirectingVector
Set tV2 = VertexArray(colVertices(2)).DirectingVector
tV1.Subtract C
tV2.Subtract C
Set tV4 = Me.Normal

For Z = 1 To Q
    Set tV3 = CVectorProduct(tV1, tV2)
    
    Sum = Sum + tV3.Norm * Sgn(CScalarProduct(tV3, tV4))
    
    Set tV3 = Nothing
    Set tV1 = Nothing
    Set tV1 = tV2
    Set tV2 = Nothing
    If Z < Q Then
        Set tV2 = VertexArray(colVertices(IIf(Z < Q - 1, Z + 2, 1))).DirectingVector
        tV2.Subtract C
    End If
Next

Set tV1 = Nothing

AreaOriented = Sum / 2
End Property

Public Property Get Area() As Double
'======================================================
' I don't believe you don't know what this property returns
'======================================================
Area = Abs(AreaOriented)
End Property

Public Property Get Convex() As Boolean
'======================================================
' Checks whether the polygon in plane is convex
'======================================================
Dim Z As Long, Q As Long
Dim tV As CVector
Dim tV1 As CVector
Dim tV2 As CVector
Dim tV3 As CVector

Q = colVertices.Count
Set tV1 = VertexArray(colVertices(Q - 1)).DirectingVector
Set tV2 = VertexArray(colVertices(Q)).DirectingVector
Set tV3 = VertexArray(colVertices(1)).DirectingVector
tV3.Subtract tV2
tV2.Subtract tV1
Set tV = CVectorProduct(tV2, tV3)

For Z = 1 To Q - 1
    Set tV1 = tV2
    Set tV2 = tV3
    Set tV3 = VertexArray(colVertices(Z + 1)).DirectingVector
    tV3.Subtract VertexArray(colVertices(Z)).GetDirectingVector
    If CScalarProduct(tV, CVectorProduct(tV2, tV3)) < 0 Then
        Convex = False
        Exit Property
    End If
Next

Convex = True
End Property

Public Function Normal() As CVector
Dim Z As Long, Q As Long
Dim V1 As CVector
Dim V2 As CVector
Dim V3 As CVector

Set Normal = New CVector

Q = colVertices.Count
For Z = 1 To Q
    Set V1 = VertexArray(colVertices(Z)).DirectingVector
    
    If Z < Q Then
        Set V2 = VertexArray(colVertices(Z + 1)).DirectingVector
    Else
        Set V2 = VertexArray(colVertices(1)).DirectingVector
    End If
    
    If Z < Q - 1 Then
        Set V3 = VertexArray.Item(colVertices(Z + 2)).DirectingVector
    Else
        Set V3 = VertexArray.Item(colVertices(Z - Q + 2)).DirectingVector
    End If
    V3.Subtract V2
    V2.Subtract V1
    
    V2.Normalize
    V3.Normalize
    Normal.Add CVectorProduct(V2, V3)
    Set V1 = Nothing
    Set V2 = Nothing
    Set V3 = Nothing
Next
End Function

Public Sub CalculateCenterOfGravity()
Dim Z As Long, Q As Long

Q = Me.Vertices.Count
If Q = 0 Then Exit Sub

m_CenterOfGravity.Clear
For Z = 1 To Q
    m_CenterOfGravity.Add VertexArray(Vertices(Z)).DirectingVector
Next
m_CenterOfGravity.ScaleVector 1 / Q
End Sub

Public Function GetCenterOfGravity()
Set GetCenterOfGravity = m_CenterOfGravity
End Function
