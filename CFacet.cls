VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Describes a facet of the polyhedron."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=====================================================
'Represents a facet of the polyhedron
'by "inheriting" a CPolygon object
'=====================================================

Option Explicit

'local variable(s) to hold property value(s)
Dim mvarPolygon As CPolygon 'local copy
Dim mColor As Long
Dim mOrientation As Integer

Public Vertices As Collection 'array of vertex indices in a global vertex array
Public Edges As Collection 'array of edge indices; each edge begins in corresponding vertex from Vertices (see above)
Public Index As Long 'the index of Me in global Facet array
Public Parent As CFacets 'handle to a parent collection

Public Unwrapped As Boolean

'did we change the order of vertices in Vertices since first creation???
Public OrientationWasChanged As Boolean

Public Property Get VertexArray() As CVertices
Set VertexArray = mvarPolygon.VertexArray
End Property

Public Property Set VertexArray(ByVal vData As CVertices)
Set mvarPolygon.VertexArray = vData
mvarPolygon.CalculateCenterOfGravity
End Property

Public Property Set Polygon(ByVal vData As CPolygon)
Set mvarPolygon = vData
Set Vertices = mvarPolygon.Vertices
mvarPolygon.CalculateCenterOfGravity
End Property
Public Property Get Polygon() As CPolygon
Attribute Polygon.VB_Description = "Describes the polygon of a facet."
Attribute Polygon.VB_UserMemId = 0
Set Polygon = mvarPolygon
End Property

Private Sub Class_Initialize()
Set mvarPolygon = New CPolygon 'creating CPolygon
Set Vertices = mvarPolygon.Vertices 'linking to the Vertices property of a child polygon
Set Edges = New Collection
Set mvarPolygon.Parent = Me
OrientationWasChanged = False 'haven't changed orientation yet...heterosexual by birth...
Orientation = 0

mColor = GetRandomColor
End Sub

Private Sub Class_Terminate()
Set mvarPolygon.VertexArray = Nothing 'killing CPolygon.VertexArray
Set mvarPolygon = Nothing 'killing CPolygon
Set Edges = Nothing
End Sub

Public Property Get Consistent() As Boolean
' is everything OK with the vertices: does each index correspond
' to a single point in global vertex array?
' Moreover: does the vertex meet only once in the facet's vertex array?
'=====================================================
Dim Z As Long, Q As Long, W As Long, Z1 As Long, VerticesCount As Long
Dim V1 As Long, V2 As Long, V3 As Long, V4 As Long

VerticesCount = Vertices.Count
If VerticesCount < 3 Then
    AuxErrorString = Replace(GetString(ResMsg_FaceHasLessThan3Vertices), "%1", Me.Index)
    Consistent = False
    Exit Property
End If

For Z = 1 To VerticesCount
    Q = Vertices(Z)
    
    If Q < 1 Or Q > VertexArray.Count Then
        AuxErrorString = Replace(GetString(ResMsg_UnexistingVertex), "%1", Q)
        Consistent = False
        Exit Property
    End If
    
    If Z < VerticesCount Then
        For W = Z + 1 To VerticesCount
            If Vertices(W) = Q Then
                AuxErrorString = Replace(Replace(GetString(ResMsg_VertexBelongsTo3SidesOfAFace), "%1", Q), "%2", Me.Index)
                Consistent = False
                Exit Property
            End If
        Next
    End If
Next

If VerticesCount > 3 Then
    For Z = 1 To VerticesCount - 2
        If Z = 1 Then Z1 = VerticesCount - 1 Else Z1 = VerticesCount
        For Q = Z + 2 To Z1
            W = Q + 1
            If W = VerticesCount + 1 Then W = 1
            If SegmentsIntersect(Me.Parent.VertexArray(Vertices(Z)).DirectingVector, Me.Parent.VertexArray(Vertices(Z + 1)).DirectingVector, Me.Parent.VertexArray(Vertices(Q)).DirectingVector, Me.Parent.VertexArray(Vertices(W)).DirectingVector) Then
                AuxErrorString = ""
                Consistent = False
                Exit Property
            End If
        Next
    Next
End If

Consistent = True
End Property

Public Property Get Planar() As Boolean
Planar = mvarPolygon.Planar 'does the face lie within some plane?
'doesn't take into account small deviations (up to Epsilon=0.0001) from the main plane
AuxErrorString = Replace(AuxErrorString, "%1", Me.Index)
End Property

'Public Sub Planarize()
'mvarPolygon.Planarize 'force projection of all vertices to a single plane;
''not recommended because it step-by-step destroys topological structure
''of a parent polyhedron
'mvarPolygon.CalculateCenterOfGravity
'End Sub

Public Property Get Orientation() As Integer
Orientation = mOrientation
End Property

Public Property Let Orientation(ByVal vNewValue As Integer)
mOrientation = vNewValue
End Property

Public Property Get Oriented() As Boolean
Oriented = mOrientation <> 0
End Property

Public Sub ChangeOrientation()
OrientationWasChanged = Not OrientationWasChanged
mvarPolygon.ChangeOrientation
End Sub

Public Function HasEdge(Vertex1 As Long, Vertex2 As Long) As Integer
'HasEdge - looks for a specified edge in Me
'Vertex1 and Vertex2 are the indices of points in global array CVertices
'
'Function returns 0, if there is no such edge in currect face,
'                         1, if there is such edge and it is properly oriented,
'                        -1, if there is such edge and it is incorrectly oriented
Dim V1 As Long, V2 As Long

V1 = FindVertex(Vertex1) 'local reference to global point Vertex1 - exists?
V2 = FindVertex(Vertex2) 'local reference to global point Vertex1 - exists?
If V1 = 0 Or V2 = 0 Then HasEdge = 0: Exit Function ' if not then exit

If V2 = NextVertex(V1) Then HasEdge = -1: Exit Function
If V2 = PreviousVertex(V1) Then HasEdge = 1: Exit Function
End Function

Public Function FindVertex(Vertex1 As Long) As Long
'Finds a local vertex reference corresponding to global index Vertex1
'Returns 0 if finds nothing
Dim Z As Long
For Z = 1 To Vertices.Count
    If Vertices(Z) = Vertex1 Then FindVertex = Z: Exit Function
Next
FindVertex = 0
End Function

'Returns the next vertex reference from sequence
Public Function NextVertex(Vertex1 As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromLocal) As Long
Dim M As Long

If Semantics = logGlobalFromGlobal Or Semantics = logLocalFromGlobal Then
    M = FindVertex(Vertex1)
    If M = 0 Then NextVertex = 0: Exit Function
Else
    M = Vertex1
End If

If M < Vertices.Count Then
    M = M + 1
Else
    M = 1
End If

If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
    NextVertex = Vertices(M)
Else
    NextVertex = M
End If
End Function

'Returns the previous vertex reference from sequence
Public Function PreviousVertex(Vertex1 As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromLocal) As Long
Dim M As Long

If Semantics = logGlobalFromGlobal Or Semantics = logLocalFromGlobal Then
    M = FindVertex(Vertex1)
    If M = 0 Then PreviousVertex = 0: Exit Function
Else
    M = Vertex1
End If

If M > 1 Then
    M = M - 1
Else
    M = Vertices.Count
End If

If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
    PreviousVertex = Vertices(M)
Else
    PreviousVertex = M
End If
End Function

'Returns the previous edge reference from sequence
Public Function PreviousEdge(ByVal N As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromLocal) As Long
Dim M As Long

If Semantics = logGlobalFromGlobal Or Semantics = logLocalFromGlobal Then
    M = FindEdge(N, logLocalFromGlobal)
    If M = 0 Then PreviousEdge = 0: Exit Function
Else
    M = N
End If

If M > 1 Then
    M = M - 1
Else
    M = Edges.Count
End If

If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
    PreviousEdge = Edges(M)
Else
    PreviousEdge = M
End If
End Function

'Returns the next LOCAL edge reference from sequence
Public Function NextEdge(ByVal N As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromLocal) As Long
Dim M As Long

If Semantics = logGlobalFromGlobal Or Semantics = logLocalFromGlobal Then
    M = FindEdge(N)
    If M = 0 Then NextEdge = 0: Exit Function
Else
    M = N
End If

If M < Edges.Count Then
    M = M + 1
Else
    M = 1
End If

If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
    NextEdge = Edges(M)
Else
    NextEdge = M
End If
End Function

'returns an edge reference corresponding to edge index N
Public Function FindEdge(ByVal N As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromGlobal) As Long
Dim V, Z As Long

For Z = 1 To Edges.Count
    If Semantics = logGlobalFromLocal Or Semantics = logLocalFromLocal Then
        If Z = N Then
            If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
                FindEdge = Edges(Z)
            Else
                FindEdge = Z
            End If
            Exit Function
        End If
    Else
        If Edges(Z) = N Then
            If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then
                FindEdge = Edges(Z)
            Else
                FindEdge = Z
            End If
            Exit Function
        End If
    End If
Next

FindEdge = 0
End Function

Public Function GetEdgeByStartPoint(ByVal N As Long, Optional ByVal Semantics As LocalOrGlobal = logLocalFromGlobal) As Long
If Semantics = logGlobalFromGlobal Or Semantics = logLocalFromGlobal Then N = FindVertex(N)
If Semantics = logGlobalFromGlobal Or Semantics = logGlobalFromLocal Then N = Edges(N)
GetEdgeByStartPoint = N
End Function

Public Function GetAdjacentFace(Vertex1 As Long, Optional ErrorResult As OrientationResultType, Optional ByVal ShouldOrient As Boolean = True, Optional ByVal ShouldRaiseErrors As Boolean = True) As CFacet
'Vertex1 specifies a starting point of the edge of current face (Me)
'Vertex1 is LOCAL reference
'Assuming Me.Oriented = True
Dim F As CFacet, Vertex2 As Long, NumOfFoundFaces As Long

NumOfFoundFaces = 0
Vertex2 = NextVertex(Vertex1) 'the second point of the edge

For Each F In Me.Parent
    If Not F Is Me Then
        Select Case F.HasEdge(Vertices(Vertex1), Vertices(Vertex2))
        Case -1
            'found; same orientation; needs correction
            NumOfFoundFaces = NumOfFoundFaces + 1
            If NumOfFoundFaces = 1 Then
                If ShouldOrient Then F.ChangeOrientation  'perform orientation fitting
                Set GetAdjacentFace = F
            End If
        Case 0
            'not found
        Case 1
            'found; distinct orientation; OK
            NumOfFoundFaces = NumOfFoundFaces + 1
            If NumOfFoundFaces = 1 Then Set GetAdjacentFace = F
        End Select
    End If
Next

If NumOfFoundFaces = 1 Then
    'Everything is OK
    ErrorResult = oreOK
ElseIf NumOfFoundFaces = 0 Then
    'Me is a terminal face; polyhedron is not closed; this is bad
    Set GetAdjacentFace = Nothing
    AuxErrorString = GetString(ResMsg_PolyhedronNotClosed)
    AuxErrorString = Replace(AuxErrorString, "%1", Vertices(Vertex1))
    AuxErrorString = Replace(AuxErrorString, "%2", Vertices(Vertex2))
    AuxErrorString = Replace(AuxErrorString, "%3", Me.Index)
    ErrorResult = oreNotClosed
Else
    'more than two faces contain this edge; this is bad
    Set GetAdjacentFace = Nothing
    AuxErrorString = GetString(ResMsg_EdgeBelongsToMoreThan2Faces)
    AuxErrorString = Replace(AuxErrorString, "%1", Vertices(Vertex1))
    AuxErrorString = Replace(AuxErrorString, "%2", Vertices(Vertex2))
    AuxErrorString = Replace(AuxErrorString, "%3", NumOfFoundFaces + 1)
    ErrorResult = oreTooManyFaces
End If

End Function

Public Function Orient(ByVal nOrientation As Integer) As OrientationResultType
'recursively orient the face by reordering the vertices
Dim Z As Long, E As OrientationResultType
Dim F As CFacet

Me.Orientation = nOrientation

For Z = 1 To Vertices.Count
    E = oreOK
    
    Set F = GetAdjacentFace(Z, E)
    If F Is Nothing Then
        Orient = E
        Exit Function
    End If
    
    If Not F.Oriented Then
        E = F.Orient(-Me.Orientation)
    End If
Next

Orient = oreOK
End Function

Public Property Get PyramidVolume() As Double
'calculates the volume of the pyramid formed by this face and point (0 0 0)
PyramidVolume = mvarPolygon.PyramidVolume
End Property

Public Property Get Area() As Double
'calculates the area of the face
Area = mvarPolygon.Area
End Property

Public Property Get Convex() As Boolean
'checks whether the face polygon is convex
Convex = mvarPolygon.Convex
End Property

Public Function Normal() As CVector
Set Normal = mvarPolygon.Normal
End Function

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let Color(ByVal vNewValue As Long)
mColor = vNewValue
End Property

Public Function GetRandomColor() As Long
Dim M As Double, LB As Double, UB As Double

Const m_Brightness = 0
Const m_Pastel = 255

M = Random(m_Brightness, 255)

LB = M - m_Pastel
If LB < m_Brightness Then LB = m_Brightness
UB = M + m_Pastel
If UB > 255 Then UB = 255

GetRandomColor = RGB(Random(LB, UB), Random(LB, UB), Random(LB, UB))
End Function

'============================================
'
'============================================

Public Function OutputScanningBranch(ByVal Stage As Double, ByVal UnwrappedFacesCount As Long) As Long
Dim Z As Long, F As CFacet, E As CEdge, ErrRes As OrientationResultType
Dim N As Double, A As Double, V As CVector, C As Long, W As CVector, P As CVector, P2 As CVector
Dim SmallStage As Double

Unwrapped = True
UnwrappedFacesCount = UnwrappedFacesCount + 1
C = (Me.Parent.Count)

glBegin GL_POLYGON
    glColor3ub Red(mColor), Green(mColor), Blue(mColor)
    For Z = 1 To Vertices.Count
        Set P = VertexArray(Vertices(Z)).DirectingVector
        glVertex3f P.X, P.Y, P.Z
        Set P = Nothing
    Next
glEnd

For Z = 1 To Vertices.Count
    Set P = VertexArray(Vertices(Z)).DirectingVector
    Set P2 = MultVectorByCurrentMatrix(P)
    
    N = P2.Norm
    If N > Me.Parent.Parent.MaxVertexNorm Then
        Me.Parent.Parent.MaxVertexNorm = N
    End If
    
    Set P = Nothing
    Set P2 = Nothing
Next

'glBegin GL_POLYGON
'    glColor3ub Red(mColor), Green(mColor), Blue(mColor)
'    For Z = 1 To Vertices.Count
'        Set P = VertexArray(Vertices(Z)).DirectingVector
'        Set P2 = MultVectorByCurrentMatrix(P)
'
'        N = P2.Norm
'        If N > Me.Parent.Parent.MaxVertexNorm Then
'            Me.Parent.Parent.MaxVertexNorm = N
'        End If
'
'        glVertex3f P.X, P.Y, P.Z
'
'        Set P = Nothing
'        Set P2 = Nothing
'    Next
'glEnd

Dim V1 As Long, V2 As Long

For Z = 1 To Vertices.Count
    Set F = GetAdjacentFace(Z, ErrRes, False, False)
    If Not F.Unwrapped Then
        glPushMatrix
            SmallStage = (Stage - (UnwrappedFacesCount - 1) / (C - 1)) * (C - 1)
            If SmallStage < 0 Then SmallStage = 0
            If SmallStage > 1 Then SmallStage = 1
            V1 = NextVertex(Z, logGlobalFromLocal)
            V2 = Vertices(Z)
            Set V = VertexArray.JoiningVector(V1, V2)
            A = Me.Parent.Parent.AngleBetweenTwoFaces(Me.Index, F.Index) * 180 / PI
            
            If SmallStage > 0 Then
                Set W = VertexArray(V1).DirectingVector
                glTranslatef W.X, W.Y, W.Z
                glRotatef A * SmallStage, V.X, V.Y, V.Z
                glTranslatef -W.X, -W.Y, -W.Z
                Set W = Nothing
            End If
            
            UnwrappedFacesCount = F.OutputScanningBranch(Stage, UnwrappedFacesCount)
        glPopMatrix
    End If
Next

OutputScanningBranch = UnwrappedFacesCount
End Function

'============================================
'
'============================================

Public Sub CreateVisualModel(Canvas As ctlPolyViewer)
Dim P() As Long
Dim Z As Long

ReDim P(1 To Vertices.Count)
For Z = 1 To Vertices.Count
    P(Z) = Vertices(Z)
Next
Canvas.AddPolygon P(), Not Me.Convex, Me.Color
End Sub

Public Function GetCenterOfGravity() As CVector
Set GetCenterOfGravity = mvarPolygon.GetCenterOfGravity
End Function

Public Sub Recalc()
mvarPolygon.CalculateCenterOfGravity
End Sub
