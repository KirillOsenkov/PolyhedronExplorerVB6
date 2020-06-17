VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Specifies an edge of a polyhedron."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'A class to represent an edge of a polyhedron
'Contains indices of two points that determine the edge;
'starting point is the first
'Contains also a method to determine the length of the edge

Option Explicit

'two points of the edge
Public StartPoint As Long
Public EndPoint As Long

'two adjacent faces
Public LeftFace As Long 'the adjacent face that is to the left of the edge when looking from outside the polyhedron
Public RightFace As Long 'the same thing for right face

Public Index As Long
Public VertexArray As CVertices 'handle to the global vertex array of the parent polyhedron

Private Sub Class_Initialize()
StartPoint = 0
EndPoint = 0
End Sub

Private Sub Class_Terminate()
Set VertexArray = Nothing
End Sub

Public Property Get Length() As Double
'self-explanatory; no comments
Length = Distance(VertexArray(StartPoint).X, VertexArray(StartPoint).Y, VertexArray(StartPoint).Z, VertexArray(EndPoint).X, VertexArray(EndPoint).Y, VertexArray(EndPoint).Z)
End Property

'====================================================================
' output 3D graphic data to a Canvas object (usually ctlPolyViewer)
'====================================================================
Public Sub CreateVisualModel(Canvas As ctlPolyViewer)
Canvas.AddLine StartPoint, EndPoint
End Sub

Public Function EdgeVector() As CVector
Set EdgeVector = VertexArray.JoiningVector(StartPoint, EndPoint)
End Function

Public Function EdgeMiddlePoint() As CVector
Set EdgeMiddlePoint = New CVector
EdgeMiddlePoint.Add VertexArray(StartPoint).GetDirectingVector
EdgeMiddlePoint.Add VertexArray(EndPoint).GetDirectingVector
EdgeMiddlePoint.ScaleVector 0.5
End Function

Public Function EdgeSubdivision(ByVal AdjacentPoint As Long, ByVal Ratio As Double) As CVector
'========================================================================
' Returns a vector that defines a point on edge
' ratio from [0; 1]
'========================================================================
Dim OtherEnd As Long
Dim Added As CVector
If AdjacentPoint = StartPoint Then OtherEnd = EndPoint Else OtherEnd = StartPoint
Set EdgeSubdivision = New CVector
Set Added = New CVector

EdgeSubdivision.InitWith VertexArray(AdjacentPoint).GetDirectingVector
EdgeSubdivision.ScaleVector Ratio
Added.InitWith VertexArray(OtherEnd).GetDirectingVector
Added.ScaleVector 1 - Ratio
EdgeSubdivision.Add Added
End Function
