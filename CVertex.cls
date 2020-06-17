VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A polyhedron vertex."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'====================================================================
'A class to represent a point in the three-dimensional space
'Contains simply three coordinates of a points as public data members
'====================================================================
Option Explicit

Private m_X As Double
Private m_Y As Double
Private m_Z As Double

Public GetDirectingVector As New CVector

Public Index As Long
Public Faces As Collection 'list of adjacent faces
Public Edges As Collection 'list of adjacent edges

Public Sub Move(ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double)
Attribute Move.VB_Description = "Change point coordinates."
'====================================================================
'Move the vertex to a new position
'====================================================================
X = dX
Y = dY
Z = dZ
End Sub

Private Sub Class_Initialize()
Set Faces = New Collection
Set Edges = New Collection
End Sub

Private Sub Class_Terminate()
Set Faces = Nothing
Set Edges = Nothing
End Sub

Public Function DirectingVector() As CVector
Set DirectingVector = New CVector
DirectingVector.X = m_X
DirectingVector.Y = m_Y
DirectingVector.Z = m_Z
End Function

Public Sub CreateVisualModel(Canvas As ctlPolyViewer, Offset As CVector, Optional ByVal Color As Long = 0, Optional ByVal Visible As Boolean = True)
'====================================================================
' output 3D graphic data to a Canvas object (usually ctlPolyViewer)
'====================================================================
Canvas.AddPoint m_X - Offset.X, m_Y - Offset.Y, m_Z - Offset.Z, Color, Visible
End Sub

Public Property Get X() As Double
X = m_X
End Property

Public Property Let X(ByVal vNewValue As Double)
m_X = vNewValue
GetDirectingVector.X = vNewValue
End Property

Public Property Get Y() As Double
Y = m_Y
End Property

Public Property Let Y(ByVal vNewValue As Double)
m_Y = vNewValue
GetDirectingVector.Y = vNewValue
End Property

Public Property Get Z() As Double
Z = m_Z
End Property

Public Property Let Z(ByVal vNewValue As Double)
m_Z = vNewValue
GetDirectingVector.Z = vNewValue
End Property
