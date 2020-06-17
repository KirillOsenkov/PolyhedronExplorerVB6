VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "All the facets of a polyhedron."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CFacet"
Attribute VB_Ext_KEY = "Member0" ,"CFacet"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'A collection of CFacet objects
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public VertexArray As CVertices
Public Parent As CPolyhedron

Public Function Add(Optional sKey As String) As CFacet
'create a new object
Dim objNewMember As CFacet
Set objNewMember = New CFacet
Set objNewMember.VertexArray = Me.VertexArray
Set objNewMember.Parent = Me
objNewMember.Index = mCol.Count + 1

'set the properties passed into the method
If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If

'return the object created
Set Add = objNewMember
Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CFacet
Attribute Item.VB_UserMemId = 0
'used when referencing an element in the collection
'vntIndexKey contains either the Index or Key to the collection,
'this is why it is declared as a Variant
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
'used when removing an element from the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: x.Remove(xyz)
mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'this property allows you to enumerate
'this collection with the For...Each syntax
Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
'creates the collection when this class is created
Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
'destroys collection when this class is terminated
Set mCol = Nothing
Set VertexArray = Nothing
End Sub

Public Sub Clear()
Do While mCol.Count > 0
    mCol.Remove 1
Loop
End Sub

Public Function Orient() As OrientationResultType
'orient the entire polyhedron recursively;
'let's see if it is connected
Dim F As CFacet, R As OrientationResultType

Set F = mCol(1)

R = F.Orient(1) 'orient the first facet recursively
If R <> oreOK Then Orient = R: Exit Function

Orient = CheckOrientation
End Function

Public Function CheckOrientation() As OrientationResultType
'check whether all the faces were oriented accordingly to each other's orientation
Dim F As CFacet
CheckOrientation = oreOK
For Each F In Me
    If Not F.Oriented Then
        CheckOrientation = oreNotConnected
        AuxErrorString = GetString(ResMsg_PolyhedronNotConnected)
    End If
Next
End Function

Public Sub ClearOrientation()
Dim F As CFacet

For Each F In Me
    If F.OrientationWasChanged Then F.ChangeOrientation
    F.Orientation = 0
Next
End Sub

Public Sub ChangeOrientation()
'turn the polyhedron inside out
'by changing orientation of all faces
Dim F As CFacet

For Each F In Me
    F.ChangeOrientation
Next
End Sub

Public Property Get SurfaceArea() As Double
'calculate summary surface area of all polyhedron faces
Dim F As CFacet, Sum As Double

For Each F In Me
    Sum = Sum + F.Area
Next

SurfaceArea = Sum
End Property
