VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEdges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contains edges of the polyhedron."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CEdge"
Attribute VB_Ext_KEY = "Member0" ,"CEdge"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Contains a collection of CEdge objects;
'also implements a method TotalLength to find out the summary length of all edges
'FindEdge function looks for a specific edge in the collection by the index of vertices
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public VertexArray As CVertices 'handle to the global vertex array of the parent polyhedron

Public Function Add(ByVal P1 As Long, ByVal P2 As Long, Optional sKey As String) As CEdge
'create a new object
Dim objNewMember As CEdge
Set objNewMember = New CEdge
Set objNewMember.VertexArray = Me.VertexArray
objNewMember.StartPoint = P1
objNewMember.EndPoint = P2
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

Public Property Get Item(vntIndexKey As Variant) As CEdge
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

Public Function FindEdge(ByVal P1 As Long, ByVal P2 As Long) As Long
'find an edge that contains two vertices P1 and P2
Dim Z As Long
Dim E As CEdge

For Z = 1 To mCol.Count
    Set E = mCol.Item(Z)
    If (E.StartPoint = P1 And E.EndPoint = P2) Or (E.StartPoint = P2 And E.EndPoint = P1) Then
        FindEdge = Z
        Exit Function
    End If
Next Z

FindEdge = 0
End Function

Public Function TotalLength() As Double
'compute total length of all edges
Dim E As CEdge, S As Double

For Each E In mCol
    S = S + E.Length
Next

TotalLength = S
End Function
