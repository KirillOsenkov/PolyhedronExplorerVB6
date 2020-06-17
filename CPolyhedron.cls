VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPolyhedron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The polyhedron."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Represents an entire polyhedron - the last and the main object in the hierarchy
Option Explicit

'local variable(s) to hold property value(s)
Private mvarFacets As CFacets 'local copy of Facets collection
Private mvarVertices As CVertices 'represents vertices of a polyhedron
Private mvarEdges As CEdges 'the same with edges

Public WasOriented As Boolean 'was the polyhedron successfully oriented yet??
Dim mVolume As Double '-1 if Volume wasn't calculated;
                                      '-2 if Volume cannot be calculated

Public FileName As String ' filename to save the structure into
Public WasSaved As Boolean ' was it saved into that file??
Private m_CenterOfGravity As CVector

Public MaxVertexNorm As Double

'#############################################
'#############################################

Public Property Set Edges(ByVal vData As CEdges)
Attribute Edges.VB_Description = "Describes edges of the polyhedron."
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Edges = Form1
    Set mvarEdges = vData
End Property

Public Property Get Edges() As CEdges
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Edges
    Set Edges = mvarEdges
End Property

Public Property Set Vertices(ByVal vData As CVertices)
Attribute Vertices.VB_Description = "Vertices of this polyhedron."
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Vertices = Form1
Set mvarVertices = vData
End Property

Public Property Get Vertices() As CVertices
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vertices
Set Vertices = mvarVertices
End Property

Public Property Set Facets(ByVal vData As CFacets)
Attribute Facets.VB_Description = "Facets of this polyhedron."
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Facets = Form1
Set mvarFacets = vData
End Property

Public Property Get Facets() As CFacets
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Facets
Set Facets = mvarFacets
End Property

'====================================================================
' Class events; constructor and destructor
'====================================================================

Private Sub Class_Initialize()
Set mvarVertices = New CVertices
Set mvarFacets = New CFacets 'initialize new property collections when a new object of the class is being created
Set mvarEdges = New CEdges
Set m_CenterOfGravity = New CVector

Set mvarFacets.VertexArray = mvarVertices 'creating all the stuff
Set mvarFacets.Parent = Me
Set mvarEdges.VertexArray = mvarVertices

WasOriented = False
WasSaved = True
mVolume = -1
End Sub

Private Sub Class_Terminate()
'====================================================================
' Terminates owned collections
'====================================================================
Set mvarFacets = Nothing 'destroy the collections
Set mvarEdges = Nothing
Set mvarVertices = Nothing
Set m_CenterOfGravity = Nothing
End Sub

Public Sub Clear()
'====================================================================
'kill everything
'====================================================================
mvarEdges.Clear
mvarFacets.Clear
mvarVertices.Clear
WasOriented = False

AuxErrorString = ""
End Sub

'====================================================================
'====================================================================
'                                                       SelfTest
'====================================================================
'====================================================================

Public Property Get CheckConsistency() As Boolean
'====================================================================
'check whether each polygon contains only correct vertex references
'====================================================================
Dim V As CFacet

For Each V In mvarFacets
    If Not V.Consistent Then
        CheckConsistency = False
    End If
Next

CheckConsistency = True
End Property

Public Property Get CheckPlanarity() As Boolean
'====================================================================
'check whether each facet is planar
'====================================================================
Dim V As CFacet

For Each V In mvarFacets
    If Not V.Planar Then
        CheckPlanarity = False
        Exit Property
    End If
Next

CheckPlanarity = True
End Property

Public Property Get CheckConvexity() As Boolean
'====================================================================
' check whether each facet is convex
'====================================================================
Dim V As CFacet

For Each V In mvarFacets
    If Not V.Convex Then
        CheckConvexity = False
        AuxErrorString = Replace(AuxErrorString, "%1", V.Index)
        Exit Property
    End If
Next

CheckConvexity = True
End Property

Public Function Planarize(Optional ByVal Quiet As Boolean = False) As Boolean
'====================================================================
' force planarization of each face;
' can often lead to a failure, not very reliable;
' although worth to give it a try :-)
'====================================================================
Const MaxIterations = 10 'we have to stop sometimes...
Dim Z As Long, V As CFacet
Dim StillPlanar As Boolean
AuxErrorString = ""

StillPlanar = Me.CheckPlanarity
'If Not StillPlanar Then MsgBox AuxErrorString, vbOKOnly + vbInformation, frmMainMDI.Caption
If Not StillPlanar Then
    If MsgBox(AuxErrorString & vbCrLf & vbCrLf & GetString(ResMsg_WishToCorrect), vbYesNo + vbQuestion, frmMainMDI.Caption) = vbYes Then
        'correction
'        Do While Z < MaxIterations And Not StillPlanar
'            For Each V In mvarFacets
'                If Not V.Planar Then V.Planarize
'            Next
'            StillPlanar = Me.CheckPlanarity
'            Z = Z + 1
'        Loop

        If StillPlanar = False Then
            If MsgBox(GetString(ResMsg_ChangesNotMade), vbOKOnly + vbExclamation, frmMainMDI.Caption) = vbYes Then
            'If MsgBox(GetString(ResMsg_ChangesNotMade) & vbCrLf & vbCrLf & GetString(ResMsg_WishToTriangulate), vbYesNo + vbQuestion, frmMainMDI.Caption) = vbYes Then
                'triangulate
                StillPlanar = Me.CheckPlanarity
            'Else
                'all bad
            End If
            Planarize = False
            Exit Function
        Else
            'successful correction
            MsgBox GetString(ResMsg_ChangesMade), vbOKOnly + vbInformation, frmMainMDI.Caption
        End If

'    Else 'refused from correction; triangulating...
'
'        If MsgBox(GetString(ResMsg_WishToTriangulate), vbYesNo + vbQuestion, frmMainMDI.Caption) = vbYes Then
'            'triangulate
'            StillPlanar = Me.CheckPlanarity
'        Else
'            'all bad
'        End If
    End If
End If

Planarize = StillPlanar
End Function

Public Function Orient(Optional ByVal Quiet As Boolean = False) As Boolean
'====================================================================
' Orient the polyhedron recursively
'====================================================================
Dim R As OrientationResultType
AuxErrorString = ""

R = mvarFacets.Orient
Select Case R
Case oreOK
    WasOriented = True
Case oreNonOrientable
    WasOriented = False
    MsgBox AuxErrorString & vbCrLf & vbCrLf & GetString(ResMsg_ExplorationImpossible), vbOKOnly + vbExclamation, frmMainMDI.Caption
Case oreNotClosed
    WasOriented = False
    MsgBox AuxErrorString & vbCrLf & vbCrLf & GetString(ResMsg_ExplorationImpossible), vbOKOnly + vbExclamation, frmMainMDI.Caption
Case oreNotConnected
    WasOriented = False
    MsgBox AuxErrorString & vbCrLf & vbCrLf & GetString(ResMsg_ExplorationImpossible), vbOKOnly + vbExclamation, frmMainMDI.Caption
Case oreTooManyFaces
    WasOriented = False
    MsgBox AuxErrorString & vbCrLf & vbCrLf & GetString(ResMsg_ExplorationImpossible), vbOKOnly + vbExclamation, frmMainMDI.Caption
End Select

Orient = WasOriented
End Function

Public Property Get Oriented() As Boolean
Oriented = WasOriented
End Property

Public Property Get CheckOrientation() As Boolean
CheckOrientation = mvarFacets.CheckOrientation = oreOK
End Property

Public Sub ClearOrientation()
'====================================================================
' Resets the orientation to default
'====================================================================
WasOriented = False
mvarFacets.ClearOrientation
End Sub

Public Sub ChangeOrientation()
'====================================================================
' Reverses orientation to the opposide
' (turns Normals inside out)
'====================================================================
mvarFacets.ChangeOrientation
End Sub

Public Function SelfTest(Optional ByVal Quiet As Boolean = False, Optional ByVal ShouldClear As Boolean = True) As Boolean
'====================================================================
' make sure that:
' - Me is a real polyhedron
' - contains only valid vertex references
' - all the faces are planar
' - the polyhedron is connected
' - polyhedron can be oriented
' if everything is OK then calculate volume
'====================================================================
If Me.IsEmptyPolyhedron Or Me.Vertices.Count < 4 Or Me.Facets.Count < 4 Then
    SelfTest = False
    If Not Quiet Then MsgBox GetString(ResMsg_EmptyPolyhedron), vbOKOnly + vbExclamation, GetString(ResMsg_UnableToLoadPolyhedron)
    Exit Function
End If

Me.Recalc

If Not Me.CheckConsistency Then
    SelfTest = False
    If Not Quiet Then MsgBox GetString(ResMsg_UnexistingVertex), vbOKOnly Or vbExclamation, GetString(ResMsg_UnableToLoadPolyhedron)
    If ShouldClear Then Me.Clear
    Exit Function
End If
If Not Me.Planarize(Quiet) Then
    SelfTest = False
    If Not Quiet Then MsgBox GetString(ResMsg_UnableToLoadPolyhedron), vbOKOnly Or vbExclamation, frmMainMDI.Caption
    If ShouldClear Then Me.Clear
    Exit Function
End If
If Not Me.Orient(Quiet) Then
    Me.ClearOrientation
End If

If Me.Oriented Then
    mVolume = CalculateVolume
    Me.AddEdges
End If

AuxErrorString = ""
End Function

'====================================================================
'====================================================================
'                                           Edge procedures
'                               restoring connectivity information
'====================================================================
'====================================================================

Public Function AddEdges() As Boolean
'====================================================================
' Add edge information to the data structure;
' Restore connectivity information
'====================================================================
On Local Error GoTo EH
Dim e As CEdge, F As CFacet
Dim Z As Long, Q As Long, EI As Long
Dim P1 As Long, P2 As Long

mvarEdges.Clear ' begin a new honest life :-)

For Each F In mvarFacets
    For Z = 1 To F.Vertices.Count
        P1 = F.Vertices(Z)
        P2 = F.Vertices(IIf(Z = F.Vertices.Count, 1, Z + 1))
        If P1 = 0 Or P2 = 0 Then GoTo EH
        EI = mvarEdges.FindEdge(P1, P2)
        If EI = 0 Then
            Set e = mvarEdges.Add(P1, P2)
                e.LeftFace = F.Index
                Q = FindFaceWithEdge(P2, P1)
                If Q <> 0 Then e.RightFace = Q Else GoTo EH
            Set e = Nothing
            EI = mvarEdges.Count
        End If
        F.Edges.Add EI
    Next
Next

AddEdgeConnections

AddEdges = True
Exit Function

EH:
MsgBox GetString(ResMsg_CannotReconstructEdges), vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
AddEdges = False
Err.Clear
End Function

Public Function AddEdgeConnections() As Boolean
'====================================================================
'
'====================================================================
On Local Error GoTo EH
Dim e As CEdge, F As CFacet, V As CVertex
Dim Z As Long, Q As Long, EI As Long
Dim P1 As Long, P2 As Long, N As Long, M As Long, I As Long

For Each V In mvarVertices
    I = V.Index
    N = FindFaceWithVertex(I)
    If N = 0 Then GoTo EH
    M = N
    Do
        Q = mvarFacets(M).PreviousVertex(I, logGlobalFromGlobal)
        If Q = 0 Then GoTo EH
        M = mvarEdges.FindEdge(I, Q)
        If M = 0 Then GoTo EH
        V.Edges.Add M
        M = FindFaceWithEdge(I, Q)
        If M = 0 Then GoTo EH
        V.Faces.Add M
    Loop Until M = N
Next

AddEdgeConnections = True
Exit Function

EH:
MsgBox GetString(ResMsg_CannotReconstructEdges) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
AddEdgeConnections = False
Err.Clear
End Function

Public Function FindFaceWithEdge(ByVal P1 As Long, ByVal P2 As Long, Optional ByVal ConsiderOrientation As Boolean = True) As Long
'====================================================================
' Two vertices P1 and P2 are given in Global Indexing
' Can return considering orientation and not considering
'====================================================================
Dim Z As Long, F As CFacet, V   As Long

For Each F In mvarFacets
    V = F.FindVertex(P1)
    If V <> 0 Then
        If ConsiderOrientation Then
            If F.Vertices(F.NextVertex(V)) = P2 Then
                FindFaceWithEdge = F.Index
                Exit Function
            End If
        Else
            If F.Vertices(F.NextVertex(V)) = P2 Or F.Vertices(F.PreviousVertex(V)) = P2 Then
                FindFaceWithEdge = F.Index
                Exit Function
            End If
        End If
    End If
Next

FindFaceWithEdge = 0
End Function

Public Function FindFaceWithVertex(ByVal P1 As Long) As Long
'====================================================================
' Finds at least one face with vertex P1 given in Global Indexing
'====================================================================
Dim F As CFacet

For Each F In mvarFacets
    If F.FindVertex(P1) > 0 Then
        FindFaceWithVertex = F.Index
        Exit Function
    End If
Next
End Function

'====================================================================
'====================================================================
' Procedures that generate visual output to the specified context
'====================================================================
'====================================================================

Public Sub CreateVisualModel(Canvas As ctlPolyViewer)
'====================================================================
' output 3D graphic data to a Canvas object (usually ctlPolyViewer)
'====================================================================
Const MinCenterOffset = 0.1
Dim V As CVertex
Dim C As CVector
Dim F As CFacet
Dim e As CEdge

Set C = m_CenterOfGravity
If C.Norm < MinCenterOffset Then Set C = CNullVector

For Each V In mvarVertices
    V.CreateVisualModel Canvas, C
Next

For Each e In mvarEdges
    e.CreateVisualModel Canvas
Next

For Each F In mvarFacets
    F.CreateVisualModel Canvas
Next

Canvas.CompileStructure
Canvas.Refresh
End Sub

Public Sub OutputSphericalImage(Canvas As ctlPolyViewer)
'====================================================================
' output 3D graphic data required to draw polyhedron's Spherical Image
' to a Canvas object (usually ctlPolyViewer)
'====================================================================
Const SphericalDetails = 20

Dim V As CVertex
Dim F As CFacet
Dim e As CEdge

Dim N As CVector
Dim A As CVector, B As CVector, C As CVector
Dim T As Double, Ang As Double, Alpha As Double, Percentage As Double

Dim P() As Long
Dim Z As Long
Dim Q As Long, PrevPoint As Long

Q = mvarFacets.Count
For Each F In mvarFacets
    Set N = F.Normal
    N.Normalize
    Canvas.AddPoint N.X, N.Y, N.Z, F.Color
Next

Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim C1 As Long, C2 As Long

For Each e In mvarEdges
    Set A = mvarFacets.Item(e.LeftFace).Normal
    Set B = mvarFacets.Item(e.RightFace).Normal
    A.Normalize
    B.Normalize
    Ang = AngleTwoCVectors(A, B) / SphericalDetails
    Set B = CGrahamSchmidtOrthogonalize(A, B)
    
    C1 = mvarFacets(e.LeftFace).Color
    C2 = mvarFacets(e.RightFace).Color
    R1 = Red(C1)
    G1 = Green(C1)
    B1 = Blue(C1)
    R2 = Red(C2)
    G2 = Green(C2)
    B2 = Blue(C2)
    
    ' Point index is assigned face index;
    ' this is NOT an error; in spherical image points and faces
    ' exchange roles
    PrevPoint = e.LeftFace '????? may be a bug with For Each order (1 to n)?
    
    For Z = 1 To SphericalDetails - 1
        Alpha = Z * Ang
        Percentage = Z / SphericalDetails
        Set C = CVectorLinearCombination(Cos(Alpha), A, Sin(Alpha), B)
        C.Normalize
        Canvas.AddPoint C.X, C.Y, C.Z, RGB(R1 + (R2 - R1) * Percentage, G1 + (G2 - G1) * Percentage, B1 + (B2 - B1) * Percentage), False
        Q = Q + 1
        Canvas.AddLine PrevPoint, Q
        PrevPoint = Q
    Next
    
    Canvas.AddLine PrevPoint, e.RightFace
Next

Canvas.CompileStructure
Canvas.Refresh
End Sub

Public Sub OutputScanningModel(Canvas As ctlPolyViewer, Stage As Double)
'====================================================================
' Output 3D graphic data for a polyhedron scanning at the
' specified stage of "Unfoldedness"
'====================================================================
Dim Z As Long

MaxVertexNorm = 0.1

For Z = 1 To mvarFacets.Count
    mvarFacets(Z).Unwrapped = False
Next

mvarFacets(1).OutputScanningBranch Stage, 0
End Sub

'====================================================================
'====================================================================
'                                            File I/O procedures...
'====================================================================
'====================================================================

Public Sub SaveAs()
'====================================================================
' SaveAs saves current polyhedron with a new name
' prompts for a name
'====================================================================
CD.FileName = Me.FileName
CD.Filter = GetString(ResMsg_PolFile) & " (*." & extApp & ")|*." & extApp
CD.Flags = &H4 + &H2
If Not IsValidPath(setLastPolPath) Then setLastPolPath = AppPath
CD.InitDir = setLastPolPath
CD.ShowSave
If CD.Cancelled Then Exit Sub

Me.FileName = CD.FileName
If Right(UCase(Me.FileName), 4) <> "." & extApp Then
    If InStr(RetrieveName(Me.FileName), ".") = 0 Then
        Me.FileName = Me.FileName & "." & extApp
    Else
        Me.FileName = Left(Me.FileName, InStr(Me.FileName, ".")) + extApp
    End If
End If
setLastPolPath = AddDirSep(RetrieveDir(Me.FileName))
If Not IsValidPath(setLastPolPath) Then setLastPolPath = AppPath

Document(ActiveWindow).Caption = RetrieveName(Me.FileName)

Save
End Sub

Public Sub Save()
'====================================================================
' Save writes current polyhedron data to disk
'====================================================================
If Me.FileName = "" Then SaveAs

Screen.MousePointer = vbHourglass
DoEvents

SaveFile

AddMRUItem Me.FileName
Screen.MousePointer = vbDefault
End Sub

Public Sub OpenAFile()
'====================================================================
'
'====================================================================
If Not WasSaved Then
    Select Case MsgBox(GetString(ResMsg_Save) & " " & Me.FileName & "?", vbYesNoCancel)
        Case vbYes
            Save
        Case vbNo
            'do nothing
        Case vbCancel
            Exit Sub
    End Select
End If

CD.Filter = GetString(ResMsg_PolFile) & " (*." & extApp & ")|*." & extApp
CD.Flags = &H1000 + &H4
If Not IsValidPath(setLastPolPath) Then setLastPolPath = AppPath
CD.InitDir = setLastPolPath
CD.ShowOpen
If CD.Cancelled = True Or Dir(CD.FileName) = "" Then Exit Sub

Me.FileName = CD.FileName
setLastPolPath = AddDirSep(RetrieveDir(Me.FileName))
If Not IsValidPath(setLastPolPath) Then setLastPolPath = AppPath

Screen.MousePointer = vbHourglass
Document(ActiveWindow).Caption = RetrieveName(Me.FileName)
DoEvents

OpenFile

AddMRUItem Me.FileName
Screen.MousePointer = vbDefault
End Sub

Public Sub OpenFileNamed(ByVal FName As String)
'====================================================================
'
'====================================================================
Me.FileName = FName

Screen.MousePointer = vbHourglass
Document(ActiveWindow).Caption = RetrieveName(Me.FileName)
DoEvents

OpenFile

AddMRUItem Me.FileName
Screen.MousePointer = vbDefault
End Sub

Public Sub SaveFile()
'====================================================================
' Physically writes polyhedron data to disk
' into the file named Me.FileName
'====================================================================
On Local Error GoTo EH:
Dim Z As Long, Q As Long, A As String
Dim FName As String

FName = Me.FileName
'If Dir(FName, 23) <> "" Then
'    If MsgBox(GetString(ResMsg_FileAlreadyExists), vbYesNo + vbQuestion, frmMainMDI.Caption) = vbNo Then Exit Sub
'End If

Open FName For Output As #1
    Print #1, "PointCount=" & mvarVertices.Count
    Print #1, "CountOfFaces=" & mvarFacets.Count
    
    With mvarVertices
        For Z = 1 To .Count
            Print #1, "X" & Z & "=" & Trim(Str(.Item(Z).X))
            Print #1, "Y" & Z & "=" & Trim(Str(.Item(Z).Y))
            Print #1, "Z" & Z & "=" & Trim(Str(.Item(Z).Z))
        Next
        Print #1,
    End With
    
    With mvarFacets
        For Z = 1 To .Count
            A = ""
            For Q = 1 To .Item(Z).Vertices.Count
                A = A & IIf(Q = 1, "", ",") & .Item(Z).Vertices(Q)
            Next
            Print #1, "Face" & Z & "=" & A
            Print #1, "ColorOfFace" & Z & "=" & .Item(Z).Color
        Next Z
    End With
Close #1

Exit Sub
EH:
End Sub

Public Sub OpenFile()
'====================================================================
' This factically opens a file from disk
' and loads the data into This polyhedron
'====================================================================
On Local Error GoTo EH:
Dim Z As Long, Q As Long, A As String
Dim FName As String
Dim AName As String, AValue As String

Dim CurPoint As Long, CurFacet As Long, CurX As Double, CurY As Double, CurZ As Double
Dim F As CFacet, VA As Variant

FName = Me.FileName
If Dir(FName) = "" Then
    MsgBox GetString(ResMsg_UnableToLoadPolyhedron), vbOKOnly + vbExclamation
    Exit Sub
End If

Me.Clear

Open FName For Input As #1
    Do While Not EOF(1)
        Line Input #1, A
        Q = InStr(A, "=")
        If Q > 0 Then
            AName = Left(A, Q - 1)
            AValue = Right(A, Len(A) - Q)
            
            If Left(AName, 1) = "X" Then CurX = Val(AValue)
            If Left(AName, 1) = "Y" Then CurY = Val(AValue)
            If Left(AName, 1) = "Z" Then
                CurZ = Val(AValue)
                Me.Vertices.Add CurX, CurY, CurZ, Right(AName, Len(AName) - 1)
            End If
            If Left(AName, 4) = "Face" Then
                Set F = Me.Facets.Add
                VA = Split(AValue, ",")
                For Z = LBound(VA) To UBound(VA)
                    F.Vertices.Add CLng(VA(Z))
                Next
                Set F = Nothing
            End If
            If Left(AName, 11) = "ColorOfFace" Then
                CurZ = Val(Right(AName, Len(AName) - 11))
                Me.Facets(CurZ).Color = AValue
            End If
        End If
    Loop
Close #1

Me.SelfTest
Document(ActiveWindow).Mode = wmdViewer

Exit Sub

EH:
End Sub

Public Sub Export()
'====================================================================
'
'====================================================================
Dim XL As Excel.Application
Dim WB As Excel.Workbook
Dim WS As Excel.Worksheet
Dim Z As Long, Q As Long
On Local Error GoTo EH:

Set XL = New Excel.Application
If Dir(MainFileName) <> "" And Right(MainFileName, 4) = "." & extXLS Then
    Set WB = XL.Workbooks.Open(MainFileName)
Else
    Set WB = XL.Workbooks.Add
End If
Set WS = WB.Worksheets.Add

For Z = 1 To WB.Worksheets.Count
    If WB.Worksheets(Z).Name = MainSheetName Then MainSheetName = ""
Next
If MainSheetName <> "" Then
    WS.Name = MainSheetName
End If

For Z = 1 To mvarVertices.Count
    WS.Cells(Z, 1).Value = Trim(Str(mvarVertices(Z).X))
    WS.Cells(Z, 2).Value = Trim(Str(mvarVertices(Z).Y))
    WS.Cells(Z, 3).Value = Trim(Str(mvarVertices(Z).Z))
Next

For Z = 1 To mvarFacets.Count
    For Q = 1 To mvarFacets(Z).Vertices.Count
        WS.Cells(Z, Q + 3).Value = mvarFacets(Z).Vertices(Q)
    Next
Next

If Dir(MainFileName) <> "" And Right(MainFileName, 4) = "." & extXLS Then
    WB.Save
Else
    WB.SaveAs MainFileName
End If

WB.Close
XL.Quit
Set WS = Nothing
Set WB = Nothing
Set XL = Nothing

If Dir(MainFileName) <> "" Then MsgBox GetString(ResMsg_DataExportComplete) & ".", vbOKOnly + vbInformation, RetrieveName(MainFileName)
MainFileName = ""
MainSheetName = ""
Exit Sub

EH:
End Sub

'======================================================
'======================================================
'                                   Operations on polyhedra
'======================================================
'======================================================

Public Function Stellate(Optional ByVal Percentage As Double = 1) As CPolyhedron
'====================================================================
' Creates a new polyhedron based on Me with pyramids attached to each face
'====================================================================
On Local Error GoTo EH

Dim Poly As CPolyhedron
Dim F As CFacet, NF As CFacet
Dim Z As Long, Q As Long, NV As Long
Dim Summit As CVector, CG As CVector

AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
If Not Me.CheckConvexity Then
    MsgBox AuxErrorString & " " & GetString(ResMsg_CannotStellatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
    AuxErrorString = ""
    Exit Function
End If
AuxErrorString = ""


Set Poly = New CPolyhedron

For Z = 1 To mvarVertices.Count
    Poly.Vertices.Add mvarVertices(Z).X, mvarVertices(Z).Y, mvarVertices(Z).Z, Z
Next

For Each F In mvarFacets
    
    Set CG = F.GetCenterOfGravity
    
    Set Summit = New CVector
    Summit.InitWith F.Normal
    Summit.Normalize
    Summit.ScaleVector Percentage
    Summit.Add CG
    
    NV = Poly.Vertices.Count + 1
    Poly.Vertices.Add Summit.X, Summit.Y, Summit.Z, NV
    
    Set Summit = Nothing
    
    For Q = 1 To F.Vertices.Count
        Set NF = Poly.Facets.Add
        With NF
            .Vertices.Add F.Vertices(Q)
            If Q < F.Vertices.Count Then
                .Vertices.Add F.Vertices(Q + 1)
            Else
                .Vertices.Add F.Vertices(1)
            End If
            .Vertices.Add NV
        End With
        NF.Recalc
        Set NF = Nothing
    Next
Next

Poly.WasOriented = True
Poly.Volume = Poly.CalculateVolume
Poly.AddEdges

Set Stellate = Poly
Exit Function

EH:
End Function

'
''====================================================================
'' Creates a new polyhedron based on Me with pyramids attached to each face
''====================================================================
'
'Public Function Stellate(Optional ByVal Percentage As Double = 1) As CPolyhedron
'On Local Error GoTo EH
'
'Dim Poly As CPolyhedron
'Dim F As CFacet, NF As CFacet
'Dim Z As Long, Q As Long, NV As Long
'Dim Summit As Vector, CG As Vector, V1 As Vector, V2 As Vector
'
'AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
'If Not Me.CheckConvexity Then
'    MsgBox AuxErrorString & " " & GetString(ResMsg_CannotStellatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
'    AuxErrorString = ""
'    Exit Function
'End If
'AuxErrorString = ""
'
'
'Set Poly = New CPolyhedron
'
'For Z = 1 To mvarVertices.Count
'    Poly.Vertices.Add mvarVertices(Z).X, mvarVertices(Z).Y, mvarVertices(Z).Z, Z
'Next
'
'For Each F In mvarFacets
'
'    CG = MakeVector(F.CenterOfGravity)
'    V1 = DirectingVector(mvarVertices, F.Vertices(1))
'    V2 = DirectingVector(mvarVertices, F.Vertices(2))
'    'Summit = (CG + [V1 - CG, V2 - CG] ) * Percentage
'    Summit = VectorSum(CG, ScaleVector(VectorProduct(VectorDifference(V1, CG), VectorDifference(V2, CG)), Percentage))
'    'Summit = ScaleVector(Summit, Percentage)
'
'    NV = Poly.Vertices.Count + 1
'    Poly.Vertices.Add Summit.X, Summit.Y, Summit.Z, NV
'
'    For Q = 1 To F.Vertices.Count
'
'        Set NF = Poly.Facets.Add
'        With NF
'            .Vertices.Add F.Vertices(Q)
'            If Q < F.Vertices.Count Then
'                .Vertices.Add F.Vertices(Q + 1)
'            Else
'                .Vertices.Add F.Vertices(1)
'            End If
'            .Vertices.Add NV
'        End With
'        Set NF = Nothing
'
'    Next
'
'Next
'
'Poly.WasOriented = True
'Poly.Volume = Poly.CalculateVolume
'Poly.AddEdges
'
'Set Stellate = Poly
'Exit Function
'
'EH:
'End Function

Public Function Dual() As CPolyhedron
'====================================================================
' Create and return a new polyhedron which is dual to Me
'====================================================================
On Local Error GoTo EH:
Dim Poly As CPolyhedron

Dim F As CFacet, NF As CFacet
Dim Z As Long, Q As Long
Dim CG As CVector

AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
If Not Me.CheckConvexity Then
    MsgBox AuxErrorString & " " & GetString(ResMsg_DualNotDefined), vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
    AuxErrorString = ""
    Exit Function
End If
AuxErrorString = ""

Set Poly = New CPolyhedron

For Z = 1 To mvarFacets.Count
    Set CG = mvarFacets(Z).GetCenterOfGravity
    Poly.Vertices.Add CG.X, CG.Y, CG.Z, Z
Next

For Z = 1 To mvarVertices.Count
    Set NF = Poly.Facets.Add
        For Q = 1 To mvarVertices(Z).Faces.Count
            NF.Vertices.Add mvarVertices(Z).Faces(Q)
        Next
    Set NF = Nothing
Next

Poly.SelfTest

If Poly.IsEmptyPolyhedron Then
    Set Dual = Nothing
    Exit Function
End If

Set Dual = Poly
Set Poly = Nothing
Exit Function

EH:
MsgBox GetString(ResMsg_DualNotDefined), vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
Set Dual = Nothing
Err.Clear
End Function

Public Function EdgeDual() As CPolyhedron
'====================================================================
' Create and return a new polyhedron which is edge-dual to Me
'====================================================================
On Local Error GoTo EH
Dim Poly As CPolyhedron
Dim NF As CFacet
Dim Z As Long, Q As Long
Dim EMP As CVector

AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
If Not Me.CheckConvexity Then
    MsgBox AuxErrorString & " " & GetString(ResMsg_EdgeDualNotDefined) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
    AuxErrorString = ""
    Exit Function
End If
AuxErrorString = ""

Set Poly = New CPolyhedron

For Z = 1 To mvarEdges.Count
    Set EMP = mvarEdges(Z).EdgeMiddlePoint
    Poly.Vertices.Add EMP.X, EMP.Y, EMP.Z, Z
Next

For Z = 1 To mvarFacets.Count
    Set NF = Poly.Facets.Add
        For Q = 1 To mvarFacets(Z).Edges.Count
            NF.Vertices.Add mvarFacets(Z).Edges(Q)
        Next
    Set NF = Nothing
Next

For Z = 1 To mvarVertices.Count
    Set NF = Poly.Facets.Add
        For Q = 1 To mvarVertices(Z).Edges.Count
            NF.Vertices.Add mvarVertices(Z).Edges(Q)
        Next
    Set NF = Nothing
Next

'Poly.WasOriented = True
'Poly.Volume = Poly.CalculateVolume
'Poly.AddEdges

Poly.SelfTest

If Poly.IsEmptyPolyhedron Then
    Set EdgeDual = Nothing
    Exit Function
End If

Set EdgeDual = Poly
Set Poly = Nothing
Exit Function

EH:
MsgBox AuxErrorString & " " & GetString(ResMsg_EdgeDualNotDefined) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
Err.Clear
AuxErrorString = ""
End Function
'
Public Function Truncate(Optional ByVal Ratio As Double = 0.5) As CPolyhedron
'====================================================================
' Create and return a polyhedron which is truncated Me
' Truncation replaces vertices by new faces, adding new vertices, faces and edges
'====================================================================
On Local Error GoTo EH
Dim Poly As CPolyhedron

Dim F As CFacet, NF As CFacet, e As CEdge
Dim Z As Long, Q As Long, NV As Long
Dim Summit As CVector, CG As CVector, V1 As CVector, V2 As CVector
Dim EMP As CVector
Dim V As CVertex

AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
If Not Me.CheckConvexity Then
    MsgBox AuxErrorString & " " & GetString(ResMsg_CannotTruncatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
    AuxErrorString = ""
    Exit Function
End If
AuxErrorString = ""

Set Poly = New CPolyhedron

Dim V1 As Long, V2 As Long, e As CEdge

For Z = 1 To mvarVertices.Count
    Set V = mvarVertices(Z)
    For Q = 1 To V.Edges.Count
        
    Next
Next

For Z = 1 To mvarEdges.Count
    Set e = mvarEdges(Z)
    V1 = e.StartPoint
    V2 = e.EndPoint
    Poly.Vertices.Add EMP.X, EMP.Y, EMP.Z, 2 * Z - 1
    Poly.Vertices.Add EMP.X, EMP.Y, EMP.Z, 2 * Z
Next

For Z = 1 To mvarFacets.Count
    Set NF = Poly.Facets.Add
        For Q = 1 To mvarFacets(Z).Edges.Count
            NV = CLng(mvarFacets(Z).Vertices(Q) = mvarEdges(mvarFacets(Z).Edges(Q)).StartPoint)
            NF.Vertices.Add 2 * mvarFacets(Z).Edges(Q) + NV
            NF.Vertices.Add 2 * mvarFacets(Z).Edges(Q) - 1 - NV
        Next
    Set NF = Nothing
Next

For Z = 1 To mvarVertices.Count
    Set NF = Poly.Facets.Add
        For Q = 1 To mvarVertices(Z).Edges.Count
            Set e = mvarEdges(mvarVertices(Z).Edges(Q))
                'V1 = DirectingVector(Poly.Vertices, 2 * E.Index - 1)
                'V2 = DirectingVector(Poly.Vertices, 2 * E.Index)
                'CG = DirectingVector(mvarVertices, Z)
                'NV = IIf(DistanceTwoVectors(V1, CG) > DistanceTwoVectors(V2, CG), 2 * E.Index, 2 * E.Index - 1)
                NF.Vertices.Add NV
            Set e = Nothing
        Next
    Set NF = Nothing
Next

Poly.SelfTest

If Poly.IsEmptyPolyhedron Then
    GoTo EH
End If

Set Truncate = Poly
Set Poly = Nothing
Exit Function

EH:
MsgBox GetString(ResMsg_CannotTruncatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption

''====================================================================
'' Create and return a polyhedron which is truncated Me
'' Truncation replaces vertices by new faces, adding new vertices, faces and edges
''====================================================================
'On Local Error GoTo EH
'Dim Poly As CPolyhedron
'Exit Function
'
'Dim F As CFacet, NF As CFacet, E As CEdge
'Dim Z As Long, Q As Long, NV As Long
'Dim Summit As CVector, CG As CVector, V1 As CVector, V2 As CVector, EMP As CVector
'
'AuxErrorString = GetString(ResMsg_ContainsNonConvexEdges)
'If Not Me.CheckConvexity Then
'    MsgBox AuxErrorString & " " & GetString(ResMsg_CannotTruncatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
'    AuxErrorString = ""
'    Exit Function
'End If
'AuxErrorString = ""
'
'Set Poly = New CPolyhedron
'
'For Z = 1 To mvarEdges.Count
'    EMP = EdgeSubdivision(mvarVertices, mvarEdges(Z).StartPoint, mvarEdges(Z).EndPoint, Ratio / 2)
'    Poly.Vertices.Add EMP.X, EMP.Y, EMP.Z, 2 * Z - 1
'    EMP = EdgeSubdivision(mvarVertices, mvarEdges(Z).StartPoint, mvarEdges(Z).EndPoint, 1 - Ratio / 2)
'    Poly.Vertices.Add EMP.X, EMP.Y, EMP.Z, 2 * Z
'Next
'
'For Z = 1 To mvarFacets.Count
'    Set NF = Poly.Facets.Add
'        For Q = 1 To mvarFacets(Z).Edges.Count
'            NV = CLng(mvarFacets(Z).Vertices(Q) = mvarEdges(mvarFacets(Z).Edges(Q)).StartPoint)
'            NF.Vertices.Add 2 * mvarFacets(Z).Edges(Q) + NV
'            NF.Vertices.Add 2 * mvarFacets(Z).Edges(Q) - 1 - NV
'        Next
'    Set NF = Nothing
'Next
'
'For Z = 1 To mvarVertices.Count
'    Set NF = Poly.Facets.Add
'        For Q = 1 To mvarVertices(Z).Edges.Count
'            Set E = mvarEdges(mvarVertices(Z).Edges(Q))
'                V1 = DirectingVector(Poly.Vertices, 2 * E.Index - 1)
'                V2 = DirectingVector(Poly.Vertices, 2 * E.Index)
'                CG = DirectingVector(mvarVertices, Z)
'                NV = IIf(DistanceTwoVectors(V1, CG) > DistanceTwoVectors(V2, CG), 2 * E.Index, 2 * E.Index - 1)
'                NF.Vertices.Add NV
'            Set E = Nothing
'        Next
'    Set NF = Nothing
'Next
'
'Poly.SelfTest
'
'If Poly.IsEmptyPolyhedron Then
'    GoTo EH
'End If
'
Set Truncate = Me
'Set Poly = Nothing
'Exit Function
'
'EH:
'MsgBox GetString(ResMsg_CannotTruncatePolyhedron) & ".", vbOKOnly + vbExclamation, Document(ActiveWindow).Caption
End Function

'====================================================================
'====================================================================
' Extracting characteristics and properties of the polyhedron
'====================================================================
'====================================================================

Public Property Get EulerCharacteristic() As Long
'====================================================================
' Self-explanatory...
'====================================================================
EulerCharacteristic = mvarVertices.Count - mvarEdges.Count + mvarFacets.Count
End Property

Public Property Get GetCenterOfGravity() As CVector
'====================================================================
'
'====================================================================
Set GetCenterOfGravity = m_CenterOfGravity
End Property

Private Sub CalculateCenterOfGravity()
m_CenterOfGravity.InitWith mvarVertices.CenterOfGravity
End Sub

Public Sub Recalc()
Dim F As CFacet
CalculateCenterOfGravity
For Each F In mvarFacets
    F.Recalc
Next
End Sub

Public Function AngleBetweenTwoFaces(ByVal F1 As Long, ByVal F2 As Long) As Double
'====================================================================
' Faces are given in Local Indexing
'====================================================================
AngleBetweenTwoFaces = AngleTwoCVectors(mvarFacets(F1).Normal, mvarFacets(F2).Normal)
End Function

Public Property Get SurfaceArea() As Double
'====================================================================
' Calculate total surface area of the polyhedron and return it
'====================================================================
SurfaceArea = mvarFacets.SurfaceArea
End Property

Public Property Let Volume(ByVal vNewVolume As Double)
mVolume = vNewVolume
End Property

Public Property Get Volume() As Double
Volume = mVolume
End Property

Public Function CalculateVolume() As Double 'does what it says
'====================================================================
' This function solves the main problem of this work:
' to calculate the volume of a polyhedron
' Besides; it also reorients the polyhedron properly
'====================================================================
Dim F As CFacet
Dim Sum As Double
Sum = 0

For Each F In mvarFacets 'for every face
    Sum = Sum + F.PyramidVolume 'add to the Sum the volume of a pyramid with vertex at the origin
Next

If Sum < 0 Then Me.ChangeOrientation 'if volume < 0 it means faces are
'oriented clockwise, which is bad; as they have to be oriented counterclockwise
'(so that the normals point outside the polyhedron),
'we have to change orientation of every face

CalculateVolume = Abs(Sum)
End Function

Public Property Get IsEmptyPolyhedron() As Boolean
'====================================================================
' Checks whether the polyhedron data structure is empty
'====================================================================
IsEmptyPolyhedron = mvarVertices.Count = 0
End Property
