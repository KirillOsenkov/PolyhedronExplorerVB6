Attribute VB_Name = "modSystem"
Option Explicit

Public Const AppName = "Volumeter"
Public AppPath As String
Public Const extApp = "POL"
Public Const extXLS = "XLS"

Public Enum Language       'interface language type
    lanEnglish
    lanRussian
End Enum

Public Enum WindowMode 'state of a main window frmMain:
    wmdViewer                    'either in polyhedron viewer mode
    wmdEditor                     'or in polyhedron editor mode
End Enum

Public Enum LocalOrGlobal
    logLocalFromLocal
    logLocalFromGlobal
    logGlobalFromLocal
    logGlobalFromGlobal
End Enum

Public setLanguage As Long
Public setLastExcelPath As String
Public setLastPolPath As String
Public ColorArray(0 To 15) As Long
Public MRUList As New Collection
Public MRUCount As Long

Public CD As CCommonDialog

Public MainFileName As String
Public MainSheetName As String


'#########################################################
'________________
'Resource string IDs

Public Const Res_Caption = 102
Public Const Res_Untitled = 104
Public Const Res_PolyhedronProperties = 106
Public Const Res_EditingMode = 108
Public Const Res_ExplorationMode = 110
Public Const Res_ImportModel = 112
Public Const Res_Cancel = 114
Public Const Res_ImportData = 116
Public Const Res_LocateXLS = 118
Public Const Res_BrowseForXLS = 120
Public Const Res_SelectXLWorksheet = 122
Public Const Res_Worksheets = 124
Public Const Res_WorksheetPreview = 126
Public Const Res_Vertices = 128
Public Const Res_Edges = 130
Public Const Res_Faces = 132
Public Const Res_EulerianSum = 134
Public Const Res_ClearModel = 136
Public Const Res_Volume = 138
Public Const Res_SurfaceArea = 140
Public Const Res_ExportModel = 142
Public Const Res_StellationRatio = 144
Public Const Res_TotalEdgeLength = 146
Public Const Res_TruncationRatio = 148
Public Const Res_Stellation = 150
Public Const Res_Truncation = 152
Public Const Res_Apply = 154
Public Const Res_SphericalImage = 156
Public Const Res_Scanning = 158
Public Const Res_SynchronizeWithPreimage = 160

Public Const ResMnu_File = 200
Public Const ResMnu_New = 202
Public Const ResMnu_Open = 204
Public Const ResMnu_Save = 206
Public Const ResMnu_SaveAs = 208
Public Const ResMnu_Close = 210
Public Const ResMnu_CloseAll = 212
Public Const ResMnu_PrintPreview = 214
Public Const ResMnu_PageSetup = 216
Public Const ResMnu_Print = 218
Public Const ResMnu_Exit = 220
Public Const ResMnu_Window = 222
Public Const ResMnu_Cascade = 224
Public Const ResMnu_TileHoriz = 226
Public Const ResMnu_TileVertic = 228
Public Const ResMnu_Arrange = 230
Public Const ResMnu_Edit = 232
Public Const ResMnu_Options = 234
Public Const ResMnu_View = 236
Public Const ResMnu_Stellate = 238
Public Const ResMnu_Dual = 240
Public Const ResMnu_EdgeDual = 242
Public Const ResMnu_Truncate = 244
Public Const ResMnu_SphericalImage = 246
Public Const ResMnu_Scanning = 248

Public Const ResView_Caption = 300
Public Const ResView_General = 302
Public Const ResView_Backcolor = 304
Public Const ResView_Transparency = 306
Public Const ResView_Points = 308
Public Const ResView_ShowPoints = 310
Public Const ResView_Color = 312
Public Const ResView_Lines = 314
Public Const ResView_ShowLines = 316
Public Const ResView_Polygons = 318
Public Const ResView_ShowPolygons = 320
Public Const ResView_Solid = 322
Public Const ResView_Random = 324
Public Const ResView_Brightness = 326
Public Const ResView_Pastel = 328
Public Const ResView_PointSize = 330

Public Const ResMsg_ContainsNonPlanarFaces = 400
Public Const ResMsg_ChangesMade = 402
Public Const ResMsg_ChangesNotMade = 404
Public Const ResMsg_WishToTriangulate = 406
Public Const ResMsg_SeeHelpForDetails = 408
Public Const ResMsg_UnexistingVertex = 410
Public Const ResMsg_UnableToLoadPolyhedron = 412
Public Const ResMsg_EmptyPolyhedron = 414
Public Const ResMsg_WishToCorrect = 416
Public Const ResMsg_ImpossibleToCalcVolume = 418
Public Const ResMsg_EdgeBelongsToMoreThan2Faces = 420
Public Const ResMsg_ExplorationImpossible = 422
Public Const ResMsg_PolyhedronNotClosed = 424
Public Const ResMsg_PolyhedronNotConnected = 426
Public Const ResMsg_FileAlreadyExists = 428
Public Const ResMsg_PolFile = 430
Public Const ResMsg_Save = 432
Public Const ResMsg_EnterWorksheetName = 434
Public Const ResMsg_DataExportComplete = 436
Public Const ResMsg_ContainsNonConvexEdges = 438
Public Const ResMsg_DualNotDefined = 440
Public Const ResMsg_CannotReconstructEdges = 442
Public Const ResMsg_EdgeDualNotDefined = 444
Public Const ResMsg_CannotTruncatePolyhedron = 446
Public Const ResMsg_CannotStellatePolyhedron = 448
Public Const ResMsg_InputIntegerPointSize = 450
Public Const ResMsg_VertexBelongsTo3SidesOfAFace = 452
Public Const ResMsg_FaceHasLessThan3Vertices = 454


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Resource icon IDs:
Public Const ResIcon_Dodecahedron = 101
Public Const ResIcon_Icosahedron = 102
Public Const ResIcon_Tetrahedron = 103
Public Const ResIcon_Octahedron = 104
Public Const ResIcon_Cube = 105
'#########################################################

Public WindowCount As Long ' the number of open child windows
Public Document() As frmMain ' the array of open windows
Public ActiveWindow As Long 'index of the active window
Public m_SphericalImageVisible As Boolean
Public AuxErrorString As String

'Miscellaneous constants
Public Const EmptyVar = -2 ^ 31 + 3 ' a dummy constant to indicate empty value
Public Const MRUMax = 6


Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long

'###########################################################
'###########################################################
'Procedures named File* are executed when the corresponding menu items are selected

Public Sub FileNew()
'======================================================
'Procedure to be executed when a new file is created
'Creating a new window...
'======================================================
WindowCount = WindowCount + 1
ReDim Preserve Document(1 To WindowCount)
'ReDim Preserve GLData(1 To WindowCount)

Set Document(WindowCount) = New frmMain
Set Document(WindowCount).Icon = GetRandomPolyhedronIcon
Document(WindowCount).Index = WindowCount
Document(WindowCount).Caption = GetString(Res_Untitled) & WindowCount
'Document(WindowCount).WindowState = Document(1).WindowState
Document(WindowCount).Show
ActiveWindow = WindowCount
End Sub

Public Sub FileOpen()
'======================================================
' This happens when an existing file needs to be opened...
'======================================================
LockWindowUpdate frmMainMDI.hWnd

FileNew
Document(ActiveWindow).Polyhedron.OpenAFile

LockWindowUpdate 0
Document(ActiveWindow).Refresh
End Sub

Public Sub FileMRUOpen(ByVal FName As String)
'======================================================
' This takes place when one of the Most Recently Used files is requested
'======================================================
LockWindowUpdate frmMainMDI.hWnd

FileNew
Document(ActiveWindow).Polyhedron.OpenFileNamed FName

LockWindowUpdate 0

Document(ActiveWindow).Refresh
End Sub

'Closing the active window
Public Sub FileClose()
If ActiveWindow >= LBound(Document) And ActiveWindow <= UBound(Document) Then
    If Not (Document(ActiveWindow) Is Nothing) Then Unload Document(ActiveWindow)
End If
End Sub

'Closing all open windows
Public Sub FileCloseAll()
Do While WindowCount > 0
    FileClose
Loop
End Sub

Public Sub FileExit()
FileCloseAll
Unload frmMainMDI
End Sub

'Load string from resource file
Public Function GetString(ByVal ID As Long, Optional ByVal LanguageID As Long = EmptyVar) As String
On Local Error Resume Next
GetString = LoadResString(ID + setLanguage)
If LanguageID <> EmptyVar Then GetString = LoadResString(ID + LanguageID)
End Function

Public Function GetRandomPolyhedronIcon() As IPictureDisp
Set GetRandomPolyhedronIcon = LoadResPicture(ResIcon_Dodecahedron + Int(Rnd * 5), vbResIcon)
End Function

Public Sub ChangeLanguage(ByVal NewLanguage As Language)
Dim Z As Long
setLanguage = NewLanguage
frmMainMDI.FillStrings
For Z = 1 To WindowCount
    Document(Z).FillStrings
Next
SaveSetting AppName, "General", "Language", CStr(setLanguage)
End Sub

Public Sub GetSettings()
setLanguage = GetSetting(AppName, "General", "Language", "0")

setLastExcelPath = GetSetting(AppName, "General", "LastExcelPath", AppPath)
If Not IsValidPath(setLastExcelPath) Then setLastExcelPath = AppPath

setLastPolPath = GetSetting(AppName, "General", "LastPollPath", AppPath)
If Not IsValidPath(setLastPolPath) Then setLastPolPath = AppPath
End Sub

Public Sub SaveSettings()
SaveSetting AppName, "General", "Language", CStr(setLanguage)
SaveSetting AppName, "General", "LastExcelPath", setLastExcelPath
SaveSetting AppName, "General", "LastPolPath", setLastPolPath
End Sub

Public Function AddDirSep(ByVal sStr As String) As String
If sStr <> "" Then If Right(sStr, 1) <> "\" Then sStr = sStr & "\"
AddDirSep = sStr
End Function

Public Function IsValidPath(ByVal FName As String) As Boolean
On Local Error Resume Next
Err.Clear
If Dir(FName, 23) <> "" And FName <> "" Then IsValidPath = True Else IsValidPath = False
If Err.Number <> 0 Then
    Err.Clear
    IsValidPath = False
End If
End Function

Public Function RetrieveDir(ByVal FName As String) As String
Dim Z As Long
On Local Error Resume Next
If FName = "" Then Exit Function
Z = InStrRev(FName, "\")
If Z = 0 Then Exit Function
RetrieveDir = AddDirSep(Left(FName, Z - 1))
End Function

Public Sub Init()
Set CD = New CCommonDialog
Randomize
AppPath = AddDirSep(App.Path)
GetSettings
FillMRU
m_SphericalImageVisible = False
End Sub

Public Sub FillMRU()
Dim Z As Long
Dim tStr As String, ShouldResave As Boolean
On Local Error Resume Next

Do While MRUList.Count > 0
    MRUList.Remove 1
Loop

MRUCount = Val(GetSetting(AppName, "MRU", "Count", "-1"))
If MRUCount <> -1 Then
    For Z = 1 To MRUCount
        tStr = GetSetting(AppName, "MRU", "File" & Z)
        If tStr <> "" And Dir(tStr) <> "" Then
            If Err.Number = 0 Then MRUList.Add tStr
        End If
        Err.Clear
    Next Z
    If MRUList.Count < MRUCount Then ShouldResave = True
    MRUCount = MRUList.Count
End If

If ShouldResave Then SaveMRU Else UpdateMRUMenu
End Sub

Public Sub SaveMRU()
Dim Z As Long, bRemoved As Boolean
On Local Error Resume Next

Z = 0
Do While Z < MRUList.Count
    Z = Z + 1
    bRemoved = False
    If Dir(MRUList(Z)) = "" Then MRUList.Remove Z: Z = Z - 1: bRemoved = True
    If Err.Number = 52 And Not bRemoved Then MRUList.Remove Z: Z = Z - 1: Err.Clear
Loop

SaveSetting AppName, "MRU", "Count", MRUList.Count
If MRUList.Count > 0 Then
    For Z = 1 To MRUList.Count
        SaveSetting AppName, "MRU", "File" & Z, MRUList(Z)
    Next Z
End If
UpdateMRUMenu
End Sub

Public Sub UpdateMRUMenu()
Dim Z As Long, Q As Long
Do While frmMainMDI.mnuMRUFile.UBound > MRUList.Count And frmMainMDI.mnuMRUFile.UBound > 1: Unload frmMainMDI.mnuMRUFile(frmMainMDI.mnuMRUFile.UBound): Loop

For Q = 1 To WindowCount
    Do While Document(Q).mnuMRUFile.UBound > MRUList.Count And Document(Q).mnuMRUFile.UBound > 1: Unload Document(Q).mnuMRUFile(frmMain.mnuMRUFile.UBound): Loop
    
    With Document(Q)
        If MRUList.Count > 0 Then
            .mnuFileSep4.Visible = True
            For Z = 1 To MRUList.Count
                If .mnuMRUFile.UBound < Z Then Load .mnuMRUFile(Z)
                .mnuMRUFile(Z).Caption = "&" & Z & " " & RetrieveName(MRUList(MRUList.Count + 1 - Z))
                .mnuMRUFile(Z).Visible = True
            Next
        Else
            .mnuMRUFile(1).Visible = False
            .mnuFileSep4.Visible = False
        End If
    End With
Next

With frmMainMDI
    If MRUList.Count > 0 Then
        .mnuFileSep4.Visible = True
        For Z = 1 To MRUList.Count
            If .mnuMRUFile.UBound < Z Then Load .mnuMRUFile(Z)
            .mnuMRUFile(Z).Caption = "&" & Z & " " & RetrieveName(MRUList(MRUList.Count + 1 - Z))
            .mnuMRUFile(Z).Visible = True
        Next
    Else
        .mnuMRUFile(1).Visible = False
        .mnuFileSep4.Visible = False
    End If
End With
End Sub

Public Sub AddMRUItem(ByVal tStr As String)
Dim Z As Long
Z = 1
Do While Z <= MRUList.Count
    If LCase(MRUList(Z)) = LCase(tStr) Then MRUList.Remove Z Else Z = Z + 1
Loop
If MRUList.Count >= MRUMax Then MRUList.Remove 1
MRUList.Add tStr
SaveMRU
End Sub

Public Function RetrieveName(ByVal FName As String) As String
If InStr(FName, "\") = 0 Then RetrieveName = FName: Exit Function
RetrieveName = Right(FName, Len(FName) - InStrRev(FName, "\"))
End Function

Public Sub FileMRUClick(ByVal Index As Long)
On Local Error GoTo EH
Dim S As String

S = MRUList(MRUList.Count + 1 - Index)
If Dir(S) = "" Then
    MsgBox GetString(ResMsg_UnableToLoadPolyhedron)
    Exit Sub
End If

FileMRUOpen S
EH:
End Sub

Public Sub FileDual()
Dim C As CPolyhedron

Set C = Document(ActiveWindow).Polyhedron.Dual
If C Is Nothing Then Exit Sub

FileNew
Set Document(ActiveWindow).Polyhedron = C
Document(ActiveWindow).Mode = wmdViewer
Document(ActiveWindow).Refresh
End Sub

Public Sub FileEdgeDual()
Dim C As CPolyhedron

Set C = Document(ActiveWindow).Polyhedron.EdgeDual
If C Is Nothing Then Exit Sub

FileNew
Set Document(ActiveWindow).Polyhedron = C
Document(ActiveWindow).Mode = wmdViewer
Document(ActiveWindow).Refresh
End Sub

Public Property Get SphericalImageVisible() As Boolean
SphericalImageVisible = m_SphericalImageVisible
End Property

Public Property Let SphericalImageVisible(ByVal vNewValue As Boolean)
If m_SphericalImageVisible Xor vNewValue Then
    If vNewValue Then frmSphericalImage.Show Else Unload frmSphericalImage
    m_SphericalImageVisible = vNewValue
End If
End Property

Public Sub Temp()
Dim V As CVector
Set V = New CVector
V.X = 1
V.Y = 2
V.Z = 3
Set V = Nothing
End Sub

Public Sub Main()
'=====================================================
'Debug procedure; starts up when specified in project properties;
'for debug purposes only
'=====================================================
Dim T As Long, Z As Long

T = timeGetTime

For Z = 1 To 100000
    Temp
Next

MsgBox timeGetTime - T

'=================
'Dim V1 As New CVector
'Dim V2 As New CVector
'Dim A As CVector
'Dim T As Long, Z As Long
'
'T = timeGetTime
'
'V1.X = E
'V1.Y = PI
'V1.Z = Sqr(2)
'V2.X = -PI
'V2.Y = 0
'V2.Z = Sqr(3)
'
'For Z = 1 To 100000
'    CVectorProduct V1, V2
'Next
'
'MsgBox timeGetTime - T

'===========================

'Dim V1 As Vector
'Dim V2 As Vector
'Dim A As Vector
'Dim T As Long, Z As Long
'
'T = timeGetTime
'
'V1.X = E
'V1.Y = PI
'V1.Z = Sqr(2)
'V2.X = -PI
'V2.Y = 0
'V2.Z = Sqr(3)
'
'For Z = 1 To 100000
'    A = VectorProduct(V1, V2)
'Next
'
'MsgBox timeGetTime - T
End Sub
