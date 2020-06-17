VERSION 5.00
Begin VB.Form frmInputMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import data"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmInputMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame fraSelectSheet 
      Caption         =   "2. Select Excel worksheet"
      Enabled         =   0   'False
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6255
      Begin VB.ListBox lstSheets 
         Height          =   2400
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkSheetPreview 
         Caption         =   "Sheet preview"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblSheetList 
         Caption         =   "Sheets:"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image imgExcel2 
         Height          =   480
         Left            =   240
         Picture         =   "frmInputMain.frx":030A
         Top             =   600
         Width           =   480
      End
      Begin VB.OLE oleSheetPreview 
         Height          =   2415
         Left            =   2640
         OLETypeAllowed  =   0  'Linked
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame fraSelectFile 
      Caption         =   "1. Select Excel file"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdLocateFile 
         Caption         =   "Locate XLS file"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   5055
      End
      Begin VB.Image imgExcel1 
         Height          =   480
         Left            =   240
         Picture         =   "frmInputMain.frx":0614
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblFilename 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmInputMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim unlCancel As Boolean 'the dialog box was cancelled

'Instances of required Excel objects:
Dim XL As Excel.Application 'this represents an entire application
Dim WB As Excel.Workbook 'this is the workbook file

Private Sub chkSheetPreview_Click()
If chkSheetPreview.Value = 1 Then
    lstSheets_Click
Else
    oleSheetPreview.Delete
End If
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdLocateFile_Click()
On Local Error GoTo EH:

If Not WB Is Nothing Then
    WB.Close False
    Set WB = Nothing
End If

CD.Filter = "Microsoft Excel (*.XLS)|*.XLS"
CD.Flags = &H1000 + &H4
CD.InitDir = setLastExcelPath
CD.ShowOpen
If CD.Cancelled = True Or Dir(CD.FileName) = "" Then
    fraSelectSheet.Enabled = False
    cmdOK.Enabled = False
    chkSheetPreview.Value = 0
    Exit Sub
End If

MainFileName = CD.FileName
lblFilename = MainFileName
If IsValidPath(RetrieveDir(MainFileName)) Then setLastExcelPath = AddDirSep(RetrieveDir(MainFileName))
fraSelectSheet.Enabled = True
cmdOK.Enabled = True

Set WB = XL.Workbooks.Open(MainFileName)

FillSheetList

EH:
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then unlCancel = True: Unload Me
If KeyCode = vbKeyReturn Then Unload Me
End Sub

Private Sub Form_Load()
Set XL = New Excel.Application

FillStrings
MainFileName = ""
MainSheetName = ""
unlCancel = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then unlCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo EH

If unlCancel Then
    Set WB = Nothing
    Set XL = Nothing
    Exit Sub
End If

ImportFromExcel

If Not WB Is Nothing Then
    WB.Close False
End If
Set WB = Nothing

If Not XL Is Nothing Then
    XL.Quit
End If
Set XL = Nothing

EH:
End Sub

Public Sub FillStrings()
Caption = GetString(Res_ImportData)
cmdCancel.Caption = GetString(Res_Cancel)
fraSelectFile.Caption = "1. " & GetString(Res_LocateXLS)
cmdLocateFile.Caption = GetString(Res_BrowseForXLS)
fraSelectSheet.Caption = "2. " & GetString(Res_SelectXLWorksheet)
lblSheetList.Caption = GetString(Res_Worksheets)
chkSheetPreview.Caption = GetString(Res_WorksheetPreview)
End Sub

Public Sub FillSheetList()
Dim F As Excel.Worksheet
On Local Error GoTo EH:

If MainFileName = "" Then Exit Sub
If Dir(MainFileName) = "" Then Exit Sub

lstSheets.Clear
For Each F In WB.Worksheets
    lstSheets.AddItem F.Name
Next

lstSheets.ListIndex = 0
lstSheets_Click

Exit Sub
EH:
End Sub

Private Sub lstSheets_Click()
On Local Error GoTo EH
MainSheetName = lstSheets.List(lstSheets.ListIndex)
If chkSheetPreview.Value = 1 Then
    oleSheetPreview.CreateLink MainFileName, MainSheetName & "!R1C1:R10C4"
    'oleSheetPreview.Update
End If
Exit Sub

EH:
End Sub

Public Sub ImportFromExcel(Optional ByVal FName As String = "", Optional ByVal SheetName As String = "")
On Local Error GoTo EH:

Dim sX As Range, sY As Range, sZ As Range
Dim tX() As Double, tY() As Double, tZ() As Double
Dim colFaceVertices() As Collection
Dim Z As Long, Q As Long, N As Long
Dim Unsuccessful As Boolean
Dim VertexCount As Long
Dim WS As Excel.Worksheet 'this is a sheet in a file

Dim tempFace As CFacet

If FName = "" Then FName = MainFileName
If SheetName = "" Then SheetName = MainSheetName
If Dir(FName) = "" Then
    'ERROR
    Exit Sub
End If

Set WS = WB.Worksheets(SheetName)

With WS
    Z = 0
    Set sX = .Cells(1, 1)
    Set sY = .Cells(1, 2)
    Set sZ = .Cells(1, 3)
    Do While IsNumber(sX.Text) And IsNumber(sY.Text) And IsNumber(sZ.Text)
        Z = Z + 1
        ReDim Preserve tX(1 To Z)
        ReDim Preserve tY(1 To Z)
        ReDim Preserve tZ(1 To Z)
        tX(Z) = Val(sX.Text)
        tY(Z) = Val(sY.Text)
        tZ(Z) = Val(sZ.Text)
        Set sX = .Cells(Z + 1, 1)
        Set sY = .Cells(Z + 1, 2)
        Set sZ = .Cells(Z + 1, 3)
    Loop
    
    VertexCount = Z
    
    Z = 0
    Set sX = .Cells(1, 4)
    Do While IsNumber(sX.Text)
        Q = 1
        Z = Z + 1
        ReDim Preserve colFaceVertices(1 To Z)
        Set colFaceVertices(Z) = New Collection
        colFaceVertices(Z).Add Val(sX.Text)
        
        Set sY = .Cells(Z, 4 + Q)
        Do While IsNumber(sY.Text)
            Q = Q + 1
            colFaceVertices(Z).Add Val(sY.Text)
            Set sY = .Cells(Z, 4 + Q)
        Loop
        
        Set sX = .Cells(Z + 1, 4)
    Loop
End With


Document(ActiveWindow).Polyhedron.Clear

For Q = 1 To VertexCount
    Document(ActiveWindow).Polyhedron.Vertices.Add tX(Q), tY(Q), tZ(Q), Q
Next
For Q = 1 To Z
    Set tempFace = Document(ActiveWindow).Polyhedron.Facets.Add
    For N = 1 To colFaceVertices(Q).Count
        tempFace.Vertices.Add colFaceVertices(Q)(N)
    Next
    'tempFace.Index = Q
    Set colFaceVertices(Q) = Nothing
    'If Not tempFace.Consistent Then
    '    Unsuccessful = True
    'End If
    Set tempFace = Nothing
Next

Document(ActiveWindow).Polyhedron.SelfTest

Set WS = Nothing

If Unsuccessful Then
    Document(ActiveWindow).Polyhedron.Clear
    MsgBox GetString(ResMsg_UnexistingVertex), vbOKOnly Or vbExclamation, GetString(ResMsg_UnableToLoadPolyhedron)
End If

Document(ActiveWindow).FillEditPane
Exit Sub

EH:
Document(ActiveWindow).Polyhedron.Clear
MsgBox GetString(ResMsg_UnexistingVertex), vbOKOnly Or vbExclamation, GetString(ResMsg_UnableToLoadPolyhedron)
End Sub
