VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Polyhedron1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainerViewer 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   4455
      ScaleHeight     =   334
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Frame fraContainerViewer 
         Caption         =   "Polyhedron properties"
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmdToEditMode 
            Caption         =   "Edit mode..."
            Height          =   495
            Left            =   600
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblTotalEdgeLength 
            Caption         =   "Total edge length:"
            Height          =   615
            Left            =   120
            TabIndex        =   20
            Top             =   3960
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblSurfaceArea 
            Caption         =   "Surface area:"
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   3360
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Image imgVolume 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":030A
            Top             =   2760
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblVolume 
            Caption         =   "Volume:"
            Height          =   615
            Left            =   720
            TabIndex        =   16
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Line linDark 
            BorderColor     =   &H80000015&
            Index           =   2
            Visible         =   0   'False
            X1              =   120
            X2              =   2040
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line linLight 
            BorderColor     =   &H80000014&
            Index           =   2
            Visible         =   0   'False
            X1              =   120
            X2              =   2040
            Y1              =   2655
            Y2              =   2655
         End
         Begin VB.Label lblEulerianSum 
            Caption         =   "V - E + F = 0"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblFacesCount 
            Caption         =   "Faces: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblEdgesCount 
            Caption         =   "Edges: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblVertexCount 
            Caption         =   "Vertices: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Line linLight 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   120
            X2              =   2040
            Y1              =   975
            Y2              =   975
         End
         Begin VB.Line linDark 
            BorderColor     =   &H80000015&
            Index           =   1
            X1              =   120
            X2              =   2040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblToEditMode 
            Caption         =   "Д"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame fraContainerEditor 
         Caption         =   "Edit polyhedron"
         Height          =   4695
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export"
            Height          =   495
            Left            =   600
            TabIndex        =   18
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton cmdClearModel 
            Caption         =   "Clear model"
            Height          =   495
            Left            =   600
            TabIndex        =   14
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import"
            Height          =   495
            Left            =   600
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdToExplorationMode 
            Caption         =   "Exploration mode..."
            Height          =   495
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblExport 
            Caption         =   "р"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Left            =   180
            TabIndex        =   19
            Top             =   1860
            Width           =   375
         End
         Begin VB.Label lblClearModel 
            AutoSize        =   -1  'True
            Caption         =   "ы"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   20.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   450
            Left            =   240
            TabIndex        =   15
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label lblImport 
            Caption         =   "п"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Left            =   180
            TabIndex        =   9
            Top             =   1260
            Width           =   375
         End
         Begin VB.Line linDark 
            BorderColor     =   &H80000015&
            Index           =   0
            X1              =   120
            X2              =   2040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line linLight 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   120
            X2              =   2040
            Y1              =   975
            Y2              =   975
         End
         Begin VB.Label lblToExplorationMode 
            Caption         =   "З"
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   18
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   180
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox picContainerPicture 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   0
      ScaleHeight     =   334
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   4815
      Begin PolyhedronExplorer.ctlPolyViewer ctlPolyViewer1 
         Height          =   4035
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7117
         PolygonColor    =   12180433
         PointSize       =   5
         Pastel          =   32
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save as "
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close all"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "d"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuDual 
         Caption         =   "Dual"
      End
      Begin VB.Menu mnuEdgeDual 
         Caption         =   "Edge-dual"
      End
      Begin VB.Menu mnuTruncate 
         Caption         =   "Truncate"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStellate 
         Caption         =   "Stellate..."
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSphericalImage 
         Caption         =   "Spherical image"
      End
      Begin VB.Menu mnuScanning 
         Caption         =   "Scanning"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuExploration 
         Caption         =   "Exploration mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditor 
         Caption         =   "Editing mode"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuLanguage 
         Caption         =   "&Language"
         Begin VB.Menu mnuLanEnglish 
            Caption         =   "&English"
         End
         Begin VB.Menu mnuLanRussian 
            Caption         =   "&Russian"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWindowTileHoriz 
         Caption         =   "Tile horizontally"
      End
      Begin VB.Menu mnuWindowTileVertic 
         Caption         =   "Tile vertically"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Arrange icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bIsDirty As Boolean 'was the document saved?
Dim WMode As WindowMode 'current windowmode
Dim OldWindowState As FormWindowStateConstants
Dim NewWindowState As FormWindowStateConstants
Dim NeedToResizeInPaint As Boolean
Dim m_Index As Long

Public Polyhedron As CPolyhedron 'The polyhedron data in current window

Private Sub cmdClearModel_Click()
Polyhedron.Clear
FillEditPane
End Sub

Private Sub cmdExport_Click()
MainSheetName = InputBox(GetString(ResMsg_EnterWorksheetName) & ":", frmMainMDI.Caption)
If MainSheetName <> "" Then
    CD.Filter = GetString(Res_LocateXLS) & " (*." & extXLS & ")|*." & extXLS
    CD.Flags = &H4
    If Not IsValidPath(setLastExcelPath) Then setLastExcelPath = AppPath
    CD.InitDir = setLastExcelPath
    CD.ShowSave
    If CD.Cancelled Then Exit Sub
    
    MainFileName = CD.FileName
    If Right(UCase(MainFileName), 4) <> "." & extXLS Then
        If InStr(RetrieveName(MainFileName), ".") = 0 Then
            MainFileName = MainFileName & "." & extXLS
        Else
            MainFileName = Left(MainFileName, InStr(MainFileName, ".")) + extXLS
        End If
    End If
    setLastExcelPath = AddDirSep(RetrieveDir(MainFileName))
    If Not IsValidPath(setLastExcelPath) Then setLastExcelPath = AppPath
    
    Document(ActiveWindow).Polyhedron.Export
    
End If
End Sub

Private Sub cmdImport_Click()
frmInputMain.Show vbModal
End Sub

Private Sub cmdToEditMode_Click()
Me.Mode = wmdEditor
End Sub

Private Sub cmdToExplorationMode_Click()
Me.Mode = wmdViewer
End Sub

Private Sub ctlPolyViewer1_Rotated(ByVal AngleHorizontal As Double, ByVal AngleVertical As Double, ByVal AngleTraversal As Double)
If SphericalImageVisible Then
    If frmSphericalImage.chkSynchronize.Value = 1 And frmSphericalImage.ParentPolyWindow = Me.Index Then
        frmSphericalImage.ctlPolyViewer1.AngleHoriz = AngleHorizontal
        frmSphericalImage.ctlPolyViewer1.AngleVertic = AngleVertical
        frmSphericalImage.ctlPolyViewer1.AngleTraversal = AngleTraversal
        frmSphericalImage.ctlPolyViewer1.Refresh
    End If
End If
End Sub

Public Sub Form_Activate()
ActiveWindow = Me.Index
ctlPolyViewer1.Activate
End Sub

Private Sub Form_Deactivate()
ctlPolyViewer1.Deactivate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ctlPolyViewer1.KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()
FillStrings
UpdateMRUMenu
bIsDirty = True

Set Polyhedron = New CPolyhedron
Mode = wmdViewer
End Sub

Private Sub Form_Paint()
If Me.WindowState = vbMinimized Then Exit Sub
If NeedToResizeInPaint Then
    DoResizeJob
    NeedToResizeInPaint = False
End If
End Sub

Public Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub

OldWindowState = NewWindowState
NewWindowState = Me.WindowState
If NewWindowState = OldWindowState Then
    DoResizeJob
    NeedToResizeInPaint = False
Else
    NeedToResizeInPaint = True
    If NewWindowState = vbNormal Then
        NeedToResizeInPaint = False
        DoResizeJob
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Z As Long

SphericalImageVisible = False

Me.ctlPolyViewer1.TerminateGL

Set Polyhedron = Nothing

ActiveWindow = Me.Index
If ActiveWindow < WindowCount Then
    Set Document(ActiveWindow) = Nothing
    For Z = ActiveWindow To WindowCount - 1
        Set Document(Z) = Document(Z + 1)
        'GLData(Z) = GLData(Z + 1)
        Document(Z).Index = Z
        If Document(Z).Caption = GetString(Res_Untitled) & Z + 1 Then Document(Z).Caption = GetString(Res_Untitled) & Z
    Next
End If

'Set Document(WindowCount) = Nothing
WindowCount = WindowCount - 1
If WindowCount > 0 Then
    ReDim Preserve Document(1 To WindowCount)
    'ReDim Preserve GLData(1 To WindowCount)
End If
If ActiveWindow > WindowCount Then ActiveWindow = WindowCount
If ActiveWindow > 0 Then Document(ActiveWindow).Form_Activate
End Sub

Private Sub mnuAbout_Click()
MsgBox "Курсовая работа"
End Sub

Private Sub mnuClose_Click()
FileClose
End Sub

Private Sub mnuCloseAll_Click()
FileCloseAll
End Sub

Private Sub mnuDual_Click()
FileDual
End Sub

Private Sub mnuEdgeDual_Click()
FileEdgeDual
End Sub

Private Sub mnuEditor_Click()
Me.Mode = wmdEditor
End Sub

Private Sub mnuExit_Click()
FileExit
End Sub

Private Sub mnuExploration_Click()
Me.Mode = wmdViewer
End Sub

Private Sub mnuLanEnglish_Click()
ChangeLanguage lanEnglish
End Sub

Private Sub mnuLanRussian_Click()
ChangeLanguage lanRussian
End Sub

Private Sub mnuMRUFile_Click(Index As Integer)
FileMRUClick Index
End Sub

Private Sub mnuNew_Click()
FileNew
End Sub

Private Sub mnuOpen_Click()
FileOpen
End Sub

Private Sub mnuOptions_Click()
mnuLanEnglish.Checked = False
mnuLanRussian.Checked = False
Select Case setLanguage
    Case lanEnglish
        mnuLanEnglish.Checked = True
    Case lanRussian
        mnuLanRussian.Checked = True
End Select
End Sub

Private Sub mnuSave_Click()
Document(ActiveWindow).Polyhedron.Save
End Sub

Private Sub mnuSaveAs_Click()
Document(ActiveWindow).Polyhedron.SaveAs
End Sub

Private Sub mnuScanning_Click()
If ActiveWindow = 0 Then ActiveWindow = Index
frmScanning.Show vbModal
End Sub

Private Sub mnuSphericalImage_Click()
If ActiveWindow = 0 Then ActiveWindow = Index
If SphericalImageVisible Then
    frmSphericalImage.Reset
    frmSphericalImage.FillStrings
End If
frmSphericalImage.ParentPolyWindow = Me.Index
frmSphericalImage.Show
End Sub

Private Sub mnuStellate_Click()
frmStellate.Show vbModal
End Sub

Private Sub mnuTruncate_Click()
frmTruncate.Show vbModal
End Sub

Private Sub mnuWindowArrange_Click()
frmMainMDI.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
frmMainMDI.Arrange vbCascade
SendResizeToAllChildWindows
End Sub

Private Sub mnuWindowTileHoriz_Click()
frmMainMDI.Arrange vbTileHorizontal
SendResizeToAllChildWindows
End Sub

Private Sub mnuWindowTileVertic_Click()
frmMainMDI.Arrange vbTileVertical
SendResizeToAllChildWindows
End Sub

Public Property Get IsDirty() As Boolean
IsDirty = bIsDirty
End Property

Public Property Let IsDirty(ByVal vNewValue As Boolean)
bIsDirty = vNewValue
End Property

Public Property Get Mode() As WindowMode
Mode = WMode
End Property

Public Property Let Mode(ByVal vNewValue As WindowMode)
WMode = vNewValue
If WMode = wmdEditor Then
    
    fraContainerEditor.Visible = True
    fraContainerViewer.Visible = False
    ctlPolyViewer1.Mode = wmdEditor
    ctlPolyViewer1.Clear
    mnuEditor.Checked = True
    mnuExploration.Checked = False
    mnuEdit.Visible = False
    
    FillEditPane
    
Else
    fraContainerEditor.Visible = False
    fraContainerViewer.Visible = True
    ctlPolyViewer1.Mode = wmdViewer
    mnuExploration.Checked = True
    mnuEditor.Checked = False
    
    If Not Me.Polyhedron Is Nothing Then
        If Not Me.Polyhedron.IsEmptyPolyhedron Then
            LoadVisualModel
            mnuEdit.Visible = True
        End If
    End If
    
    FillBasicInfo
    
End If
End Property

Public Sub DoResizeJob()
Dim SW As Long, SH As Long
Const MinContainerWidth = 160
Const MinViewerHeight = 32

SW = ScaleWidth
SH = ScaleHeight

SW = SW - picContainerViewer.Width
If SW > MinViewerHeight Then picContainerPicture.Width = SW - 4

With picContainerPicture
    SW = .ScaleWidth
    SH = .ScaleHeight
    
    If SH >= SW Then
        If SW > MinViewerHeight Then ctlPolyViewer1.Move 0, (SH - SW) \ 2, SW - 1, SW
    Else
        If SH > MinViewerHeight Then ctlPolyViewer1.Move (SW - SH) \ 2, 0, SH, SH - 1
    End If
End With

With fraContainerViewer
    If picContainerViewer.ScaleHeight > 2 * .Top And picContainerViewer.ScaleWidth > 2 * .Left Then
        .Height = picContainerViewer.ScaleHeight - 2 * .Top
    End If
    fraContainerEditor.Move .Left, .Top, .Width, .Height
End With

End Sub

Public Sub FillBasicInfo()
If Me.Polyhedron Is Nothing Then Exit Sub
With Me.Polyhedron
    If .IsEmptyPolyhedron Then
        lblVertexCount.Visible = False
        lblEdgesCount.Visible = False
        lblFacesCount.Visible = False
        lblEulerianSum.Visible = False
        linDark(2).Visible = False
        linLight(2).Visible = False
        imgVolume.Visible = False
        lblVolume.Visible = False
        lblSurfaceArea.Visible = False
        lblTotalEdgeLength.Visible = False
    Else
        lblVertexCount.Caption = GetString(Res_Vertices) & ": " & .Vertices.Count
        lblEdgesCount.Caption = GetString(Res_Edges) & ": " & .Edges.Count
        lblFacesCount.Caption = GetString(Res_Faces) & ": " & .Facets.Count
        lblEulerianSum.Caption = GetString(Res_EulerianSum) & ": " & .EulerCharacteristic ' .Vertices.Count - .Edges.Count + .Facets.Count
        lblVertexCount.Visible = True
        lblEdgesCount.Visible = True
        lblFacesCount.Visible = True
        lblEulerianSum.Visible = True
        linDark(2).Visible = True
        linLight(2).Visible = True
        imgVolume.Visible = True
        lblVolume.Visible = True
        lblSurfaceArea.Visible = True
        lblTotalEdgeLength.Visible = True
        
        If Document(ActiveWindow).Polyhedron.Oriented Then
            'Dim T As Long: T = timeGetTime
            lblVolume.Caption = GetString(Res_Volume) & ": " & Format(Document(ActiveWindow).Polyhedron.Volume, "# ##0.####")
            'MsgBox timeGetTime - T
        Else
            lblVolume.Caption = GetString(ResMsg_ImpossibleToCalcVolume)
        End If
        
        lblSurfaceArea.Caption = GetString(Res_SurfaceArea) & ": " & Format(Document(ActiveWindow).Polyhedron.SurfaceArea, "# ##0.####")
        lblTotalEdgeLength.Caption = GetString(Res_TotalEdgeLength) & ": " & Format(Document(ActiveWindow).Polyhedron.Edges.TotalLength, "# ##0.####")
        
    End If
End With
End Sub

Public Sub LoadVisualModel()
Me.Polyhedron.CreateVisualModel ctlPolyViewer1
End Sub

Public Sub FillEditPane()
If Me.Polyhedron Is Nothing Then
    lblClearModel.Visible = False
    cmdClearModel.Visible = False
    lblExport.Visible = False
    cmdExport.Visible = False
    Exit Sub
End If
With Me.Polyhedron
    If .IsEmptyPolyhedron Then
        lblClearModel.Visible = False
        cmdClearModel.Visible = False
        lblExport.Visible = False
        cmdExport.Visible = False
    Else
        lblClearModel.Visible = True
        cmdClearModel.Visible = True
        lblExport.Visible = True
        cmdExport.Visible = True
    End If
End With
End Sub

Public Sub FillStrings()
'Load appropriate string names for interface controls according to current language
mnuClose.Caption = GetString(ResMnu_Close)
mnuCloseAll.Caption = GetString(ResMnu_CloseAll)
mnuExit.Caption = GetString(ResMnu_Exit)
mnuFile.Caption = GetString(ResMnu_File)
mnuNew.Caption = GetString(ResMnu_New)
mnuOpen.Caption = GetString(ResMnu_Open)
mnuSave.Caption = GetString(ResMnu_Save)
mnuSaveAs.Caption = GetString(ResMnu_SaveAs)
mnuWindow.Caption = GetString(ResMnu_Window)
mnuWindowCascade.Caption = GetString(ResMnu_Cascade)
mnuWindowTileHoriz.Caption = GetString(ResMnu_TileHoriz)
mnuWindowTileVertic.Caption = GetString(ResMnu_TileVertic)
mnuWindowArrange.Caption = GetString(ResMnu_Arrange)
mnuEdit.Caption = GetString(ResMnu_Edit)
mnuStellate.Caption = GetString(ResMnu_Stellate)
mnuDual.Caption = GetString(ResMnu_Dual)
mnuEdgeDual.Caption = GetString(ResMnu_EdgeDual)
mnuTruncate.Caption = GetString(ResMnu_Truncate)
mnuSphericalImage.Caption = GetString(ResMnu_SphericalImage)
mnuScanning.Caption = GetString(ResMnu_Scanning)

mnuView.Caption = GetString(ResMnu_View)
mnuEditor.Caption = GetString(Res_EditingMode)
mnuExploration.Caption = GetString(Res_ExplorationMode)
fraContainerViewer.Caption = GetString(Res_PolyhedronProperties)
fraContainerEditor.Caption = GetString(Res_EditingMode)
cmdToEditMode.Caption = GetString(Res_EditingMode)
cmdToExplorationMode.Caption = GetString(Res_ExplorationMode)
cmdImport.Caption = GetString(Res_ImportModel)
cmdExport.Caption = GetString(Res_ExportModel)
mnuOptions.Caption = GetString(ResMnu_Options)
cmdClearModel.Caption = GetString(Res_ClearModel)

FillBasicInfo

End Sub

Public Property Get Index() As Long
Index = m_Index
End Property

Public Property Let Index(ByVal vNewValue As Long)
m_Index = vNewValue
End Property

Public Sub SendResizeToAllChildWindows()
Dim Z As Long

For Z = 1 To WindowCount
    Document(Z).ctlPolyViewer1.ShouldResize
Next
End Sub
