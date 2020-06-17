VERSION 5.00
Begin VB.Form frmTruncate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Truncate"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmTruncate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4005
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin VB.HScrollBar hsbRatio 
      Height          =   375
      LargeChange     =   5
      Left            =   120
      Max             =   100
      TabIndex        =   0
      Top             =   5520
      Width           =   4860
   End
   Begin PolyhedronExplorer.ctlPolyViewer ctlPolyViewer1 
      Height          =   4860
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   8573
      PolygonColor    =   16761024
   End
   Begin VB.Label lblRatio 
      Caption         =   "Truncation ratio: 67%"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   4815
   End
End
Attribute VB_Name = "frmTruncate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Percentage As Double
Dim unlCancel As Boolean

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
Dim Poly As CPolyhedron
unlCancel = False

FillStrings

Percentage = 2 / 3

Set Poly = Document(ActiveWindow).Polyhedron.Truncate(Percentage)

If Poly Is Nothing Then
    unlCancel = True
    ctlPolyViewer1.Enabled = False
    cmdOK.Enabled = False
    hsbRatio.Enabled = False
    Exit Sub
Else
    Poly.CreateVisualModel ctlPolyViewer1
    hsbRatio.Value = Percentage * 100
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim W As Long

ctlPolyViewer1.TerminateGL
If unlCancel Then Exit Sub

Select Case Percentage
Case 0
    'do nothing
Case 1
    Me.Hide
    FileEdgeDual
Case Else
    Me.Hide
    W = ActiveWindow
    FileNew
    Set Document(ActiveWindow).Polyhedron = Document(W).Polyhedron.Truncate(Percentage)
    Document(ActiveWindow).Mode = wmdViewer
    Document(ActiveWindow).Refresh
End Select
End Sub

Private Sub hsbRatio_Change()
Percentage = hsbRatio.Value / 100
Update
End Sub

Private Sub hsbRatio_Scroll()
Percentage = hsbRatio.Value / 100
Update
End Sub

Public Sub Update()
lblRatio.Caption = GetString(Res_TruncationRatio) & ": " & Format(Percentage, "##0%")
lblRatio.Refresh
RecalcTruncatedPolyhedron
ctlPolyViewer1.CompileStructure
ctlPolyViewer1.Redraw
End Sub

Public Sub RecalcTruncatedPolyhedron()
Dim Z As Long, Q As Long
Dim EMP As CVector
'
'Q = ctlPolyViewer1.DataIndex
'With Document(ActiveWindow).Polyhedron
'For Z = 1 To .Edges.Count
'    EMP = EdgeSubdivision(.Vertices, .Edges(Z).StartPoint, .Edges(Z).EndPoint, Percentage / 2)
'    GLData(Q).Points3D(2 * Z - 1).X = EMP.X
'    GLData(Q).Points3D(2 * Z - 1).Y = EMP.Y
'    GLData(Q).Points3D(2 * Z - 1).Z = EMP.Z
'    EMP = EdgeSubdivision(.Vertices, .Edges(Z).StartPoint, .Edges(Z).EndPoint, 1 - Percentage / 2)
'    GLData(Q).Points3D(2 * Z).X = EMP.X
'    GLData(Q).Points3D(2 * Z).Y = EMP.Y
'    GLData(Q).Points3D(2 * Z).Z = EMP.Z
'Next
'End With
End Sub

Public Sub FillStrings()
lblRatio.Caption = GetString(Res_TruncationRatio) & ": 67%"
cmdCancel.Caption = GetString(Res_Cancel)
Caption = GetString(Res_Truncation)
End Sub

