VERSION 5.00
Begin VB.Form frmStellate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stellate polyhedron"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmStellate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hsbRatio 
      Height          =   375
      LargeChange     =   20
      Left            =   120
      Max             =   200
      Min             =   -100
      TabIndex        =   2
      Top             =   5520
      Width           =   4860
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4005
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin PolyhedronExplorer.ctlPolyViewer ctlPolyViewer1 
      Height          =   4860
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   8573
      PolygonColor    =   16761024
   End
   Begin VB.Label lblRatio 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   4815
   End
End
Attribute VB_Name = "frmStellate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxPercentage = 2
Public Percentage As Double

Dim unlCancel As Boolean
Dim OldVCount As Long

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

OldVCount = Document(ActiveWindow).Polyhedron.Vertices.Count
Percentage = 1
Set Poly = Document(ActiveWindow).Polyhedron.Stellate(Percentage)
If Poly Is Nothing Then
    ctlPolyViewer1.Enabled = False
    cmdOK.Enabled = False
    hsbRatio.Enabled = False
    unlCancel = True
    Exit Sub
Else
    Poly.CreateVisualModel ctlPolyViewer1
    hsbRatio.Value = Percentage * 100 / MaxPercentage
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim W As Long

ctlPolyViewer1.SelfDestruct
If unlCancel Then Exit Sub

Me.Hide
W = ActiveWindow
FileNew
Set Document(ActiveWindow).Polyhedron = Document(W).Polyhedron.Stellate(Percentage)
Document(ActiveWindow).Mode = wmdViewer
Document(ActiveWindow).Refresh
End Sub

Private Sub hsbRatio_Change()
Percentage = hsbRatio.Value / 100 * MaxPercentage
Update
End Sub

Private Sub hsbRatio_Scroll()
hsbRatio_Change
End Sub

Public Sub Update()
lblRatio.Caption = GetString(Res_StellationRatio) & ": " & Format(Percentage, "0.0##")
lblRatio.Refresh
RecalcStellatedPolyhedron
ctlPolyViewer1.CompileStructure
ctlPolyViewer1.Redraw
End Sub

Public Sub RecalcStellatedPolyhedron()
Dim F As CFacet, NF As CFacet
Dim Z As Long, Q As Long

Dim CG As CVector
Dim tSummit As CVector
Dim tV1 As CVector
Dim tV2 As CVector

Q = ctlPolyViewer1.DataIndex
With Document(ActiveWindow).Polyhedron
    Z = 0
    For Each F In .Facets
        Set tSummit = F.Normal
        tSummit.ScaleVector Percentage
        tSummit.Add F.GetCenterOfGravity
'        Set CG = F.CenterOfGravity
'        Set tV1 = .Vertices(F.Vertices(1)).DirectingVector ' requires delete
'        Set tV2 = .Vertices(F.Vertices(2)).DirectingVector ' requires delete
'        tV1.Subtract CG
'        tV2.Subtract CG
'        Set tSummit = CVectorProduct(tV1, tV2)
'        tSummit.ScaleVector Percentage
'        tSummit.Add CG
        
        Z = Z + 1
        
        GLData(Q).Points3D(OldVCount + Z).X = tSummit.X
        GLData(Q).Points3D(OldVCount + Z).Y = tSummit.Y
        GLData(Q).Points3D(OldVCount + Z).Z = tSummit.Z
        
        Set tSummit = Nothing
'        Set tV1 = Nothing
'        Set tV2 = Nothing
    Next
End With
End Sub

Public Sub FillStrings()
lblRatio.Caption = GetString(Res_StellationRatio) & ": 1"
cmdCancel.Caption = GetString(Res_Cancel)
Caption = GetString(Res_Stellation)
End Sub
