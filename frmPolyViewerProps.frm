VERSION 5.00
Begin VB.Form frmPolyViewerProps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3D graphics display properties"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmPolyViewerProps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame fraPolygonProps 
      Caption         =   "Polygon properties"
      Height          =   975
      Left            =   3240
      TabIndex        =   10
      Top             =   1200
      Width           =   3015
      Begin VB.Frame fraColor 
         Caption         =   "Color"
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton cmdPolygonColor 
            BackColor       =   &H0098D9F8&
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox chkRandomPolygonColor 
            Caption         =   "Random"
            Height          =   255
            Left            =   720
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.HScrollBar hsbBrightness 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   23
            Top             =   840
            Value           =   1
            Width           =   1335
         End
         Begin VB.HScrollBar hsbPastel 
            Height          =   255
            Left            =   1320
            Max             =   255
            Min             =   1
            TabIndex        =   22
            Top             =   1320
            Value           =   1
            Width           =   1335
         End
         Begin VB.Label lblBrightness 
            Caption         =   "Brightness"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblPastel 
            Caption         =   "Pastel colors"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkSingleColorPolygons 
         Caption         =   "Single color polygons"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chkPolygons 
         Caption         =   "Show polygons"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraLineProps 
      Caption         =   "Line properties"
      Height          =   975
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkLines 
         Caption         =   "Show lines"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdLineColor 
         BackColor       =   &H0095A8CC&
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblLineColor 
         Caption         =   "Color"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame fraPointProps 
      Caption         =   "Point properties"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
      Begin VB.TextBox txtPointSize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdPointColor 
         BackColor       =   &H00B5C7FD&
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkPoints 
         Caption         =   "Show points"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblPointSize 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblPointColor 
         Caption         =   "Color"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General appearance"
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar hsbTransparency 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   19
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox chkTransparency 
         Caption         =   "Transparency"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdBackColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblBackColor 
         Caption         =   "Background color"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPolyViewerProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldBackColor As Long
Dim OldTransparency As Long
Dim OldShowPoints As Long
Dim OldPointColor As Long
Dim OldPointSize As Long
Dim OldShowLines As Boolean
Dim OldLineColor As Long
Dim OldShowPolygons As Boolean
Dim OldPolygonColor As Long
Dim OldRandomPolygonColor As Boolean
Dim OldSingleColorPolygons As Boolean
Dim OldBrightness As Long
Dim OldPastel As Long

Public Canvas As ctlPolyViewer
Dim unlCancel As Boolean

Private Sub cmdApply_Click()
Apply
End Sub

Private Sub cmdCancel_Click()
unlCancel = True
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Public Sub FillStrings()
Caption = GetString(ResView_Caption)
cmdCancel.Caption = GetString(Res_Cancel)
cmdApply.Caption = GetString(Res_Apply)

lblBackColor.Caption = GetString(ResView_Backcolor)
lblBrightness = GetString(ResView_Brightness)
lblLineColor = GetString(ResView_Color)
lblPastel = GetString(ResView_Pastel)
lblPointColor = GetString(ResView_Color)
lblPointSize = GetString(ResView_PointSize)

fraGeneral.Caption = GetString(ResView_General)
fraPointProps.Caption = GetString(ResView_Points)
fraLineProps.Caption = GetString(ResView_Lines)
fraPolygonProps.Caption = GetString(ResView_Polygons)
fraColor.Caption = GetString(ResView_Color)

chkTransparency.Caption = GetString(ResView_Transparency)
chkLines.Caption = GetString(ResView_ShowLines)
chkPoints.Caption = GetString(ResView_ShowPoints)
chkPolygons.Caption = GetString(ResView_ShowPolygons)
chkRandomPolygonColor.Caption = GetString(ResView_Random)
chkSingleColorPolygons.Caption = GetString(ResView_Solid)

End Sub

Private Sub chkLines_Click()
cmdLineColor.Enabled = CBool(chkLines.Value)
End Sub

Private Sub chkPoints_Click()
cmdPointColor.Enabled = CBool(chkPoints.Value)
txtPointSize.Enabled = CBool(chkPoints.Value)
End Sub

Private Sub chkPolygons_Click()
cmdPolygonColor.Enabled = CBool(chkPolygons.Value)
End Sub

Private Sub chkTransparency_Click()
hsbTransparency.Enabled = CBool(chkTransparency.Value)
End Sub

Private Sub cmdBackColor_Click()
SetFocus
Enabled = False
CD.Color = cmdBackColor.BackColor
CD.ShowColor
cmdBackColor.BackColor = CD.Color
Enabled = True
SetFocus
End Sub

Private Sub cmdLineColor_Click()
SetFocus
Enabled = False
CD.Color = cmdLineColor.BackColor
CD.ShowColor
cmdLineColor.BackColor = CD.Color
Enabled = True
SetFocus
End Sub

Private Sub cmdPointColor_Click()
SetFocus
Enabled = False
CD.Color = cmdPointColor.BackColor
CD.ShowColor
cmdPointColor.BackColor = CD.Color
Enabled = True
SetFocus
End Sub

Private Sub cmdPolygonColor_Click()
SetFocus
Enabled = False
CD.Color = cmdPolygonColor.BackColor
CD.ShowColor
cmdPolygonColor.BackColor = CD.Color
Enabled = True
SetFocus
End Sub

Private Sub Form_Load()
unlCancel = False
FillStrings
If SphericalImageVisible Then frmSphericalImage.AlwaysOnTop = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Canvas.Enabled = True
If unlCancel Then Exit Sub

If Not CheckCorrectness Then Cancel = 1: Canvas.Enabled = False: Exit Sub
Apply
End Sub

Public Function CheckCorrectness() As Boolean
If Not IsNumber(txtPointSize.Text) Then
    MsgBox GetString(ResMsg_InputIntegerPointSize), vbOKOnly + vbExclamation, GetString(ResView_PointSize)
    CheckCorrectness = False
    Exit Function
End If
If Val(txtPointSize.Text) < 1 Or Val(txtPointSize.Text) > 16 Then
    MsgBox GetString(ResMsg_InputIntegerPointSize), vbOKOnly + vbExclamation, GetString(ResView_PointSize)
    CheckCorrectness = False
    Exit Function
End If

CheckCorrectness = True
End Function

Public Sub Apply()
If Not CheckCorrectness Then Exit Sub

Canvas.BackColor = cmdBackColor.BackColor
If chkTransparency.Value = 1 Then
    Canvas.Transparency = 1 - hsbTransparency.Value / 100
Else
    Canvas.Transparency = 1
End If
Canvas.DrawPoints = chkPoints.Value = 1
Canvas.DrawLines = chkLines.Value = 1
Canvas.DrawFaces = chkPolygons.Value = 1
If IsNumber(txtPointSize.Text) Then
    If Val(txtPointSize.Text) >= 1 And Val(txtPointSize.Text) <= 16 Then Canvas.PointSize = Int(Val(txtPointSize.Text))
End If
Canvas.PointColor = cmdPointColor.BackColor
Canvas.LineColor = cmdLineColor.BackColor
Canvas.PolygonColor = cmdPolygonColor.BackColor
Canvas.SingleColorPolygons = chkSingleColorPolygons.Value = 1
Canvas.RandomPolygonColor = chkRandomPolygonColor.Value = 1
Canvas.RandomColorBrightness = hsbBrightness.Value
Canvas.RandomColorPastel = hsbPastel.Value

Canvas.ReInitColors
Canvas.CompileStructure
Canvas.Refresh
End Sub

Public Sub Fill()
cmdBackColor.BackColor = Canvas.BackColor
chkTransparency.Value = -CInt(Canvas.Transparency <> 1)
hsbTransparency.Value = (1 - Canvas.Transparency) * 100
hsbTransparency.Enabled = chkTransparency.Value <> 0
chkPoints.Value = -CInt(Canvas.DrawPoints)
chkLines.Value = -CInt(Canvas.DrawLines)
chkPolygons.Value = -CInt(Canvas.DrawFaces)
txtPointSize.Text = Canvas.PointSize
txtPointSize.Enabled = Canvas.DrawPoints
cmdPointColor.Enabled = CBool(chkPoints.Value)
cmdLineColor.Enabled = CBool(chkLines.Value)
cmdPolygonColor.Enabled = CBool(chkPolygons.Value)
cmdPointColor.BackColor = Canvas.PointColor
cmdLineColor.BackColor = Canvas.LineColor
cmdPolygonColor.BackColor = Canvas.PolygonColor
chkSingleColorPolygons.Value = -CInt(Canvas.SingleColorPolygons)
chkRandomPolygonColor.Value = -CInt(Canvas.RandomPolygonColor)
hsbBrightness.Value = Canvas.RandomColorBrightness
hsbPastel.Value = Canvas.RandomColorPastel
End Sub
