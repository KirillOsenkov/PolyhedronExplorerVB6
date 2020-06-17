VERSION 5.00
Begin VB.Form frmScanning 
   Caption         =   "Scanning"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmScanning.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.HScrollBar hsbAnimStep 
      Height          =   375
      LargeChange     =   5
      Left            =   1200
      Max             =   1000
      TabIndex        =   3
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton cmdAnimation 
      Caption         =   "Анимация"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin PolyhedronExplorer.ctlPolyViewer ctlPolyViewer1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8493
   End
End
Attribute VB_Name = "frmScanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ScanningAnimating As Boolean

Private Sub cmdAnimation_Click()
Dim T As Double
hsbAnimStep.Enabled = False
cmdAnimation.Enabled = False
ScanningAnimating = True
hsbAnimStep.Value = hsbAnimStep.Max
For T = 0 To 1 Step 0.001
    If ScanningAnimating Then ctlPolyViewer1.OutputScanningModel Document(ActiveWindow).Polyhedron, T Else Exit Sub
    'If T * 1000 Mod 40 = 0 Then hsbAnimStep.Value = T * hsbAnimStep.Max
    ctlPolyViewer1.Refresh
    DoEvents
Next
ScanningAnimating = False
hsbAnimStep.Enabled = True
cmdAnimation.Enabled = True
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdOK_Click
End Sub

Private Sub Form_Load()
ScanningAnimating = False
FillStrings
ctlPolyViewer1.OutputScanningModel Document(ActiveWindow).Polyhedron, 0
End Sub

Private Sub Form_Resize()
Dim nX As Long, nY As Long

nX = ScaleWidth - 2 * ctlPolyViewer1.Top
nY = ScaleHeight - 3 * ctlPolyViewer1.Top - cmdOK.Height
If nX > nY Then nX = nY
If nY > nX Then nY = nX
If nX < 4 Or nY < 4 Then Exit Sub

ctlPolyViewer1.Move (ScaleWidth - nX) \ 2, ctlPolyViewer1.Top, nX, nY
cmdOK.Move ctlPolyViewer1.Left + ctlPolyViewer1.Width - cmdOK.Width, ctlPolyViewer1.Top + ctlPolyViewer1.Height + ctlPolyViewer1.Top
cmdAnimation.Move ctlPolyViewer1.Left, cmdOK.Top
hsbAnimStep.Move cmdAnimation.Left + cmdAnimation.Width + ctlPolyViewer1.Top, cmdAnimation.Top
If cmdOK.Left - hsbAnimStep.Left > ctlPolyViewer1.Top Then
    hsbAnimStep.Width = cmdOK.Left - hsbAnimStep.Left - ctlPolyViewer1.Top
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ScanningAnimating = False
ctlPolyViewer1.TerminateGL
End Sub

Public Sub FillStrings()
Caption = GetString(Res_Scanning) & " - " & Document(ActiveWindow).Caption
End Sub

Private Sub hsbAnimStep_Change()
ctlPolyViewer1.OutputScanningModel Document(ActiveWindow).Polyhedron, hsbAnimStep.Value / hsbAnimStep.Max
End Sub

Private Sub hsbAnimStep_Scroll()
hsbAnimStep_Change
End Sub
