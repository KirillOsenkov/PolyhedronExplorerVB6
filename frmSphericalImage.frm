VERSION 5.00
Begin VB.Form frmSphericalImage 
   Caption         =   "Spherical image"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "frmSphericalImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin PolyhedronExplorer.ctlPolyViewer ctlPolyViewer1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8705
      DrawPoints      =   -1  'True
      PointSize       =   12
      LineWidth       =   2
   End
   Begin VB.PictureBox picOnTop 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   120
      Picture         =   "frmSphericalImage.frx":0442
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox picOnTop 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "frmSphericalImage.frx":059C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkSynchronize 
      Caption         =   "Synchronize with preimage"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   5160
      Value           =   1  'Checked
      Width           =   3375
   End
End
Attribute VB_Name = "frmSphericalImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "User32" (ByVal H&, ByVal hb&, ByVal X&, ByVal Y&, ByVal CX&, ByVal CY&, ByVal F&) As Long
Private m_AlwaysOnTop As Boolean

Public ParentPolyWindow As Long

Private Sub chkSynchronize_Click()
ctlPolyViewer1_Rotated ctlPolyViewer1.AngleHoriz, ctlPolyViewer1.AngleVertic, ctlPolyViewer1.AngleTraversal
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub ctlPolyViewer1_AfterInitGL()
glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
glEnable GL_BLEND
End Sub

Private Sub ctlPolyViewer1_AfterRepaint()

Const Prec = 100, Radius = 1.01
Const SphereRed = 0.7
Const SphereGreen = 0.95
Const SphereBlue = 0.95
Const SphereTransparent = 0.9

'glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
glEnable GL_BLEND
'glEnable GL_POINT_SMOOTH
glEnable GL_LINE_SMOOTH

Dim Q As Long, Phi As Double

glBegin GL_POLYGON
    glColor4f SphereRed, SphereGreen, SphereBlue, SphereTransparent
    
    For Q = 0 To Prec
        Phi = Q * 2 * PI / Prec
        glVertex3f Radius * Cos(Phi), Radius * Sin(Phi), 0
    Next
glEnd
End Sub

Private Sub ctlPolyViewer1_Rotated(ByVal AngleHorizontal As Double, ByVal AngleVertical As Double, ByVal AngleTraversal As Double)
If chkSynchronize.Value = 1 Then
    Document(ParentPolyWindow).ctlPolyViewer1.AngleHoriz = AngleHorizontal
    Document(ParentPolyWindow).ctlPolyViewer1.AngleVertic = AngleVertical
    Document(ParentPolyWindow).ctlPolyViewer1.AngleTraversal = AngleTraversal
    Document(ParentPolyWindow).ctlPolyViewer1.Refresh
End If
End Sub

Private Sub Form_Activate()
ctlPolyViewer1.Activate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdOK_Click
End Sub

Private Sub Form_Load()
FillStrings

'Reset
ctlPolyViewer1.AngleHoriz = Document(ActiveWindow).ctlPolyViewer1.AngleHoriz
ctlPolyViewer1.AngleTraversal = Document(ActiveWindow).ctlPolyViewer1.AngleTraversal
ctlPolyViewer1.AngleVertic = Document(ActiveWindow).ctlPolyViewer1.AngleVertic
Document(ActiveWindow).Polyhedron.OutputSphericalImage ctlPolyViewer1

Visible = True

m_AlwaysOnTop = False
AlwaysOnTop = True
m_SphericalImageVisible = True
End Sub

Private Sub Form_Resize()
Dim nX As Long, nY As Long
nX = ScaleWidth - 2 * ctlPolyViewer1.Top
nY = ScaleHeight - 3 * ctlPolyViewer1.Top - cmdOK.Height
If nX > nY Then nX = nY
If nY > nX Then nY = nX
If nX < 4 Or nY < 4 Then Exit Sub
ctlPolyViewer1.Move (ScaleWidth - nX) \ 2, ctlPolyViewer1.Top, nX, nY
cmdOK.Move ctlPolyViewer1.Left + ctlPolyViewer1.Width - cmdOK.Width, 2 * ctlPolyViewer1.Top + ctlPolyViewer1.Height
picOnTop(0).Move ctlPolyViewer1.Left, cmdOK.Top + (cmdOK.Height - picOnTop(0).Height) \ 2 + 1
picOnTop(1).Move picOnTop(0).Left, picOnTop(0).Top
chkSynchronize.Move picOnTop(0).Left + picOnTop(0).Width + 2 * ctlPolyViewer1.Top, cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
ctlPolyViewer1.SelfDestruct
m_SphericalImageVisible = False
End Sub

Public Sub FillStrings()
Caption = GetString(Res_SphericalImage) & " - " & Document(ActiveWindow).Caption
chkSynchronize.Caption = GetString(Res_SynchronizeWithPreimage)
End Sub

Public Sub Reset()
ctlPolyViewer1.Clear
Document(ActiveWindow).Polyhedron.OutputSphericalImage ctlPolyViewer1
ctlPolyViewer1.AngleHoriz = Document(ActiveWindow).ctlPolyViewer1.AngleHoriz
ctlPolyViewer1.AngleTraversal = Document(ActiveWindow).ctlPolyViewer1.AngleTraversal
ctlPolyViewer1.AngleVertic = Document(ActiveWindow).ctlPolyViewer1.AngleVertic
ctlPolyViewer1.Refresh
End Sub

Private Sub picOnTop_Click(Index As Integer)
AlwaysOnTop = Index = 0
End Sub

Public Property Get AlwaysOnTop() As Boolean
AlwaysOnTop = m_AlwaysOnTop
End Property

Public Property Let AlwaysOnTop(ByVal vNewValue As Boolean)
'=====================================================
'hwnd_topmost=-1
'hwnd_notopmost=-2
'topwindow=3
'=====================================================
SetWindowPos hWnd, -CLng(vNewValue) - 2, 0, 0, 0, 0, 3
m_AlwaysOnTop = vNewValue
picOnTop(1).Visible = vNewValue
picOnTop(0).Visible = Not vNewValue
End Property
