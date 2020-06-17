VERSION 5.00
Begin VB.UserControl ctlPolyViewer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlPolyViewer.ctx":0000
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "ctlPolyViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'Default Property Values:
Const m_def_ForeColor = vbBlack
Const m_def_Perspective = 0
Const m_def_BorderBox = 0
Const m_def_BackColor = vbWhite
Const m_def_Axes = 0
Const m_def_DrawLines = True
Const m_def_DrawFaces = True
Const m_def_DrawPoints = False
Const m_def_PolygonColor = &HC0C0C0
Const m_def_PointColor = 0
Const m_def_PointSize = 3
Const m_def_Pastel = 48
Const m_def_Brightness = 128
Const m_def_SingleColorPolygons = True
Const m_def_RandomPolygonColor = True
Const BoxSizeEpsilon = 0.1
'Property Variables:
Dim m_ForeColor As Long
Dim m_Perspective As Boolean
Dim m_BorderBox As Boolean
Dim m_Axes As Boolean
Dim m_DataIndex As Long
Dim m_Transparency As Double
Dim m_DrawLines As Boolean
Dim m_DrawFaces As Boolean
Dim m_DrawPoints As Boolean
Dim m_BoxSize As Double
Dim m_SingleColorPolygons As Boolean
Dim m_RandomPolygonColor As Boolean
Dim m_PointColor As Long
Dim m_LineColor As Long
Dim m_PolygonColor As Long
Dim m_PointSize As Long
Dim m_LineWidth As Long
Dim m_Pastel As Single
Dim m_Brightness As Single

Dim WMode As WindowMode
'Dim PolygonColor As Long
'Dim RandomPolygonColor As Boolean
'Public ListID As Long
'Public ListLinesID As Long
Dim WasSelfDestroyed As Boolean
Public CenterOffset As Double
Public GLWasInit As Boolean

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public AngleHoriz As Double
Public AngleVertic As Double
Public AngleTraversal As Double
Public ScaleFactor As Double

Private OldAngleHoriz As Double
Private OldAngleVertic As Double
Private OldScaleFactor As Double

Private OX As Long, OY As Long
Private Dragging As Boolean

Private rOX As Long, rOY As Long
Private rDragging As Boolean

Private theOldDC As Long, theOldRC As Long, DCNeedsBackup As Boolean, DCMadeBackup As Boolean

Public Event BeforeRepaint()
Public Event AfterRepaint()
Public Event AfterInitGL()
Public Event Rotated(ByVal AngleHorizontal As Double, ByVal AngleVertical As Double, ByVal AngleTraversal As Double)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
UserControl.BackColor() = New_BackColor

If GLWasInit Then
    Dim OldGLDC As Long
    Dim OldGLRC As Long
    OldGLDC = wglGetCurrentDC
    OldGLRC = wglGetCurrentContext
    GLMakeCurrent
    glClearColor Red(New_BackColor) / 255, Green(New_BackColor) / 255, Blue(New_BackColor) / 255, 1
    wglMakeCurrent OldGLDC, OldGLRC
End If

PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
UserControl.Enabled() = New_Enabled
PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set UserControl.Font = New_Font
PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As FormBorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As FormBorderStyleConstants)
If New_BorderStyle < 0 Or New_BorderStyle > 1 Then Exit Property
UserControl.BorderStyle() = New_BorderStyle
PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
If GLWasInit Then
    Dim OldGLDC As Long
    Dim OldGLRC As Long
    OldGLDC = wglGetCurrentDC
    OldGLRC = wglGetCurrentContext
    GLMakeCurrent
    Redraw
    wglMakeCurrent OldGLDC, OldGLRC
Else
    UserControl.Refresh
End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Perspective() As Boolean
Perspective = m_Perspective
End Property

Public Property Let Perspective(ByVal New_Perspective As Boolean)
m_Perspective = New_Perspective
PropertyChanged "Perspective"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BorderBox() As Boolean
BorderBox = m_BorderBox
End Property

Public Property Let BorderBox(ByVal New_BorderBox As Boolean)
m_BorderBox = New_BorderBox
PropertyChanged "BorderBox"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Axes() As Boolean
Axes = m_Axes
End Property

Public Property Let Axes(ByVal New_Axes As Boolean)
m_Axes = New_Axes
PropertyChanged "Axes"
End Property

Public Property Get Mode() As WindowMode
Mode = WMode
End Property

Public Property Let Mode(ByVal vNewValue As WindowMode)
If WMode = vNewValue Then Exit Property
WMode = vNewValue
'TerminateGL
End Property

Private Sub mnuProperties_Click()
Me.Enabled = False
Load frmPolyViewerProps
Set frmPolyViewerProps.Canvas = Me
frmPolyViewerProps.Fill
frmPolyViewerProps.Show vbModal
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
'GLMakeCurrent
End Sub

Private Sub UserControl_Initialize()
'If WasSelfDestroyed Then Exit Sub
Randomize Timer
mnuProperties.Caption = GetString(ResView_Caption)
'WasSelfDestroyed = False
WMode = wmdViewer
GLWasInit = False
ScaleFactor = 1
m_LineWidth = 1

AddGLData Me

InitGL

'GLDataCount = GLDataCount + 1
'ReDim Preserve GLData(1 To GLDataCount)
'GLData(GLDataCount).Index = GLDataCount
'Set GLData(GLDataCount).Viewer = Me
'm_DataIndex = GLDataCount
End Sub

Public Sub SelfDestruct()
UserControl_Terminate
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Const AngleDelta = 5
Const ScaleDelta = 0

Select Case KeyCode
Case vbKeyUp
    AngleVertic = AngleVertic - AngleDelta
    If AngleVertic < 0 Then AngleVertic = AngleVertic + 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyDown
    AngleVertic = AngleVertic + AngleDelta
    If AngleVertic > 360 Then AngleVertic = AngleVertic - 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyLeft
    AngleHoriz = AngleHoriz - AngleDelta
    If AngleHoriz < 0 Then AngleHoriz = AngleHoriz + 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyRight
    AngleHoriz = AngleHoriz + AngleDelta
    If AngleHoriz > 360 Then AngleHoriz = AngleHoriz - 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyPageUp
    AngleTraversal = AngleTraversal - AngleDelta
    If AngleTraversal < 0 Then AngleTraversal = AngleTraversal + 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyPageDown
    AngleTraversal = AngleTraversal + AngleDelta
    If AngleTraversal > 360 Then AngleTraversal = AngleTraversal - 360
    Redraw
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
Case vbKeyHome, vbKeyAdd
    ScaleFactor = ScaleFactor + ScaleDelta
    If ScaleFactor > 10 Then ScaleFactor = 10
    Redraw
Case vbKeyEnd, vbKeySubtract
    ScaleFactor = ScaleFactor - ScaleDelta
    If ScaleFactor < 0.1 Then ScaleFactor = 0.1
    Redraw
End Select
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################
'##########################################################################







Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dragging Then
    'GLFreeCurrent
    Dragging = False
End If
If Button = 1 And Shift = 0 And GLWasInit Then
    GLMakeCurrent
    OX = X
    OY = Y
    OldAngleHoriz = AngleHoriz
    OldAngleVertic = AngleVertic
    Dragging = True
End If
If Button = 2 And Shift = 0 And GLWasInit Then
    GLMakeCurrent
    OX = X
    OY = Y
    OldScaleFactor = ScaleFactor
    Dragging = True
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dragging And Button = 1 And Shift = 0 Then
    AngleHoriz = OldAngleHoriz + (X - OX)
    AngleVertic = OldAngleVertic + (Y - OY)
    NormalizeAngle AngleHoriz
    NormalizeAngle AngleVertic
    Refresh
    RaiseEvent Rotated(AngleHoriz, AngleVertic, AngleTraversal)
End If
If Dragging And Button = 2 And Shift = 0 Then
    ScaleFactor = OldScaleFactor + (Y - OY) / 100
    Refresh
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dragging Then
    Dragging = False
    'GLFreeCurrent
End If
If Button = 2 And GLWasInit And X = OX And Y = OY Then
    PopupMenu mnuMain, , X, Y
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub NormalizeAngle(ByRef A As Double)
If A > 360 Then A = A - 360
If A < 0 Then A = A + 360
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddPoint(X As Double, Y As Double, Z As Double, Optional ByVal Color As Long = -1, Optional ByVal Visible As Boolean = True) As Long
Const C1 = 256
If Color = -1 Then
    'Color = RGB(Rnd * 128, Rnd * 128, Rnd * 128)
    'Color = RGB(Rnd * 128 + 128, Rnd * 128 + 128, Rnd * 128 + 128)
    Color = GetRandomColor
    'Color = RGB(0, Rnd * 256, Rnd * 256)
    'Color = RGB(Rnd * 255, 0, Rnd * 255)
    'Color = RGB(Rnd * 256, Rnd * 256, 0)
    'Color = RGB(0, 0, Rnd * 256)
    'Color = RGB(0, Rnd * 256, 0)
    'Color = RGB(Rnd * 256, 0, 0)
    'Color = m_ForeColor
End If
With GLData(m_DataIndex)
    .Points3DCount = .Points3DCount + 1
    ReDim Preserve .Points3D(1 To .Points3DCount)
    .Points3D(.Points3DCount).X = X
    .Points3D(.Points3DCount).Y = Y
    .Points3D(.Points3DCount).Z = Z
    .Points3D(.Points3DCount).Color = Color
    .Points3D(.Points3DCount).Visible = Visible
    
    AddPoint = .Points3DCount
End With
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub AddLine(P1 As Long, P2 As Long)
With GLData(m_DataIndex)
    .Lines3DCount = .Lines3DCount + 1
    ReDim Preserve .Lines3D(1 To .Lines3DCount)
    .Lines3D(.Lines3DCount).P1 = P1
    .Lines3D(.Lines3DCount).P2 = P2
End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub AddPolygon(P() As Long, Optional ByVal NeedsTesselation As Boolean = False, Optional ByVal Color As Long = -1)
Dim X() As Double, Y() As Double, Z() As Double, Count As Long, Q As Long

Dim Az As Double, UB As Long, LB As Long

With GLData(m_DataIndex)

    Count = UBound(P)
    'For Q = 1 To Count
    '    Az = Az + .Points3D(Q).Z
    'Next
    'Az = Az / Count

    .Polygons3DCount = .Polygons3DCount + 1
    ReDim Preserve .Polygons3D(1 To .Polygons3DCount)
    .Polygons3D(.Polygons3DCount).P() = P()
    .Polygons3D(.Polygons3DCount).PointCount = Count
    '.Polygons3D(.Polygons3DCount).ZOrder = Az
    ReDim Preserve .Polygons3D(.Polygons3DCount).Color(1 To Count)
    
'    UB = .PolygonOrder.Count
'    If UB > 0 Then
'        LB = 1
'        Do
'            Q = (UB + LB) \ 2
'            If Az > .Polygons3D(.PolygonOrder(Q)).ZOrder Then
'                LB = Q + 1
'                If LB > UB Then .PolygonOrder.Add .Polygons3DCount, , , Q
'            Else
'                UB = Q - 1
'                If LB > UB Then .PolygonOrder.Add .Polygons3DCount, , Q
'            End If
'        Loop Until UB < LB
'    Else
'        .PolygonOrder.Add 1
'    End If
    
    If m_SingleColorPolygons Then
        If Color = -1 Then
            If m_RandomPolygonColor Then
                .Polygons3D(.Polygons3DCount).Col = GetRandomColor
            Else
                .Polygons3D(.Polygons3DCount).Col = m_PolygonColor
            End If
        Else
            .Polygons3D(.Polygons3DCount).Col = Color
        End If
    Else
        For Q = 1 To Count
            If m_RandomPolygonColor Then
                .Polygons3D(.Polygons3DCount).Color(Q) = GetRandomColor
            Else
                .Polygons3D(.Polygons3DCount).Color(Q) = .Points3D(.Polygons3D(.Polygons3DCount).P(Q)).Color
            End If
        Next
    End If
    
    If NeedsTesselation Then
        ReDim X(1 To Count)
        ReDim Y(1 To Count)
        ReDim Z(1 To Count)
        For Q = 1 To Count
            X(Q) = .Points3D(P(Q)).X
            Y(Q) = .Points3D(P(Q)).Y
            Z(Q) = .Points3D(P(Q)).Z
        Next
        TessIndex = m_DataIndex
        TesselatePolygon X, Y, Z, Count, GLData(m_DataIndex).Polygons3D(GLData(m_DataIndex).Polygons3DCount)
    End If
End With
End Sub

Public Sub UserControl_Paint()
If Not GLWasInit Then Exit Sub

Dim OldGLDC As Long
Dim OldGLRC As Long
OldGLDC = wglGetCurrentDC
OldGLRC = wglGetCurrentContext

GLMakeCurrent
Redraw
wglMakeCurrent OldGLDC, OldGLRC
End Sub

Private Sub UserControl_InitProperties()
UserControl.BackColor = m_def_BackColor
m_DrawFaces = m_def_DrawFaces
m_DrawLines = m_def_DrawLines
m_DrawPoints = m_def_DrawPoints
m_SingleColorPolygons = m_def_SingleColorPolygons
m_RandomPolygonColor = m_def_RandomPolygonColor
m_Transparency = 1
m_BoxSize = 2
m_PointColor = m_def_PointColor
m_LineColor = 0
m_LineWidth = 1
m_PolygonColor = m_def_PolygonColor
m_PointSize = m_def_PointSize
m_Pastel = m_def_Pastel
m_Brightness = m_def_Brightness
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
m_DrawFaces = PropBag.ReadProperty("DrawFaces", m_def_DrawFaces)
m_DrawLines = PropBag.ReadProperty("DrawLines", m_def_DrawLines)
m_DrawPoints = PropBag.ReadProperty("DrawPoints", m_def_DrawPoints)
m_Transparency = PropBag.ReadProperty("Transparency", 1)
m_SingleColorPolygons = PropBag.ReadProperty("SingleColorPolygons", m_def_SingleColorPolygons)
m_RandomPolygonColor = PropBag.ReadProperty("RandomPolygonColor", m_def_RandomPolygonColor)
m_PointColor = PropBag.ReadProperty("PointColor", m_def_PointColor)
m_LineColor = PropBag.ReadProperty("LineColor", 0)
m_PolygonColor = PropBag.ReadProperty("PolygonColor", m_def_PolygonColor)
m_PointSize = PropBag.ReadProperty("PointSize", m_def_PointSize)
m_Pastel = PropBag.ReadProperty("Pastel", m_def_Pastel)
m_Brightness = PropBag.ReadProperty("Brightness", m_def_Brightness)
m_LineWidth = PropBag.ReadProperty("LineWidth", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BackColor", UserControl.BackColor, m_def_BackColor
PropBag.WriteProperty "DrawFaces", m_DrawFaces, m_def_DrawFaces
PropBag.WriteProperty "DrawLines", m_DrawLines, m_def_DrawLines
PropBag.WriteProperty "DrawPoints", m_DrawPoints, m_def_DrawPoints
PropBag.WriteProperty "Transparency", m_Transparency, 1
PropBag.WriteProperty "SingleColorPolygons", m_SingleColorPolygons, m_def_SingleColorPolygons
PropBag.WriteProperty "RandomPolygonColor", m_RandomPolygonColor, m_def_RandomPolygonColor
PropBag.WriteProperty "PointColor", m_PointColor, m_def_PointColor
PropBag.WriteProperty "LineColor", m_LineColor, 0
PropBag.WriteProperty "PolygonColor", m_PolygonColor, m_def_PolygonColor
PropBag.WriteProperty "PointSize", m_PointSize, m_def_PointSize
PropBag.WriteProperty "Pastel", m_Pastel, m_def_Pastel
PropBag.WriteProperty "Brightness", m_Brightness, m_def_Brightness
PropBag.WriteProperty "LineWidth", m_LineWidth, 1
End Sub

Public Sub UserControl_Resize()
If Not GLWasInit Then Exit Sub
If UserControl.ScaleHeight < 1 Then UserControl.Height = 1

Dim OldGLDC As Long
Dim OldGLRC As Long
OldGLDC = wglGetCurrentDC
OldGLRC = wglGetCurrentContext

GLMakeCurrent
ResizeGL
Redraw
wglMakeCurrent OldGLDC, OldGLRC
End Sub

Public Sub ResizeGL(Optional ShouldSetViewport As Boolean = True)
If Not GLWasInit Then Exit Sub

If ShouldSetViewport Then glViewport 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1

glMatrixMode GL_PROJECTION
glLoadIdentity
glFrustum -BoxSize, BoxSize, -BoxSize, BoxSize, CenterOffset + BoxSize, 100 + CenterOffset + BoxSize * 3
'gluPerspective 45, 1, 1, 100
glMatrixMode GL_MODELVIEW

End Sub

Public Sub CompileStructure()
Dim Z As Long
Dim Q As Long, M As Long, N As Long, C As Long

If Not GLWasInit Then Exit Sub

Dim OldGLDC As Long
Dim OldGLRC As Long
OldGLDC = wglGetCurrentDC
OldGLRC = wglGetCurrentContext

GLMakeCurrent

NewList m_DataIndex
glNewList GLData(m_DataIndex).ListID, GL_COMPILE

With GLData(m_DataIndex)
    
    'FACES=====================================================
    If DrawFaces And .Points3DCount > 0 Then
        For Z = 1 To .Polygons3DCount
            If .Polygons3D(Z).NeedsTesselation Then
                
                If m_SingleColorPolygons Then
                    C = .Polygons3D(Z).Col
                    glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                End If
                
                With .Polygons3D(Z)
                    For M = 1 To .TriangleCount
                        glBegin GL_TRIANGLES
                            For N = 1 To .Triangles(M).Count
                                If Not m_SingleColorPolygons Then
                                    C = .Fans(M).P(N).Color
                                    glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                                End If
                                glVertex3f .Triangles(M).P(N).X, .Triangles(M).P(N).Y, .Triangles(M).P(N).Z
                            Next
                        glEnd
                    Next
                    
                    For M = 1 To .FanCount
                        glBegin GL_TRIANGLE_FAN
                            For N = 1 To .Fans(M).Count
                                If Not m_SingleColorPolygons Then
                                    C = .Fans(M).P(N).Color
                                    glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                                End If
                                glVertex3f .Fans(M).P(N).X, .Fans(M).P(N).Y, .Fans(M).P(N).Z
                            Next
                        glEnd
                    Next
                
                    For M = 1 To .StripCount
                        glBegin GL_TRIANGLE_STRIP
                            For N = 1 To .Strips(M).Count
                                If Not m_SingleColorPolygons Then
                                    C = .Fans(M).P(N).Color
                                    glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                                End If
                                glVertex3f .Strips(M).P(N).X, .Strips(M).P(N).Y, .Strips(M).P(N).Z
                            Next
                        glEnd
                    Next
                
                End With
                
            Else 'do not tesselate
                glBegin GL_POLYGON
                
                If m_SingleColorPolygons Then
                    C = .Polygons3D(Z).Col
                    glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                End If
                
                For Q = 1 To .Polygons3D(Z).PointCount
                    If Not m_SingleColorPolygons Then
                        C = .Polygons3D(Z).Color(Q)
                        glColor4f Red(C) / 255, Green(C) / 255, Blue(C) / 255, m_Transparency
                    End If
                    glVertex3f .Points3D(.Polygons3D(Z).P(Q)).X, .Points3D(.Polygons3D(Z).P(Q)).Y, .Points3D(.Polygons3D(Z).P(Q)).Z
                Next
                
                glEnd
            End If 'tesselate?
        
        Next
    End If 'draw faces
    '//FACES=====================================================
    
    
    'EDGES======================================================
'    If DrawLines Then
'
'        'glClear GL_DEPTH_BUFFER_BIT
'
'        glPolygonMode GL_FRONT, GL_LINE
'        glEnable GL_CULL_FACE
'
'        glColor4f Red(m_LineColor) / 255, Green(m_LineColor) / 255, Blue(m_LineColor) / 255, 1
'        glDepthFunc GL_LEQUAL
'
'        For Z = 1 To .Polygons3DCount
'            glBegin GL_POLYGON
'            For Q = 1 To .Polygons3D(Z).PointCount
'                glVertex3f .Points3D(.Polygons3D(Z).P(Q)).X, .Points3D(.Polygons3D(Z).P(Q)).Y, .Points3D(.Polygons3D(Z).P(Q)).Z
'            Next
'            glEnd
'        Next
'
'        glDepthFunc GL_LESS
'        glDisable GL_CULL_FACE
'        glPolygonMode GL_FRONT, GL_FILL
'    End If
    '//EDGES======================================================

'    If .Lines3DCount > 0 And DrawLines Then
'        glTranslatef 0, 0, 0.2
'
'        glBegin GL_LINES
'            glColor4f (Red(m_LineColor)) / 255, (Green(m_LineColor)) / 255, (Blue(m_LineColor)) / 255, m_Transparency
'            For Z = 1 To .Lines3DCount
'                glVertex3f .Points3D(.Lines3D(Z).P1).X, .Points3D(.Lines3D(Z).P1).Y, .Points3D(.Lines3D(Z).P1).Z
'                glVertex3f .Points3D(.Lines3D(Z).P2).X, .Points3D(.Lines3D(Z).P2).Y, .Points3D(.Lines3D(Z).P2).Z
'            Next
'        glEnd
'    End If
    
    
    
    'VERTICES====================================================
    If DrawPoints Then
        If m_PointSize <> 1 Then glPointSize m_PointSize
        'glColor4f Red(m_PointColor) / 255, Green(m_PointColor) / 255, Blue(m_PointColor) / 255, m_Transparency
        glBegin GL_POINTS
            For Z = 1 To .Points3DCount
                If .Points3D(Z).Visible Then
                    glColor4f Red(.Points3D(Z).Color) / 255, Green(.Points3D(Z).Color) / 255, Blue(.Points3D(Z).Color) / 255, m_Transparency
                    glVertex3f .Points3D(Z).X, .Points3D(Z).Y, .Points3D(Z).Z
                End If
            Next
        glEnd
    End If
    '//VERTICES====================================================
End With
glEndList

NewLinesList m_DataIndex
If DrawLines Then
    glNewList GLData(m_DataIndex).ListLinesID, GL_COMPILE
    
    glLineWidth LineWidth
    
    With GLData(m_DataIndex)
        'glColor4f Red(m_LineColor) / 255, Green(m_LineColor) / 255, Blue(m_LineColor) / 255, m_Transparency

        glBegin GL_LINES
            For Z = 1 To .Lines3DCount
                glColor4f Red(.Points3D(.Lines3D(Z).P1).Color) / 255, Green(.Points3D(.Lines3D(Z).P1).Color) / 255, Blue(.Points3D(.Lines3D(Z).P1).Color) / 255, m_Transparency
                glVertex3f .Points3D(.Lines3D(Z).P1).X, .Points3D(.Lines3D(Z).P1).Y, .Points3D(.Lines3D(Z).P1).Z
                glColor4f Red(.Points3D(.Lines3D(Z).P1).Color) / 255, Green(.Points3D(.Lines3D(Z).P1).Color) / 255, Blue(.Points3D(.Lines3D(Z).P1).Color) / 255, m_Transparency
                glVertex3f .Points3D(.Lines3D(Z).P2).X, .Points3D(.Lines3D(Z).P2).Y, .Points3D(.Lines3D(Z).P2).Z
            Next
        glEnd
    End With
    glEndList
End If

UpdateBoxSize

wglMakeCurrent OldGLDC, OldGLRC
End Sub

Public Sub Redraw()
If Not GLWasInit Then Exit Sub

glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

glPushMatrix
    
    glLoadIdentity
    If ScaleFactor <> 1 Then
        'glScalef ScaleFactor, ScaleFactor, ScaleFactor
        glTranslatef 0, 0, -(ScaleFactor - 1) * 10
    End If
    
    glTranslatef 0, 0, -(CenterOffset + m_BoxSize * 2)
    RaiseEvent BeforeRepaint
    
    If glIsList(GLData(m_DataIndex).ListID) <> 0 Then
        glLoadIdentity
        If ScaleFactor <> 1 Then
            'glScalef ScaleFactor, ScaleFactor, ScaleFactor
            glTranslatef 0, 0, -(ScaleFactor - 1) * 10
        End If
        glTranslatef 0, 0, -(CenterOffset + m_BoxSize * 2)
        glRotatef AngleVertic, 1, 0, 0
        glRotatef AngleHoriz, 0, 1, 0
        glRotatef AngleTraversal, 0, 0, 1
        glCallList GLData(m_DataIndex).ListID
    End If
    
    If DrawLines And glIsList(GLData(m_DataIndex).ListLinesID) <> 0 Then
        glLoadIdentity
        If ScaleFactor <> 1 Then
            'glScalef ScaleFactor, ScaleFactor, ScaleFactor
            glTranslatef 0, 0, -(ScaleFactor - 1) * 10
        End If
        glTranslatef 0, 0, -(CenterOffset + m_BoxSize * 2) + 0.015
        glRotatef AngleVertic, 1, 0, 0
        glRotatef AngleHoriz, 0, 1, 0
        glRotatef AngleTraversal, 0, 0, 1
        glCallList GLData(m_DataIndex).ListLinesID
    End If
    
    glLoadIdentity
    If ScaleFactor <> 1 Then
        'glScalef ScaleFactor, ScaleFactor, ScaleFactor
        glTranslatef 0, 0, -(ScaleFactor - 1) * 10
    End If
    glTranslatef 0, 0, -(CenterOffset + m_BoxSize * 2)
    RaiseEvent AfterRepaint
    
glPopMatrix

wglSwapBuffers GLData(m_DataIndex).glDC
End Sub

'=================================================
'
'=================================================

Public Sub InitGL()
On Local Error GoTo EH:

If GLWasInit Then Exit Sub

ClearGLData DataIndex

'==================================

GLData(DataIndex).glWnd = UserControl.hWnd
If Not InitializeOpenGL(GLData(DataIndex)) Then
    MsgBox "Unable to initialize OpenGL."
    Exit Sub
End If

GLWasInit = True

'==================================

'glClearDepth 1
Me.Transparency = m_Transparency
BackColor = UserControl.BackColor
CenterOffset = 4

'==================================

'If m_Transparency = 1 Then
glEnable GL_DEPTH_TEST

'glEnable GL_LIGHTING
'glEnable GL_LIGHT0
'glEnable GL_LIGHT1
'glEnable GL_LIGHT2

'glEnable GL_POLYGON_SMOOTH
'glEnable GL_LINE_SMOOTH
'glEnable GL_POINT_SMOOTH
'glLineWidth 1

'glEnable GL_CULL_FACE
glPolygonMode GL_BACK, GL_FILL
glPolygonMode GL_FRONT, GL_FILL

RaiseEvent AfterInitGL

'==================================

ResizeGL False
Redraw

Exit Sub
'==================================
EH:
End Sub

'=================================================
'
'=================================================

Public Sub TerminateGL()
If GLWasInit Then
    GLWasInit = False
    TerminateOpenGL GLData(m_DataIndex)
    RemoveGLData m_DataIndex
    Refresh
End If
End Sub

'If Not GLWasInit Then Exit Sub
'
'GLWasInit = False
'
'
''TerminateOpenGL GLData(m_DataIndex)
''ClearGLData DataIndex
''UserControl.Cls
'End Sub

Private Sub UserControl_Terminate()
'If WasSelfDestroyed Then Exit Sub

TerminateGL

'WasSelfDestroyed = True
End Sub

'=================================================
'
'=================================================

Public Sub Clear()
ClearGLData DataIndex
Refresh
End Sub

'=================================================
'
'=================================================

Public Property Get DataIndex() As Long
DataIndex = m_DataIndex
End Property

Public Property Let DataIndex(ByVal vNewValue As Long)
m_DataIndex = vNewValue
End Property

Public Sub Activate()
If Not GLWasInit Then Exit Sub
GLMakeCurrent
Redraw
End Sub

Public Sub Deactivate()
'GLFreeCurrent
End Sub

Public Property Get Transparency() As Double
Transparency = m_Transparency
End Property

Public Property Let Transparency(ByVal vNewValue As Double)
If vNewValue < 0 Or vNewValue > 1 Then Exit Property

m_Transparency = vNewValue

If GLWasInit Then
    Dim OldGLDC As Long
    Dim OldGLRC As Long
    OldGLDC = wglGetCurrentDC
    OldGLRC = wglGetCurrentContext

    GLMakeCurrent
    If vNewValue = 1 Then
        glDisable GL_BLEND
    Else
        glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        glEnable GL_BLEND
    End If
    wglMakeCurrent OldGLDC, OldGLRC
End If

PropertyChanged "Transparency"
End Property

Public Property Get DrawLines() As Boolean
DrawLines = m_DrawLines
End Property

Public Property Let DrawLines(ByVal vNewValue As Boolean)
m_DrawLines = vNewValue
PropertyChanged "DrawLines"
End Property

Public Property Get DrawPoints() As Boolean
DrawPoints = m_DrawPoints
End Property

Public Property Let DrawPoints(ByVal vNewValue As Boolean)
m_DrawPoints = vNewValue
PropertyChanged "DrawPoints"
End Property

Public Property Get DrawFaces() As Boolean
DrawFaces = m_DrawFaces
End Property

Public Property Let DrawFaces(ByVal vNewValue As Boolean)
m_DrawFaces = vNewValue
PropertyChanged "DrawFaces"
End Property

Public Property Get BoxSize() As Double
BoxSize = m_BoxSize
End Property

Public Property Let BoxSize(ByVal vNewValue As Double)
If vNewValue < Epsilon Then Exit Property
m_BoxSize = vNewValue

If Not GLWasInit Then Exit Property
Dim OldGLDC As Long
Dim OldGLRC As Long
OldGLDC = wglGetCurrentDC
OldGLRC = wglGetCurrentContext

GLMakeCurrent
ResizeGL False
wglMakeCurrent OldGLDC, OldGLRC
End Property

Public Sub GLMakeCurrent()
wglMakeCurrent GLData(m_DataIndex).glDC, GLData(m_DataIndex).glRC
'Dim OldDC As Long, OldRC As Long, DIndex As Long
'
'If Not GLWasInit Then Exit Sub
'If DCMadeBackup Then GLFreeCurrent
'
'theOldDC = 0
'theOldRC = 0
'OldDC = wglGetCurrentDC
'OldRC = wglGetCurrentContext
'DIndex = m_DataIndex
'
'DCNeedsBackup = (OldDC <> GLData(DIndex).glDC) Or (OldRC <> GLData(DIndex).glRC)
'
'If DCNeedsBackup Then
'    DCMadeBackup = True
'    theOldDC = OldDC
'    theOldRC = OldRC
'
'    wglMakeCurrent GLData(DIndex).glDC, GLData(DIndex).glRC
'End If
End Sub

Public Sub GLFreeCurrent(Optional ByVal ShouldMakeCurrent As Boolean = True)
wglMakeCurrent 0, 0
'If DCMadeBackup And GLWasInit Then
'    DCMadeBackup = False
'    If ShouldMakeCurrent Then wglMakeCurrent theOldDC, theOldRC
'    theOldDC = 0
'    theOldRC = 0
'End If
End Sub

Public Property Get SingleColorPolygons() As Boolean
SingleColorPolygons = m_SingleColorPolygons
End Property

Public Property Let SingleColorPolygons(ByVal vNewValue As Boolean)
m_SingleColorPolygons = vNewValue
End Property

Public Property Get RandomPolygonColor() As Boolean
RandomPolygonColor = m_RandomPolygonColor
End Property

Public Property Let RandomPolygonColor(ByVal vNewValue As Boolean)
m_RandomPolygonColor = vNewValue
End Property

Public Property Get LineColor() As OLE_COLOR
LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal vNewValue As OLE_COLOR)
m_LineColor = vNewValue
End Property

Public Property Get PolygonColor() As OLE_COLOR
PolygonColor = m_PolygonColor
End Property

Public Property Let PolygonColor(ByVal vNewValue As OLE_COLOR)
m_PolygonColor = vNewValue
End Property

Public Property Get PointColor() As OLE_COLOR
PointColor = m_PointColor
End Property

Public Property Let PointColor(ByVal vNewValue As OLE_COLOR)
m_PointColor = vNewValue
End Property

Public Property Get PointSize() As Long
PointSize = m_PointSize
End Property

Public Property Let PointSize(ByVal vNewValue As Long)
m_PointSize = vNewValue
End Property

Public Property Get LineWidth() As Long
LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal vNewValue As Long)
m_LineWidth = vNewValue
End Property

Public Sub UpdateBoxSize()
Dim M As Double, N As Double, Z As Long

With GLData(m_DataIndex)
    For Z = 1 To .Points3DCount
        N = .Points3D(Z).X * .Points3D(Z).X + .Points3D(Z).Y * .Points3D(Z).Y + .Points3D(Z).Z * .Points3D(Z).Z
        If M < N Then M = N
    Next
End With
M = Sqr(M) + BoxSizeEpsilon * 2
BoxSize = M
End Sub

Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
UserControl_KeyDown KeyCode, Shift
End Sub

Public Property Get RandomColorPastel() As Single
RandomColorPastel = m_Pastel
End Property

Public Property Let RandomColorPastel(ByVal vNewValue As Single)
m_Pastel = vNewValue
End Property

Private Function Random(ByVal A As Double, ByVal B As Double) As Double
Random = A + (B - A) * Rnd
End Function

Public Property Get RandomColorBrightness() As Single
RandomColorBrightness = m_Brightness
End Property

Public Property Let RandomColorBrightness(ByVal vNewValue As Single)
m_Brightness = vNewValue
End Property

Public Function GetRandomColor() As Long
Dim M As Double, LB As Double, UB As Double

M = Random(m_Brightness, 255)

LB = M - m_Pastel
If LB < m_Brightness Then LB = m_Brightness
UB = M + m_Pastel
If UB > 255 Then UB = 255

GetRandomColor = RGB(Random(LB, UB), Random(LB, UB), Random(LB, UB))
End Function

Public Sub ReInitColors()
Dim Z As Long, Q As Long

With GLData(m_DataIndex)
    For Z = 1 To .Points3DCount
        '.Points3D(Z).Color = GetRandomColor
    Next
    
    For Z = 1 To .Polygons3DCount
        If m_SingleColorPolygons Then
            'If m_RandomPolygonColor Then .Polygons3D(Z).Col = GetRandomColor
        Else
            For Q = 1 To Count
                If m_RandomPolygonColor Then
                    .Polygons3D(Z).Color(Q) = GetRandomColor
                Else
                    .Polygons3D(Z).Color(Q) = .Points3D(.Polygons3D(Z).P(Q)).Color
                End If
            Next
        End If
    Next
End With
End Sub

Public Sub OutputScanningModel(P As CPolyhedron, ByVal Stage As Double)
If Not GLWasInit Then Exit Sub
Dim OldGLDC As Long
Dim OldGLRC As Long
OldGLDC = wglGetCurrentDC
OldGLRC = wglGetCurrentContext

GLMakeCurrent

NewList m_DataIndex
glNewList GLData(m_DataIndex).ListID, GL_COMPILE
    P.OutputScanningModel Me, Stage
glEndList

BoxSize = P.MaxVertexNorm
ResizeGL False
Refresh

wglMakeCurrent OldGLDC, OldGLRC
End Sub

Public Sub ShouldResize()
If GLWasInit Then
    Dim OldGLDC As Long
    Dim OldGLRC As Long
    OldGLDC = wglGetCurrentDC
    OldGLRC = wglGetCurrentContext
    GLMakeCurrent
    ResizeGL
    wglMakeCurrent OldGLDC, OldGLRC
End If
End Sub
