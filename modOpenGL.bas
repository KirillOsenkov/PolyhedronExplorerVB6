Attribute VB_Name = "modOpenGL"
Option Explicit
#Const SmoothPoints = True

Public Type Point3D
    X As Double
    Y As Double
    Z As Double
    Color As Long
    Red As Single
    Green As Single
    Blue As Single
    Alpha As Single
    Visible As Boolean
End Type

Public Type Line3D
    P1 As Long
    P2 As Long
End Type

Public Type GLFan
    P() As Point3D
    Color As Long
    Count As Long
End Type

Public Type GLStrip
    P() As Point3D
    Color As Long
    Count As Long
End Type

Public Type GLTriangle
    P() As Point3D
    Color As Long
    Count As Long
End Type

Public Type Polygon3D
    P() As Long
    Col As Long
    Color() As Long
    Red() As Single
    Green() As Single
    Blue() As Single
    Alpha() As Single
    PointCount As Long
    NeedsTesselation As Boolean
    Fans() As GLFan
    Strips() As GLStrip
    Triangles() As GLTriangle
    FanCount As Long
    StripCount As Long
    TriangleCount As Long
    ZOrder As Double
End Type

Public Type Data3D
    Points3D() As Point3D
    Lines3D() As Line3D
    Polygons3D() As Polygon3D
    Points3DCount As Long
    Lines3DCount As Long
    Polygons3DCount As Long
    PolygonOrder As New Collection
    glWnd As Long
    glDC As Long
    glRC As Long
    Viewer As Object
    Index As Long
    ListID As Long
    ListLinesID As Long
End Type

Public TessPoly As Polygon3D
Public TessState As Long
Public TessIndex As Long

Public GLData() As Data3D
Public GLDataCount As Long

'===========================================

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Public Type POINTFLOAT
    X As Single
    Y As Single
End Type

Public Type GLYPHMETRICSFLOAT
    gmfBlackBoxX As Single
    gmfBlackBoxY As Single
    gmfptGlyphOrigin As POINTFLOAT
    gmfCellIncX As Single
    gmfCellIncY As Single
End Type

Public Declare Function wglUseFontOutlines Lib "OpenGL32.dll" (hdc As Long, first As Long, Count As Long, listBase As Long, deviation As Single, extrusion As Single, Format As Long, lpgmf As GLYPHMETRICSFLOAT) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public ARIAL36&
Public TIMES36&
Public Scene1 As Long
Const MAX_STRING = 1024
Const WGL_FONT_POLYGONS = 1


'OPENGL32.DLL

'***********************************************************

' Version
Public Const GL_VERSION_1_1 = 1

' AccumOp
Public Const GL_ACCUM = &H100
Public Const GL_LOAD = &H101
Public Const GL_RETURN = &H102
Public Const GL_MULT = &H103
Public Const GL_ADD = &H104

' AlphaFunction
Public Const GL_NEVER = &H200
Public Const GL_LESS = &H201
Public Const GL_EQUAL = &H202
Public Const GL_LEQUAL = &H203
Public Const GL_GREATER = &H204
Public Const GL_NOTEQUAL = &H205
Public Const GL_GEQUAL = &H206
Public Const GL_ALWAYS = &H207

' AttribMask
Public Const GL_CURRENT_BIT = &H1
Public Const GL_POINT_BIT = &H2
Public Const GL_LINE_BIT = &H4
Public Const GL_POLYGON_BIT = &H8
Public Const GL_POLYGON_STIPPLE_BIT = &H10
Public Const GL_PIXEL_MODE_BIT = &H20
Public Const GL_LIGHTING_BIT = &H40
Public Const GL_FOG_BIT = &H80
Public Const GL_DEPTH_BUFFER_BIT = &H100
Public Const GL_ACCUM_BUFFER_BIT = &H200
Public Const GL_STENCIL_BUFFER_BIT = &H400
Public Const GL_VIEWPORT_BIT = &H800
Public Const GL_TRANSFORM_BIT = &H1000
Public Const GL_ENABLE_BIT = &H2000
Public Const GL_COLOR_BUFFER_BIT = &H4000
Public Const GL_HINT_BIT = &H8000
Public Const GL_EVAL_BIT = &H10000
Public Const GL_LIST_BIT = &H20000
Public Const GL_TEXTURE_BIT = &H40000
Public Const GL_SCISSOR_BIT = &H80000
Public Const GL_ALL_ATTRIB_BITS = &HFFFFF

' BeginMode
Public Const GL_POINTS = &H0
Public Const GL_LINES = &H1
Public Const GL_LINE_LOOP = &H2
Public Const GL_LINE_STRIP = &H3
Public Const GL_TRIANGLES = &H4
Public Const GL_TRIANGLE_STRIP = &H5
Public Const GL_TRIANGLE_FAN = &H6
Public Const GL_QUADS = &H7
Public Const GL_QUAD_STRIP = &H8
Public Const GL_POLYGON = &H9

' BlendingFactorDest
Public Const GL_ZERO = 0
Public Const GL_ONE = 1
Public Const GL_SRC_COLOR = &H300
Public Const GL_ONE_MINUS_SRC_COLOR = &H301
Public Const GL_SRC_ALPHA = &H302
Public Const GL_ONE_MINUS_SRC_ALPHA = &H303
Public Const GL_DST_ALPHA = &H304
Public Const GL_ONE_MINUS_DST_ALPHA = &H305

' BlendingFactorSrc
'      GL_ZERO
'      GL_ONE
Public Const GL_DST_COLOR = &H306
Public Const GL_ONE_MINUS_DST_COLOR = &H307
Public Const GL_SRC_ALPHA_SATURATE = &H308
'      GL_SRC_ALPHA
'      GL_ONE_MINUS_SRC_ALPHA
'      GL_DST_ALPHA
'      GL_ONE_MINUS_DST_ALPHA

' Boolean
Public Const GL_TRUE = 1
Public Const GL_FALSE = 0

' ClearBufferMask
'      GL_COLOR_BUFFER_BIT
'      GL_ACCUM_BUFFER_BIT
'      GL_STENCIL_BUFFER_BIT
'      GL_DEPTH_BUFFER_BIT

' ClientArrayType
'      GL_VERTEX_ARRAY
'      GL_NORMAL_ARRAY
'      GL_COLOR_ARRAY
'      GL_INDEX_ARRAY
'      GL_TEXTURE_COORD_ARRAY
'      GL_EDGE_FLAG_ARRAY

' ClipPlaneName
Public Const GL_CLIP_PLANE0 = &H3000
Public Const GL_CLIP_PLANE1 = &H3001
Public Const GL_CLIP_PLANE2 = &H3002
Public Const GL_CLIP_PLANE3 = &H3003
Public Const GL_CLIP_PLANE4 = &H3004
Public Const GL_CLIP_PLANE5 = &H3005

' ColorMaterialFace
'      GL_FRONT
'      GL_BACK
'      GL_FRONT_AND_BACK

' ColorMaterialParameter
'      GL_AMBIENT
'      GL_DIFFUSE
'      GL_SPECULAR
'      GL_EMISSION
'      GL_AMBIENT_AND_DIFFUSE

' ColorPointerType
'      GL_BYTE
'      GL_UNSIGNED_BYTE
'      GL_SHORT
'      GL_UNSIGNED_SHORT
'      GL_INT
'      GL_UNSIGNED_INT
'      GL_FLOAT
'      GL_DOUBLE

' CullFaceMode
'      GL_FRONT
'      GL_BACK
'      GL_FRONT_AND_BACK

' DataType
Public Const GL_BYTE = &H1400
Public Const GL_UNSIGNED_BYTE = &H1401
Public Const GL_SHORT = &H1402
Public Const GL_UNSIGNED_SHORT = &H1403
Public Const GL_INT = &H1404
Public Const GL_UNSIGNED_INT = &H1405
Public Const GL_FLOAT = &H1406
Public Const GL_2_BYTES = &H1407
Public Const GL_3_BYTES = &H1408
Public Const GL_4_BYTES = &H1409
Public Const GL_DOUBLE = &H140A

' DepthFunction
'      GL_NEVER
'      GL_LESS
'      GL_EQUAL
'      GL_LEQUAL
'      GL_GREATER
'      GL_NOTEQUAL
'      GL_GEQUAL
'      GL_ALWAYS

' DrawBufferMode
Public Const GL_NONE = 0
Public Const GL_FRONT_LEFT = &H400
Public Const GL_FRONT_RIGHT = &H401
Public Const GL_BACK_LEFT = &H402
Public Const GL_BACK_RIGHT = &H403
Public Const GL_FRONT = &H404
Public Const GL_BACK = &H405
Public Const GL_LEFT = &H406
Public Const GL_RIGHT = &H407
Public Const GL_FRONT_AND_BACK = &H408
Public Const GL_AUX0 = &H409
Public Const GL_AUX1 = &H40A
Public Const GL_AUX2 = &H40B
Public Const GL_AUX3 = &H40C

' Enable
'      GL_FOG
'      GL_LIGHTING
'      GL_TEXTURE_1D
'      GL_TEXTURE_2D
'      GL_LINE_STIPPLE
'      GL_POLYGON_STIPPLE
'      GL_CULL_FACE
'      GL_ALPHA_TEST
'      GL_BLEND
'      GL_INDEX_LOGIC_OP
'      GL_COLOR_LOGIC_OP
'      GL_DITHER
'      GL_STENCIL_TEST
'      GL_DEPTH_TEST
'      GL_CLIP_PLANE0
'      GL_CLIP_PLANE1
'      GL_CLIP_PLANE2
'      GL_CLIP_PLANE3
'      GL_CLIP_PLANE4
'      GL_CLIP_PLANE5
'      GL_LIGHT0
'      GL_LIGHT1
'      GL_LIGHT2
'      GL_LIGHT3
'      GL_LIGHT4
'      GL_LIGHT5
'      GL_LIGHT6
'      GL_LIGHT7
'      GL_TEXTURE_GEN_S
'      GL_TEXTURE_GEN_T
'      GL_TEXTURE_GEN_R
'      GL_TEXTURE_GEN_Q
'      GL_MAP1_VERTEX_3
'      GL_MAP1_VERTEX_4
'      GL_MAP1_COLOR_4
'      GL_MAP1_INDEX
'      GL_MAP1_NORMAL
'      GL_MAP1_TEXTURE_COORD_1
'      GL_MAP1_TEXTURE_COORD_2
'      GL_MAP1_TEXTURE_COORD_3
'      GL_MAP1_TEXTURE_COORD_4
'      GL_MAP2_VERTEX_3
'      GL_MAP2_VERTEX_4
'      GL_MAP2_COLOR_4
'      GL_MAP2_INDEX
'      GL_MAP2_NORMAL
'      GL_MAP2_TEXTURE_COORD_1
'      GL_MAP2_TEXTURE_COORD_2
'      GL_MAP2_TEXTURE_COORD_3
'      GL_MAP2_TEXTURE_COORD_4
'      GL_POINT_SMOOTH
'      GL_LINE_SMOOTH
'      GL_POLYGON_SMOOTH
'      GL_SCISSOR_TEST
'      GL_COLOR_MATERIAL
'      GL_NORMALIZE
'      GL_AUTO_NORMAL
'      GL_VERTEX_ARRAY
'      GL_NORMAL_ARRAY
'      GL_COLOR_ARRAY
'      GL_INDEX_ARRAY
'      GL_TEXTURE_COORD_ARRAY
'      GL_EDGE_FLAG_ARRAY
'      GL_POLYGON_OFFSET_POINT
'      GL_POLYGON_OFFSET_LINE
'      GL_POLYGON_OFFSET_FILL

' ErrorCode
Public Const GL_NO_ERROR = 0
Public Const GL_INVALID_ENUM = &H500
Public Const GL_INVALID_VALUE = &H501
Public Const GL_INVALID_OPERATION = &H502
Public Const GL_STACK_OVERFLOW = &H503
Public Const GL_STACK_UNDERFLOW = &H504
Public Const GL_OUT_OF_MEMORY = &H505

' FeedBackMode
Public Const GL_2D = &H600
Public Const GL_3D = &H601
Public Const GL_3D_COLOR = &H602
Public Const GL_3D_COLOR_TEXTURE = &H603
Public Const GL_4D_COLOR_TEXTURE = &H604

' FeedBackToken
Public Const GL_PASS_THROUGH_TOKEN = &H700
Public Const GL_POINT_TOKEN = &H701
Public Const GL_LINE_TOKEN = &H702
Public Const GL_POLYGON_TOKEN = &H703
Public Const GL_BITMAP_TOKEN = &H704
Public Const GL_DRAW_PIXEL_TOKEN = &H705
Public Const GL_COPY_PIXEL_TOKEN = &H706
Public Const GL_LINE_RESET_TOKEN = &H707

' FogMode
'      GL_LINEAR
Public Const GL_EXP = &H800
Public Const GL_EXP2 = &H801

' FogParameter
'      GL_FOG_COLOR
'      GL_FOG_DENSITY
'      GL_FOG_END
'      GL_FOG_INDEX
'      GL_FOG_MODE
'      GL_FOG_START

' FrontFaceDirection
Public Const GL_CW = &H900
Public Const GL_CCW = &H901

' GetMapTarget
Public Const GL_COEFF = &HA00
Public Const GL_ORDER = &HA01
Public Const GL_DOMAIN = &HA02

' GetPixelMap
'      GL_PIXEL_MAP_I_TO_I
'      GL_PIXEL_MAP_S_TO_S
'      GL_PIXEL_MAP_I_TO_R
'      GL_PIXEL_MAP_I_TO_G
'      GL_PIXEL_MAP_I_TO_B
'      GL_PIXEL_MAP_I_TO_A
'      GL_PIXEL_MAP_R_TO_R
'      GL_PIXEL_MAP_G_TO_G
'      GL_PIXEL_MAP_B_TO_B
'      GL_PIXEL_MAP_A_TO_A

' GetPointerTarget
'      GL_VERTEX_ARRAY_POINTER
'      GL_NORMAL_ARRAY_POINTER
'      GL_COLOR_ARRAY_POINTER
'      GL_INDEX_ARRAY_POINTER
'      GL_TEXTURE_COORD_ARRAY_POINTER
'      GL_EDGE_FLAG_ARRAY_POINTER

' GetTarget
Public Const GL_CURRENT_COLOR = &HB00
Public Const GL_CURRENT_INDEX = &HB01
Public Const GL_CURRENT_NORMAL = &HB02
Public Const GL_CURRENT_TEXTURE_COORDS = &HB03
Public Const GL_CURRENT_RASTER_COLOR = &HB04
Public Const GL_CURRENT_RASTER_INDEX = &HB05
Public Const GL_CURRENT_RASTER_TEXTURE_COORDS = &HB06
Public Const GL_CURRENT_RASTER_POSITION = &HB07
Public Const GL_CURRENT_RASTER_POSITION_VALID = &HB08
Public Const GL_CURRENT_RASTER_DISTANCE = &HB09
Public Const GL_POINT_SMOOTH = &HB10
Public Const GL_POINT_SIZE = &HB11
Public Const GL_POINT_SIZE_RANGE = &HB12
Public Const GL_POINT_SIZE_GRANULARITY = &HB13
Public Const GL_LINE_SMOOTH = &HB20
Public Const GL_LINE_WIDTH = &HB21
Public Const GL_LINE_WIDTH_RANGE = &HB22
Public Const GL_LINE_WIDTH_GRANULARITY = &HB23
Public Const GL_LINE_STIPPLE = &HB24
Public Const GL_LINE_STIPPLE_PATTERN = &HB25
Public Const GL_LINE_STIPPLE_REPEAT = &HB26
Public Const GL_LIST_MODE = &HB30
Public Const GL_MAX_LIST_NESTING = &HB31
Public Const GL_LIST_BASE = &HB32
Public Const GL_LIST_INDEX = &HB33
Public Const GL_POLYGON_MODE = &HB40
Public Const GL_POLYGON_SMOOTH = &HB41
Public Const GL_POLYGON_STIPPLE = &HB42
Public Const GL_EDGE_FLAG = &HB43
Public Const GL_CULL_FACE = &HB44
Public Const GL_CULL_FACE_MODE = &HB45
Public Const GL_FRONT_FACE = &HB46
Public Const GL_LIGHTING = &HB50
Public Const GL_LIGHT_MODEL_LOCAL_VIEWER = &HB51
Public Const GL_LIGHT_MODEL_TWO_SIDE = &HB52
Public Const GL_LIGHT_MODEL_AMBIENT = &HB53
Public Const GL_SHADE_MODEL = &HB54
Public Const GL_COLOR_MATERIAL_FACE = &HB55
Public Const GL_COLOR_MATERIAL_PARAMETER = &HB56
Public Const GL_COLOR_MATERIAL = &HB57
Public Const GL_FOG = &HB60
Public Const GL_FOG_INDEX = &HB61
Public Const GL_FOG_DENSITY = &HB62
Public Const GL_FOG_START = &HB63
Public Const GL_FOG_END = &HB64
Public Const GL_FOG_MODE = &HB65
Public Const GL_FOG_COLOR = &HB66
Public Const GL_DEPTH_RANGE = &HB70
Public Const GL_DEPTH_TEST = &HB71
Public Const GL_DEPTH_WRITEMASK = &HB72
Public Const GL_DEPTH_CLEAR_VALUE = &HB73
Public Const GL_DEPTH_FUNC = &HB74
Public Const GL_ACCUM_CLEAR_VALUE = &HB80
Public Const GL_STENCIL_TEST = &HB90
Public Const GL_STENCIL_CLEAR_VALUE = &HB91
Public Const GL_STENCIL_FUNC = &HB92
Public Const GL_STENCIL_VALUE_MASK = &HB93
Public Const GL_STENCIL_FAIL = &HB94
Public Const GL_STENCIL_PASS_DEPTH_FAIL = &HB95
Public Const GL_STENCIL_PASS_DEPTH_PASS = &HB96
Public Const GL_STENCIL_REF = &HB97
Public Const GL_STENCIL_WRITEMASK = &HB98
Public Const GL_MATRIX_MODE = &HBA0
Public Const GL_NORMALIZE = &HBA1
Public Const GL_VIEWPORT = &HBA2
Public Const GL_MODELVIEW_STACK_DEPTH = &HBA3
Public Const GL_PROJECTION_STACK_DEPTH = &HBA4
Public Const GL_TEXTURE_STACK_DEPTH = &HBA5
Public Const GL_MODELVIEW_MATRIX = &HBA6
Public Const GL_PROJECTION_MATRIX = &HBA7
Public Const GL_TEXTURE_MATRIX = &HBA8
Public Const GL_ATTRIB_STACK_DEPTH = &HBB0
Public Const GL_CLIENT_ATTRIB_STACK_DEPTH = &HBB1
Public Const GL_ALPHA_TEST = &HBC0
Public Const GL_ALPHA_TEST_FUNC = &HBC1
Public Const GL_ALPHA_TEST_REF = &HBC2
Public Const GL_DITHER = &HBD0
Public Const GL_BLEND_DST = &HBE0
Public Const GL_BLEND_SRC = &HBE1
Public Const GL_BLEND = &HBE2
Public Const GL_LOGIC_OP_MODE = &HBF0
Public Const GL_INDEX_LOGIC_OP = &HBF1
Public Const GL_COLOR_LOGIC_OP = &HBF2
Public Const GL_AUX_BUFFERS = &HC00
Public Const GL_DRAW_BUFFER = &HC01
Public Const GL_READ_BUFFER = &HC02
Public Const GL_SCISSOR_BOX = &HC10
Public Const GL_SCISSOR_TEST = &HC11
Public Const GL_INDEX_CLEAR_VALUE = &HC20
Public Const GL_INDEX_WRITEMASK = &HC21
Public Const GL_COLOR_CLEAR_VALUE = &HC22
Public Const GL_COLOR_WRITEMASK = &HC23
Public Const GL_INDEX_MODE = &HC30
Public Const GL_RGBA_MODE = &HC31
Public Const GL_DOUBLEBUFFER = &HC32
Public Const GL_STEREO = &HC33
Public Const GL_RENDER_MODE = &HC40
Public Const GL_PERSPECTIVE_CORRECTION_HINT = &HC50
Public Const GL_POINT_SMOOTH_HINT = &HC51
Public Const GL_LINE_SMOOTH_HINT = &HC52
Public Const GL_POLYGON_SMOOTH_HINT = &HC53
Public Const GL_FOG_HINT = &HC54
Public Const GL_TEXTURE_GEN_S = &HC60
Public Const GL_TEXTURE_GEN_T = &HC61
Public Const GL_TEXTURE_GEN_R = &HC62
Public Const GL_TEXTURE_GEN_Q = &HC63
Public Const GL_PIXEL_MAP_I_TO_I = &HC70
Public Const GL_PIXEL_MAP_S_TO_S = &HC71
Public Const GL_PIXEL_MAP_I_TO_R = &HC72
Public Const GL_PIXEL_MAP_I_TO_G = &HC73
Public Const GL_PIXEL_MAP_I_TO_B = &HC74
Public Const GL_PIXEL_MAP_I_TO_A = &HC75
Public Const GL_PIXEL_MAP_R_TO_R = &HC76
Public Const GL_PIXEL_MAP_G_TO_G = &HC77
Public Const GL_PIXEL_MAP_B_TO_B = &HC78
Public Const GL_PIXEL_MAP_A_TO_A = &HC79
Public Const GL_PIXEL_MAP_I_TO_I_SIZE = &HCB0
Public Const GL_PIXEL_MAP_S_TO_S_SIZE = &HCB1
Public Const GL_PIXEL_MAP_I_TO_R_SIZE = &HCB2
Public Const GL_PIXEL_MAP_I_TO_G_SIZE = &HCB3
Public Const GL_PIXEL_MAP_I_TO_B_SIZE = &HCB4
Public Const GL_PIXEL_MAP_I_TO_A_SIZE = &HCB5
Public Const GL_PIXEL_MAP_R_TO_R_SIZE = &HCB6
Public Const GL_PIXEL_MAP_G_TO_G_SIZE = &HCB7
Public Const GL_PIXEL_MAP_B_TO_B_SIZE = &HCB8
Public Const GL_PIXEL_MAP_A_TO_A_SIZE = &HCB9
Public Const GL_UNPACK_SWAP_BYTES = &HCF0
Public Const GL_UNPACK_LSB_FIRST = &HCF1
Public Const GL_UNPACK_ROW_LENGTH = &HCF2
Public Const GL_UNPACK_SKIP_ROWS = &HCF3
Public Const GL_UNPACK_SKIP_PIXELS = &HCF4
Public Const GL_UNPACK_ALIGNMENT = &HCF5
Public Const GL_PACK_SWAP_BYTES = &HD00
Public Const GL_PACK_LSB_FIRST = &HD01
Public Const GL_PACK_ROW_LENGTH = &HD02
Public Const GL_PACK_SKIP_ROWS = &HD03
Public Const GL_PACK_SKIP_PIXELS = &HD04
Public Const GL_PACK_ALIGNMENT = &HD05
Public Const GL_MAP_COLOR = &HD10
Public Const GL_MAP_STENCIL = &HD11
Public Const GL_INDEX_SHIFT = &HD12
Public Const GL_INDEX_OFFSET = &HD13
Public Const GL_RED_SCALE = &HD14
Public Const GL_RED_BIAS = &HD15
Public Const GL_ZOOM_X = &HD16
Public Const GL_ZOOM_Y = &HD17
Public Const GL_GREEN_SCALE = &HD18
Public Const GL_GREEN_BIAS = &HD19
Public Const GL_BLUE_SCALE = &HD1A
Public Const GL_BLUE_BIAS = &HD1B
Public Const GL_ALPHA_SCALE = &HD1C
Public Const GL_ALPHA_BIAS = &HD1D
Public Const GL_DEPTH_SCALE = &HD1E
Public Const GL_DEPTH_BIAS = &HD1F
Public Const GL_MAX_EVAL_ORDER = &HD30
Public Const GL_MAX_LIGHTS = &HD31
Public Const GL_MAX_CLIP_PLANES = &HD32
Public Const GL_MAX_TEXTURE_SIZE = &HD33
Public Const GL_MAX_PIXEL_MAP_TABLE = &HD34
Public Const GL_MAX_ATTRIB_STACK_DEPTH = &HD35
Public Const GL_MAX_MODELVIEW_STACK_DEPTH = &HD36
Public Const GL_MAX_NAME_STACK_DEPTH = &HD37
Public Const GL_MAX_PROJECTION_STACK_DEPTH = &HD38
Public Const GL_MAX_TEXTURE_STACK_DEPTH = &HD39
Public Const GL_MAX_VIEWPORT_DIMS = &HD3A
Public Const GL_MAX_CLIENT_ATTRIB_STACK_DEPTH = &HD3B
Public Const GL_SUBPIXEL_BITS = &HD50
Public Const GL_INDEX_BITS = &HD51
Public Const GL_RED_BITS = &HD52
Public Const GL_GREEN_BITS = &HD53
Public Const GL_BLUE_BITS = &HD54
Public Const GL_ALPHA_BITS = &HD55
Public Const GL_DEPTH_BITS = &HD56
Public Const GL_STENCIL_BITS = &HD57
Public Const GL_ACCUM_RED_BITS = &HD58
Public Const GL_ACCUM_GREEN_BITS = &HD59
Public Const GL_ACCUM_BLUE_BITS = &HD5A
Public Const GL_ACCUM_ALPHA_BITS = &HD5B
Public Const GL_NAME_STACK_DEPTH = &HD70
Public Const GL_AUTO_NORMAL = &HD80
Public Const GL_MAP1_COLOR_4 = &HD90
Public Const GL_MAP1_INDEX = &HD91
Public Const GL_MAP1_NORMAL = &HD92
Public Const GL_MAP1_TEXTURE_COORD_1 = &HD93
Public Const GL_MAP1_TEXTURE_COORD_2 = &HD94
Public Const GL_MAP1_TEXTURE_COORD_3 = &HD95
Public Const GL_MAP1_TEXTURE_COORD_4 = &HD96
Public Const GL_MAP1_VERTEX_3 = &HD97
Public Const GL_MAP1_VERTEX_4 = &HD98
Public Const GL_MAP2_COLOR_4 = &HDB0
Public Const GL_MAP2_INDEX = &HDB1
Public Const GL_MAP2_NORMAL = &HDB2
Public Const GL_MAP2_TEXTURE_COORD_1 = &HDB3
Public Const GL_MAP2_TEXTURE_COORD_2 = &HDB4
Public Const GL_MAP2_TEXTURE_COORD_3 = &HDB5
Public Const GL_MAP2_TEXTURE_COORD_4 = &HDB6
Public Const GL_MAP2_VERTEX_3 = &HDB7
Public Const GL_MAP2_VERTEX_4 = &HDB8
Public Const GL_MAP1_GRID_DOMAIN = &HDD0
Public Const GL_MAP1_GRID_SEGMENTS = &HDD1
Public Const GL_MAP2_GRID_DOMAIN = &HDD2
Public Const GL_MAP2_GRID_SEGMENTS = &HDD3
Public Const GL_TEXTURE_1D = &HDE0
Public Const GL_TEXTURE_2D = &HDE1
Public Const GL_FEEDBACK_BUFFER_POINTER = &HDF0
Public Const GL_FEEDBACK_BUFFER_SIZE = &HDF1
Public Const GL_FEEDBACK_BUFFER_TYPE = &HDF2
Public Const GL_SELECTION_BUFFER_POINTER = &HDF3
Public Const GL_SELECTION_BUFFER_SIZE = &HDF4
'      GL_TEXTURE_BINDING_1D
'      GL_TEXTURE_BINDING_2D
'      GL_VERTEX_ARRAY
'      GL_NORMAL_ARRAY
'      GL_COLOR_ARRAY
'      GL_INDEX_ARRAY
'      GL_TEXTURE_COORD_ARRAY
'      GL_EDGE_FLAG_ARRAY
'      GL_VERTEX_ARRAY_SIZE
'      GL_VERTEX_ARRAY_TYPE
'      GL_VERTEX_ARRAY_STRIDE
'      GL_NORMAL_ARRAY_TYPE
'      GL_NORMAL_ARRAY_STRIDE
'      GL_COLOR_ARRAY_SIZE
'      GL_COLOR_ARRAY_TYPE
'      GL_COLOR_ARRAY_STRIDE
'      GL_INDEX_ARRAY_TYPE
'      GL_INDEX_ARRAY_STRIDE
'      GL_TEXTURE_COORD_ARRAY_SIZE
'      GL_TEXTURE_COORD_ARRAY_TYPE
'      GL_TEXTURE_COORD_ARRAY_STRIDE
'      GL_EDGE_FLAG_ARRAY_STRIDE
'      GL_POLYGON_OFFSET_FACTOR
'      GL_POLYGON_OFFSET_UNITS

' GetTextureParameter
'      GL_TEXTURE_MAG_FILTER
'      GL_TEXTURE_MIN_FILTER
'      GL_TEXTURE_WRAP_S
'      GL_TEXTURE_WRAP_T
Public Const GL_TEXTURE_WIDTH = &H1000
Public Const GL_TEXTURE_HEIGHT = &H1001
Public Const GL_TEXTURE_INTERNAL_FORMAT = &H1003
Public Const GL_TEXTURE_BORDER_COLOR = &H1004
Public Const GL_TEXTURE_BORDER = &H1005
'      GL_TEXTURE_RED_SIZE
'      GL_TEXTURE_GREEN_SIZE
'      GL_TEXTURE_BLUE_SIZE
'      GL_TEXTURE_ALPHA_SIZE
'      GL_TEXTURE_LUMINANCE_SIZE
'      GL_TEXTURE_INTENSITY_SIZE
'      GL_TEXTURE_PRIORITY
'      GL_TEXTURE_RESIDENT

' HintMode
Public Const GL_DONT_CARE = &H1100
Public Const GL_FASTEST = &H1101
Public Const GL_NICEST = &H1102

' HintTarget
'      GL_PERSPECTIVE_CORRECTION_HINT
'      GL_POINT_SMOOTH_HINT
'      GL_LINE_SMOOTH_HINT
'      GL_POLYGON_SMOOTH_HINT
'      GL_FOG_HINT

' IndexPointerType
'      GL_SHORT
'      GL_INT
'      GL_FLOAT
'      GL_DOUBLE

' LightModelParameter
'      GL_LIGHT_MODEL_AMBIENT
'      GL_LIGHT_MODEL_LOCAL_VIEWER
'      GL_LIGHT_MODEL_TWO_SIDE

' LightName
Public Const GL_LIGHT0 = &H4000
Public Const GL_LIGHT1 = &H4001
Public Const GL_LIGHT2 = &H4002
Public Const GL_LIGHT3 = &H4003
Public Const GL_LIGHT4 = &H4004
Public Const GL_LIGHT5 = &H4005
Public Const GL_LIGHT6 = &H4006
Public Const GL_LIGHT7 = &H4007

' LightParameter
Public Const GL_AMBIENT = &H1200
Public Const GL_DIFFUSE = &H1201
Public Const GL_SPECULAR = &H1202
Public Const GL_POSITION = &H1203
Public Const GL_SPOT_DIRECTION = &H1204
Public Const GL_SPOT_EXPONENT = &H1205
Public Const GL_SPOT_CUTOFF = &H1206
Public Const GL_CONSTANT_ATTENUATION = &H1207
Public Const GL_LINEAR_ATTENUATION = &H1208
Public Const GL_QUADRATIC_ATTENUATION = &H1209

' InterleavedArrays
'      GL_V2F
'      GL_V3F
'      GL_C4UB_V2F
'      GL_C4UB_V3F
'      GL_C3F_V3F
'      GL_N3F_V3F
'      GL_C4F_N3F_V3F
'      GL_T2F_V3F
'      GL_T4F_V4F
'      GL_T2F_C4UB_V3F
'      GL_T2F_C3F_V3F
'      GL_T2F_N3F_V3F
'      GL_T2F_C4F_N3F_V3F
'      GL_T4F_C4F_N3F_V4F

' ListMode
Public Const GL_COMPILE = &H1300
Public Const GL_COMPILE_AND_EXECUTE = &H1301

' ListNameType
'      GL_BYTE
'      GL_UNSIGNED_BYTE
'      GL_SHORT
'      GL_UNSIGNED_SHORT
'      GL_INT
'      GL_UNSIGNED_INT
'      GL_FLOAT
'      GL_2_BYTES
'      GL_3_BYTES
'      GL_4_BYTES

' LogicOp
Public Const GL_CLEAR = &H1500
Public Const GL_AND = &H1501
Public Const GL_AND_REVERSE = &H1502
Public Const GL_COPY = &H1503
Public Const GL_AND_INVERTED = &H1504
Public Const GL_NOOP = &H1505
Public Const GL_XOR = &H1506
Public Const GL_OR = &H1507
Public Const GL_NOR = &H1508
Public Const GL_EQUIV = &H1509
Public Const GL_INVERT = &H150A
Public Const GL_OR_REVERSE = &H150B
Public Const GL_COPY_INVERTED = &H150C
Public Const GL_OR_INVERTED = &H150D
Public Const GL_NAND = &H150E
Public Const GL_SET = &H150F

' MapTarget
'      GL_MAP1_COLOR_4
'      GL_MAP1_INDEX
'      GL_MAP1_NORMAL
'      GL_MAP1_TEXTURE_COORD_1
'      GL_MAP1_TEXTURE_COORD_2
'      GL_MAP1_TEXTURE_COORD_3
'      GL_MAP1_TEXTURE_COORD_4
'      GL_MAP1_VERTEX_3
'      GL_MAP1_VERTEX_4
'      GL_MAP2_COLOR_4
'      GL_MAP2_INDEX
'      GL_MAP2_NORMAL
'      GL_MAP2_TEXTURE_COORD_1
'      GL_MAP2_TEXTURE_COORD_2
'      GL_MAP2_TEXTURE_COORD_3
'      GL_MAP2_TEXTURE_COORD_4
'      GL_MAP2_VERTEX_3
'      GL_MAP2_VERTEX_4

' MaterialFace
'      GL_FRONT
'      GL_BACK
'      GL_FRONT_AND_BACK

' MaterialParameter
Public Const GL_EMISSION = &H1600
Public Const GL_SHININESS = &H1601
Public Const GL_AMBIENT_AND_DIFFUSE = &H1602
Public Const GL_COLOR_INDEXES = &H1603
'      GL_AMBIENT
'      GL_DIFFUSE
'      GL_SPECULAR

' MatrixMode
Public Const GL_MODELVIEW = &H1700
Public Const GL_PROJECTION = &H1701
Public Const GL_TEXTURE = &H1702

' MeshMode1
'      GL_POINT
'      GL_LINE

' MeshMode2
'      GL_POINT
'      GL_LINE
'      GL_FILL

' NormalPointerType
'      GL_BYTE
'      GL_SHORT
'      GL_INT
'      GL_FLOAT
'      GL_DOUBLE

' PixelCopyType
Public Const GL_COLOR = &H1800
Public Const GL_DEPTH = &H1801
Public Const GL_STENCIL = &H1802

' PixelFormat
Public Const GL_COLOR_INDEX = &H1900
Public Const GL_STENCIL_INDEX = &H1901
Public Const GL_DEPTH_COMPONENT = &H1902
Public Const GL_RED = &H1903
Public Const GL_GREEN = &H1904
Public Const GL_BLUE = &H1905
Public Const GL_ALPHA = &H1906
Public Const GL_RGB = &H1907
Public Const GL_RGBA = &H1908
Public Const GL_LUMINANCE = &H1909
Public Const GL_LUMINANCE_ALPHA = &H190A

' PixelMap
'      GL_PIXEL_MAP_I_TO_I
'      GL_PIXEL_MAP_S_TO_S
'      GL_PIXEL_MAP_I_TO_R
'      GL_PIXEL_MAP_I_TO_G
'      GL_PIXEL_MAP_I_TO_B
'      GL_PIXEL_MAP_I_TO_A
'      GL_PIXEL_MAP_R_TO_R
'      GL_PIXEL_MAP_G_TO_G
'      GL_PIXEL_MAP_B_TO_B
'      GL_PIXEL_MAP_A_TO_A

' PixelStore
'      GL_UNPACK_SWAP_BYTES
'      GL_UNPACK_LSB_FIRST
'      GL_UNPACK_ROW_LENGTH
'      GL_UNPACK_SKIP_ROWS
'      GL_UNPACK_SKIP_PIXELS
'      GL_UNPACK_ALIGNMENT
'      GL_PACK_SWAP_BYTES
'      GL_PACK_LSB_FIRST
'      GL_PACK_ROW_LENGTH
'      GL_PACK_SKIP_ROWS
'      GL_PACK_SKIP_PIXELS
'      GL_PACK_ALIGNMENT

' PixelTransfer
'      GL_MAP_COLOR
'      GL_MAP_STENCIL
'      GL_INDEX_SHIFT
'      GL_INDEX_OFFSET
'      GL_RED_SCALE
'      GL_RED_BIAS
'      GL_GREEN_SCALE
'      GL_GREEN_BIAS
'      GL_BLUE_SCALE
'      GL_BLUE_BIAS
'      GL_ALPHA_SCALE
'      GL_ALPHA_BIAS
'      GL_DEPTH_SCALE
'      GL_DEPTH_BIAS

' PixelType
Public Const GL_BITMAP = &H1A00
'      GL_BYTE
'      GL_UNSIGNED_BYTE
'      GL_SHORT
'      GL_UNSIGNED_SHORT
'      GL_INT
'      GL_UNSIGNED_INT
'      GL_FLOAT

' PolygonMode
Public Const GL_POINT = &H1B00
Public Const GL_LINE = &H1B01
Public Const GL_FILL = &H1B02

' ReadBufferMode
'      GL_FRONT_LEFT
'      GL_FRONT_RIGHT
'      GL_BACK_LEFT
'      GL_BACK_RIGHT
'      GL_FRONT
'      GL_BACK
'      GL_LEFT
'      GL_RIGHT
'      GL_AUX0
'      GL_AUX1
'      GL_AUX2
'      GL_AUX3

' RenderingMode
Public Const GL_RENDER = &H1C00
Public Const GL_FEEDBACK = &H1C01
Public Const GL_SELECT = &H1C02

' ShadingModel
Public Const GL_FLAT = &H1D00
Public Const GL_SMOOTH = &H1D01

' StencilFunction
'      GL_NEVER
'      GL_LESS
'      GL_EQUAL
'      GL_LEQUAL
'      GL_GREATER
'      GL_NOTEQUAL
'      GL_GEQUAL
'      GL_ALWAYS

' StencilOp
'      GL_ZERO
Public Const GL_KEEP = &H1E00
Public Const GL_REPLACE = &H1E01
Public Const GL_INCR = &H1E02
Public Const GL_DECR = &H1E03
'      GL_INVERT

' StringName
Public Const GL_VENDOR = &H1F00
Public Const GL_RENDERER = &H1F01
Public Const GL_VERSION = &H1F02
Public Const GL_EXTENSIONS = &H1F03

' TextureCoordName
Public Const GL_S = &H2000
Public Const GL_T = &H2001
Public Const GL_R = &H2002
Public Const GL_Q = &H2003

' TexCoordPointerType
'      GL_SHORT
'      GL_INT
'      GL_FLOAT
'      GL_DOUBLE

' TextureEnvMode
Public Const GL_MODULATE = &H2100
Public Const GL_DECAL = &H2101
'      GL_BLEND
'      GL_REPLACE

' TextureEnvParameter
Public Const GL_TEXTURE_ENV_MODE = &H2200
Public Const GL_TEXTURE_ENV_COLOR = &H2201

' TextureEnvTarget
Public Const GL_TEXTURE_ENV = &H2300

' TextureGenMode
Public Const GL_EYE_LINEAR = &H2400
Public Const GL_OBJECT_LINEAR = &H2401
Public Const GL_SPHERE_MAP = &H2402

' TextureGenParameter
Public Const GL_TEXTURE_GEN_MODE = &H2500
Public Const GL_OBJECT_PLANE = &H2501
Public Const GL_EYE_PLANE = &H2502

' TextureMagFilter
Public Const GL_NEAREST = &H2600
Public Const GL_LINEAR = &H2601

' TextureMinFilter
'      GL_NEAREST
'      GL_LINEAR
Public Const GL_NEAREST_MIPMAP_NEAREST = &H2700
Public Const GL_LINEAR_MIPMAP_NEAREST = &H2701
Public Const GL_NEAREST_MIPMAP_LINEAR = &H2702
Public Const GL_LINEAR_MIPMAP_LINEAR = &H2703

' TextureParameterName
Public Const GL_TEXTURE_MAG_FILTER = &H2800
Public Const GL_TEXTURE_MIN_FILTER = &H2801
Public Const GL_TEXTURE_WRAP_S = &H2802
Public Const GL_TEXTURE_WRAP_T = &H2803
'      GL_TEXTURE_BORDER_COLOR
'      GL_TEXTURE_PRIORITY

' TextureTarget
'      GL_TEXTURE_1D
'      GL_TEXTURE_2D
'      GL_PROXY_TEXTURE_1D
'      GL_PROXY_TEXTURE_2D

' TextureWrapMode
Public Const GL_CLAMP = &H2900
Public Const GL_REPEAT = &H2901

' VertexPointerType
'      GL_SHORT
'      GL_INT
'      GL_FLOAT
'      GL_DOUBLE

' ClientAttribMask
Public Const GL_CLIENT_PIXEL_STORE_BIT = &H1
Public Const GL_CLIENT_VERTEX_ARRAY_BIT = &H2
Public Const GL_CLIENT_ALL_ATTRIB_BITS = &HFFFFFFFF

' polygon_offset
Public Const GL_POLYGON_OFFSET_FACTOR = &H8038
Public Const GL_POLYGON_OFFSET_UNITS = &H2A00
Public Const GL_POLYGON_OFFSET_POINT = &H2A01
Public Const GL_POLYGON_OFFSET_LINE = &H2A02
Public Const GL_POLYGON_OFFSET_FILL = &H8037

' texture
Public Const GL_ALPHA4 = &H803B
Public Const GL_ALPHA8 = &H803C
Public Const GL_ALPHA12 = &H803D
Public Const GL_ALPHA16 = &H803E
Public Const GL_LUMINANCE4 = &H803F
Public Const GL_LUMINANCE8 = &H8040
Public Const GL_LUMINANCE12 = &H8041
Public Const GL_LUMINANCE16 = &H8042
Public Const GL_LUMINANCE4_ALPHA4 = &H8043
Public Const GL_LUMINANCE6_ALPHA2 = &H8044
Public Const GL_LUMINANCE8_ALPHA8 = &H8045
Public Const GL_LUMINANCE12_ALPHA4 = &H8046
Public Const GL_LUMINANCE12_ALPHA12 = &H8047
Public Const GL_LUMINANCE16_ALPHA16 = &H8048
Public Const GL_INTENSITY = &H8049
Public Const GL_INTENSITY4 = &H804A
Public Const GL_INTENSITY8 = &H804B
Public Const GL_INTENSITY12 = &H804C
Public Const GL_INTENSITY16 = &H804D
Public Const GL_R3_G3_B2 = &H2A10
Public Const GL_RGB4 = &H804F
Public Const GL_RGB5 = &H8050
Public Const GL_RGB8 = &H8051
Public Const GL_RGB10 = &H8052
Public Const GL_RGB12 = &H8053
Public Const GL_RGB16 = &H8054
Public Const GL_RGBA2 = &H8055
Public Const GL_RGBA4 = &H8056
Public Const GL_RGB5_A1 = &H8057
Public Const GL_RGBA8 = &H8058
Public Const GL_RGB10_A2 = &H8059
Public Const GL_RGBA12 = &H805A
Public Const GL_RGBA16 = &H805B
Public Const GL_TEXTURE_RED_SIZE = &H805C
Public Const GL_TEXTURE_GREEN_SIZE = &H805D
Public Const GL_TEXTURE_BLUE_SIZE = &H805E
Public Const GL_TEXTURE_ALPHA_SIZE = &H805F
Public Const GL_TEXTURE_LUMINANCE_SIZE = &H8060
Public Const GL_TEXTURE_INTENSITY_SIZE = &H8061
Public Const GL_PROXY_TEXTURE_1D = &H8063
Public Const GL_PROXY_TEXTURE_2D = &H8064

' texture_object
Public Const GL_TEXTURE_PRIORITY = &H8066
Public Const GL_TEXTURE_RESIDENT = &H8067
Public Const GL_TEXTURE_BINDING_1D = &H8068
Public Const GL_TEXTURE_BINDING_2D = &H8069

' vertex_array
Public Const GL_VERTEX_ARRAY = &H8074
Public Const GL_NORMAL_ARRAY = &H8075
Public Const GL_COLOR_ARRAY = &H8076
Public Const GL_INDEX_ARRAY = &H8077
Public Const GL_TEXTURE_COORD_ARRAY = &H8078
Public Const GL_EDGE_FLAG_ARRAY = &H8079
Public Const GL_VERTEX_ARRAY_SIZE = &H807A
Public Const GL_VERTEX_ARRAY_TYPE = &H807B
Public Const GL_VERTEX_ARRAY_STRIDE = &H807C
Public Const GL_NORMAL_ARRAY_TYPE = &H807E
Public Const GL_NORMAL_ARRAY_STRIDE = &H807F
Public Const GL_COLOR_ARRAY_SIZE = &H8081
Public Const GL_COLOR_ARRAY_TYPE = &H8082
Public Const GL_COLOR_ARRAY_STRIDE = &H8083
Public Const GL_INDEX_ARRAY_TYPE = &H8085
Public Const GL_INDEX_ARRAY_STRIDE = &H8086
Public Const GL_TEXTURE_COORD_ARRAY_SIZE = &H8088
Public Const GL_TEXTURE_COORD_ARRAY_TYPE = &H8089
Public Const GL_TEXTURE_COORD_ARRAY_STRIDE = &H808A
Public Const GL_EDGE_FLAG_ARRAY_STRIDE = &H808C
Public Const GL_VERTEX_ARRAY_POINTER = &H808E
Public Const GL_NORMAL_ARRAY_POINTER = &H808F
Public Const GL_COLOR_ARRAY_POINTER = &H8090
Public Const GL_INDEX_ARRAY_POINTER = &H8091
Public Const GL_TEXTURE_COORD_ARRAY_POINTER = &H8092
Public Const GL_EDGE_FLAG_ARRAY_POINTER = &H8093
Public Const GL_V2F = &H2A20
Public Const GL_V3F = &H2A21
Public Const GL_C4UB_V2F = &H2A22
Public Const GL_C4UB_V3F = &H2A23
Public Const GL_C3F_V3F = &H2A24
Public Const GL_N3F_V3F = &H2A25
Public Const GL_C4F_N3F_V3F = &H2A26
Public Const GL_T2F_V3F = &H2A27
Public Const GL_T4F_V4F = &H2A28
Public Const GL_T2F_C4UB_V3F = &H2A29
Public Const GL_T2F_C3F_V3F = &H2A2A
Public Const GL_T2F_N3F_V3F = &H2A2B
Public Const GL_T2F_C4F_N3F_V3F = &H2A2C
Public Const GL_T4F_C4F_N3F_V4F = &H2A2D

' Extensions
Public Const GL_EXT_vertex_array = 1
Public Const GL_WIN_swap_hint = 1
Public Const GL_EXT_bgra = 1
Public Const GL_EXT_paletted_texture = 1

' EXT_vertex_array
Public Const GL_VERTEX_ARRAY_EXT = &H8074
Public Const GL_NORMAL_ARRAY_EXT = &H8075
Public Const GL_COLOR_ARRAY_EXT = &H8076
Public Const GL_INDEX_ARRAY_EXT = &H8077
Public Const GL_TEXTURE_COORD_ARRAY_EXT = &H8078
Public Const GL_EDGE_FLAG_ARRAY_EXT = &H8079
Public Const GL_VERTEX_ARRAY_SIZE_EXT = &H807A
Public Const GL_VERTEX_ARRAY_TYPE_EXT = &H807B
Public Const GL_VERTEX_ARRAY_STRIDE_EXT = &H807C
Public Const GL_VERTEX_ARRAY_COUNT_EXT = &H807D
Public Const GL_NORMAL_ARRAY_TYPE_EXT = &H807E
Public Const GL_NORMAL_ARRAY_STRIDE_EXT = &H807F
Public Const GL_NORMAL_ARRAY_COUNT_EXT = &H8080
Public Const GL_COLOR_ARRAY_SIZE_EXT = &H8081
Public Const GL_COLOR_ARRAY_TYPE_EXT = &H8082
Public Const GL_COLOR_ARRAY_STRIDE_EXT = &H8083
Public Const GL_COLOR_ARRAY_COUNT_EXT = &H8084
Public Const GL_INDEX_ARRAY_TYPE_EXT = &H8085
Public Const GL_INDEX_ARRAY_STRIDE_EXT = &H8086
Public Const GL_INDEX_ARRAY_COUNT_EXT = &H8087
Public Const GL_TEXTURE_COORD_ARRAY_SIZE_EXT = &H8088
Public Const GL_TEXTURE_COORD_ARRAY_TYPE_EXT = &H8089
Public Const GL_TEXTURE_COORD_ARRAY_STRIDE_EXT = &H808A
Public Const GL_TEXTURE_COORD_ARRAY_COUNT_EXT = &H808B
Public Const GL_EDGE_FLAG_ARRAY_STRIDE_EXT = &H808C
Public Const GL_EDGE_FLAG_ARRAY_COUNT_EXT = &H808D
Public Const GL_VERTEX_ARRAY_POINTER_EXT = &H808E
Public Const GL_NORMAL_ARRAY_POINTER_EXT = &H808F
Public Const GL_COLOR_ARRAY_POINTER_EXT = &H8090
Public Const GL_INDEX_ARRAY_POINTER_EXT = &H8091
Public Const GL_TEXTURE_COORD_ARRAY_POINTER_EXT = &H8092
Public Const GL_EDGE_FLAG_ARRAY_POINTER_EXT = &H8093
Public Const GL_DOUBLE_EXT = GL_DOUBLE

' EXT_bgra
Public Const GL_BGR_EXT = &H80E0
Public Const GL_BGRA_EXT = &H80E1

' EXT_paletted_texture

' These must match the GL_COLOR_TABLE_*_SGI enumerants
Public Const GL_COLOR_TABLE_FORMAT_EXT = &H80D8
Public Const GL_COLOR_TABLE_WIDTH_EXT = &H80D9
Public Const GL_COLOR_TABLE_RED_SIZE_EXT = &H80DA
Public Const GL_COLOR_TABLE_GREEN_SIZE_EXT = &H80DB
Public Const GL_COLOR_TABLE_BLUE_SIZE_EXT = &H80DC
Public Const GL_COLOR_TABLE_ALPHA_SIZE_EXT = &H80DD
Public Const GL_COLOR_TABLE_LUMINANCE_SIZE_EXT = &H80DE
Public Const GL_COLOR_TABLE_INTENSITY_SIZE_EXT = &H80DF

Public Const GL_COLOR_INDEX1_EXT = &H80E2
Public Const GL_COLOR_INDEX2_EXT = &H80E3
Public Const GL_COLOR_INDEX4_EXT = &H80E4
Public Const GL_COLOR_INDEX8_EXT = &H80E5
Public Const GL_COLOR_INDEX12_EXT = &H80E6
Public Const GL_COLOR_INDEX16_EXT = &H80E7

' For compatibility with OpenGL v1.0
Public Const GL_LOGIC_OP = GL_INDEX_LOGIC_OP
Public Const GL_TEXTURE_COMPONENTS = GL_TEXTURE_INTERNAL_FORMAT


'--------------------------------------------------------
'GLU
Public Const GLU_VERSION_1_1 = 1
Public Const GLU_VERSION_1_2 = 1
Public Const GLU_INVALID_ENUM = 100900
Public Const GLU_INVALID_VALUE = 100901
Public Const GLU_OUT_OF_MEMORY = 100902
Public Const GLU_INCOMPATIBLE_GL_VERSION = 100903
Public Const GLU_VERSION = 100800
Public Const GLU_EXTENSIONS = 100801
Public Const GLU_TRUE = GL_TRUE
Public Const GLU_FALSE = GL_FALSE
Public Const GLU_SMOOTH = 100000
Public Const GLU_FLAT = 100001
Public Const GLU_NONE = 100002
Public Const GLU_POINT = 100010
Public Const GLU_LINE = 100011
Public Const GLU_FILL = 100012
Public Const GLU_SILHOUETTE = 100013
Public Const GLU_OUTSIDE = 100020
Public Const GLU_INSIDE = 100021
Public Const GLU_TESS_MAX_COORD = 1E+150
Public Const GLU_TESS_WINDING_RULE = 100140
Public Const GLU_TESS_BOUNDARY_ONLY = 100141
Public Const GLU_TESS_TOLERANCE = 100142
Public Const GLU_TESS_WINDING_ODD = 100130
Public Const GLU_TESS_WINDING_NONZERO = 100131
Public Const GLU_TESS_WINDING_POSITIVE = 100132
Public Const GLU_TESS_WINDING_NEGATIVE = 100133
Public Const GLU_TESS_WINDING_ABS_GEQ_TWO = 100134
Public Const GLU_TESS_BEGIN = 100100
Public Const GLU_TESS_VERTEX = 100101
Public Const GLU_TESS_END = 100102
Public Const GLU_TESS_ERROR = 100103
Public Const GLU_TESS_EDGE_FLAG = 100104
Public Const GLU_TESS_COMBINE = 100105
Public Const GLU_TESS_BEGIN_DATA = 100106
Public Const GLU_TESS_VERTEX_DATA = 100107
Public Const GLU_TESS_END_DATA = 100108
Public Const GLU_TESS_ERROR_DATA = 100109
Public Const GLU_TESS_EDGE_FLAG_DATA = 100110
Public Const GLU_TESS_COMBINE_DATA = 100111
Public Const GLU_TESS_ERROR1 = 100151
Public Const GLU_TESS_ERROR2 = 100152
Public Const GLU_TESS_ERROR3 = 100153
Public Const GLU_TESS_ERROR4 = 100154
Public Const GLU_TESS_ERROR5 = 100155
Public Const GLU_TESS_ERROR6 = 100156
Public Const GLU_TESS_ERROR7 = 100157
Public Const GLU_TESS_ERROR8 = 100158
Public Const GLU_TESS_MISSING_BEGIN_POLYGON = GLU_TESS_ERROR1
Public Const GLU_TESS_MISSING_BEGIN_CONTOUR = GLU_TESS_ERROR2
Public Const GLU_TESS_MISSING_END_POLYGON = GLU_TESS_ERROR3
Public Const GLU_TESS_MISSING_END_CONTOUR = GLU_TESS_ERROR4
Public Const GLU_TESS_COORD_TOO_LARGE = GLU_TESS_ERROR5
Public Const GLU_TESS_NEED_COMBINE_CALLBACK = GLU_TESS_ERROR6
Public Const GLU_AUTO_LOAD_MATRIX = 100200
Public Const GLU_CULLING = 100201
Public Const GLU_SAMPLING_TOLERANCE = 100203
Public Const GLU_DISPLAY_MODE = 100204
Public Const GLU_PARAMETRIC_TOLERANCE = 100202
Public Const GLU_SAMPLING_METHOD = 100205
Public Const GLU_U_STEP = 100206
Public Const GLU_V_STEP = 100207
Public Const GLU_PATH_LENGTH = 100215
Public Const GLU_PARAMETRIC_ERROR = 100216
Public Const GLU_DOMAIN_DISTANCE = 100217
Public Const GLU_MAP1_TRIM_2 = 100210
Public Const GLU_MAP1_TRIM_3 = 100211
Public Const GLU_OUTLINE_POLYGON = 100240
Public Const GLU_OUTLINE_PATCH = 100241
Public Const GLU_NURBS_ERROR1 = 100251
Public Const GLU_NURBS_ERROR2 = 100252
Public Const GLU_NURBS_ERROR3 = 100253
Public Const GLU_NURBS_ERROR4 = 100254
Public Const GLU_NURBS_ERROR5 = 100255
Public Const GLU_NURBS_ERROR6 = 100256
Public Const GLU_NURBS_ERROR7 = 100257
Public Const GLU_NURBS_ERROR8 = 100258
Public Const GLU_NURBS_ERROR9 = 100259
Public Const GLU_NURBS_ERROR10 = 100260
Public Const GLU_NURBS_ERROR11 = 100261
Public Const GLU_NURBS_ERROR12 = 100262
Public Const GLU_NURBS_ERROR13 = 100263
Public Const GLU_NURBS_ERROR14 = 100264
Public Const GLU_NURBS_ERROR15 = 100265
Public Const GLU_NURBS_ERROR16 = 100266
Public Const GLU_NURBS_ERROR17 = 100267
Public Const GLU_NURBS_ERROR18 = 100268
Public Const GLU_NURBS_ERROR19 = 100269
Public Const GLU_NURBS_ERROR20 = 100270
Public Const GLU_NURBS_ERROR21 = 100271
Public Const GLU_NURBS_ERROR22 = 100272
Public Const GLU_NURBS_ERROR23 = 100273
Public Const GLU_NURBS_ERROR24 = 100274
Public Const GLU_NURBS_ERROR25 = 100275
Public Const GLU_NURBS_ERROR26 = 100276
Public Const GLU_NURBS_ERROR27 = 100277
Public Const GLU_NURBS_ERROR28 = 100278
Public Const GLU_NURBS_ERROR29 = 100279
Public Const GLU_NURBS_ERROR30 = 100280
Public Const GLU_NURBS_ERROR31 = 100281
Public Const GLU_NURBS_ERROR32 = 100282
Public Const GLU_NURBS_ERROR33 = 100283
Public Const GLU_NURBS_ERROR34 = 100284
Public Const GLU_NURBS_ERROR35 = 100285
Public Const GLU_NURBS_ERROR36 = 100286
Public Const GLU_NURBS_ERROR37 = 100287
Public Const GLU_CW = 100120
Public Const GLU_CCW = 100121
Public Const GLU_INTERIOR = 100122
Public Const GLU_EXTERIOR = 100123
Public Const GLU_UNKNOWN = 100124
Public Const GLU_BEGIN = GLU_TESS_BEGIN
Public Const GLU_VERTEX = GLU_TESS_VERTEX
Public Const GLU_END = GLU_TESS_END
Public Const GLU_ERROR = GLU_TESS_ERROR
Public Const GLU_EDGE_FLAG = GLU_TESS_EDGE_FLAG

'***********************************************************

Public Declare Sub glAccum Lib "OpenGL32.dll" (ByVal OP As Integer, ByVal Value As Single)
Public Declare Sub glAlphaFunc Lib "OpenGL32.dll" (ByVal Func As Integer, ByVal Ref As Single)
Public Declare Function glAreTexturesResident Lib "OpenGL32.dll" (ByVal N As Integer, ByRef Textures As Integer, ByRef Residences As Byte) As Byte
Public Declare Sub glArrayElement Lib "OpenGL32.dll" (ByVal I As Integer)
Public Declare Sub glBegin Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glBindTexture Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Texture As Integer)
Public Declare Sub glBitmap Lib "OpenGL32.dll" (ByVal Width As Integer, ByVal Height As Integer, ByVal Xorig As Single, ByVal Yorig As Single, ByVal Xmove As Single, ByVal Ymove As Single, ByRef Bitmap As Byte)
Public Declare Sub glBlendFunc Lib "OpenGL32.dll" (ByVal sFactor As Integer, ByVal dFactor As Integer)
Public Declare Sub glCallList Lib "OpenGL32.dll" (ByVal List As Integer)
Public Declare Sub glCallLists Lib "OpenGL32.dll" (ByVal N As Integer, ByVal iType As Integer, ByRef Lists As Any)
Public Declare Sub glClear Lib "OpenGL32.dll" (ByVal Mask As Integer)
Public Declare Sub glClearAccum Lib "OpenGL32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal Alpha As Single)
Public Declare Sub glClearColor Lib "OpenGL32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal Alpha As Single)
Public Declare Sub glClearDepth Lib "OpenGL32.dll" (ByVal Depth As Double)
Public Declare Sub glClearIndex Lib "OpenGL32.dll" (ByVal C As Single)
Public Declare Sub glClearStencil Lib "OpenGL32.dll" (ByVal S As Integer)
Public Declare Sub glClipPlane Lib "OpenGL32.dll" (ByVal Plane As Integer, ByRef Equation As Double)
Public Declare Sub glColor3b Lib "OpenGL32.dll" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
Public Declare Sub glColor3bv Lib "OpenGL32.dll" (ByRef V As Byte)
Public Declare Sub glColor3d Lib "OpenGL32.dll" (ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double)
Public Declare Sub glColor3dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glColor3f Lib "OpenGL32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single)
Public Declare Sub glColor3fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glColor3i Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
Public Declare Sub glColor3iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor3s Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
Public Declare Sub glColor3sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor3ub Lib "OpenGL32.dll" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
Public Declare Sub glColor3ubv Lib "OpenGL32.dll" (ByRef V As Byte)
Public Declare Sub glColor3ui Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
Public Declare Sub glColor3uiv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor3us Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
Public Declare Sub glColor3usv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor4b Lib "OpenGL32.dll" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal Alpha As Byte)
Public Declare Sub glColor4bv Lib "OpenGL32.dll" (ByRef V As Byte)
Public Declare Sub glColor4d Lib "OpenGL32.dll" (ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double, ByVal Alpha As Double)
Public Declare Sub glColor4dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glColor4f Lib "OpenGL32.dll" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal Alpha As Single)
Public Declare Sub glColor4fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glColor4i Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal Alpha As Integer)
Public Declare Sub glColor4iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor4s Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal Alpha As Integer)
Public Declare Sub glColor4sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor4ub Lib "OpenGL32.dll" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal Alpha As Byte)
Public Declare Sub glColor4ubv Lib "OpenGL32.dll" (ByRef V As Byte)
Public Declare Sub glColor4ui Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal Alpha As Integer)
Public Declare Sub glColor4uiv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColor4us Lib "OpenGL32.dll" (ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal Alpha As Integer)
Public Declare Sub glColor4usv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glColorMask Lib "OpenGL32.dll" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal Alpha As Byte)
Public Declare Sub glColorMaterial Lib "OpenGL32.dll" (ByVal face As Integer, ByVal Mode As Integer)
Public Declare Sub glColorPointer Lib "OpenGL32.dll" (ByVal Size As Integer, ByVal iType As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glCopyPixels Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal iType As Integer)
Public Declare Sub glCopyTexImage1D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal Internalformat As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Border As Integer)
Public Declare Sub glCopyTexImage2D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal Internalformat As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Border As Integer)
Public Declare Sub glCopyTexSubImage1D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal XOffset As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer)
Public Declare Sub glCopyTexSubImage2D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal XOffset As Integer, ByVal YOffset As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
Public Declare Sub glCullFace Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glDeleteLists Lib "OpenGL32.dll" (ByVal List As Integer, ByVal Range As Integer)
Public Declare Sub glDeleteTextures Lib "OpenGL32.dll" (ByVal N As Integer, ByRef Textures As Integer)
Public Declare Sub glDepthFunc Lib "OpenGL32.dll" (ByVal Func As Integer)
Public Declare Sub glDepthMask Lib "OpenGL32.dll" (ByVal Flag As Byte)
Public Declare Sub glDepthRange Lib "OpenGL32.dll" (ByVal Znear As Double, ByVal Zfar As Double)
Public Declare Sub glDisable Lib "OpenGL32.dll" (ByVal Cap As Integer)
Public Declare Sub glDisableClientState Lib "OpenGL32.dll" (ByVal iArray As Integer)
Public Declare Sub glDrawArrays Lib "OpenGL32.dll" (ByVal Mode As Integer, ByVal first As Integer, ByVal Count As Integer)
Public Declare Sub glDrawBuffer Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glDrawElements Lib "OpenGL32.dll" (ByVal Mode As Integer, ByVal Count As Integer, ByVal iType As Integer, ByRef Indices As Any)
Public Declare Sub glDrawPixels Lib "OpenGL32.dll" (ByVal Width As Integer, ByVal Height As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glEdgeFlag Lib "OpenGL32.dll" (ByVal Flag As Byte)
Public Declare Sub glEdgeFlagPointer Lib "OpenGL32.dll" (ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glEdgeFlagv Lib "OpenGL32.dll" (ByRef Flag As Byte)
Public Declare Sub glEnable Lib "OpenGL32.dll" (ByVal Cap As Integer)
Public Declare Sub glEnableClientState Lib "OpenGL32.dll" (ByVal iArray As Integer)
Public Declare Sub glEnd Lib "OpenGL32.dll" ()
Public Declare Sub glEndList Lib "OpenGL32.dll" ()
Public Declare Sub glEvalCoord1d Lib "OpenGL32.dll" (ByVal u As Double)
Public Declare Sub glEvalCoord1dv Lib "OpenGL32.dll" (ByRef u As Double)
Public Declare Sub glEvalCoord1f Lib "OpenGL32.dll" (ByVal u As Single)
Public Declare Sub glEvalCoord1fv Lib "OpenGL32.dll" (ByRef u As Single)
Public Declare Sub glEvalCoord2d Lib "OpenGL32.dll" (ByVal u As Double, ByVal V As Double)
Public Declare Sub glEvalCoord2dv Lib "OpenGL32.dll" (ByRef u As Double)
Public Declare Sub glEvalCoord2f Lib "OpenGL32.dll" (ByVal u As Single, ByVal V As Single)
Public Declare Sub glEvalCoord2fv Lib "OpenGL32.dll" (ByRef u As Single)
Public Declare Sub glEvalMesh1 Lib "OpenGL32.dll" (ByVal Mode As Integer, ByVal I1 As Integer, ByVal I2 As Integer)
Public Declare Sub glEvalMesh2 Lib "OpenGL32.dll" (ByVal Mode As Integer, ByVal I1 As Integer, ByVal I2 As Integer, ByVal J1 As Integer, ByVal J2 As Integer)
Public Declare Sub glEvalPoint1 Lib "OpenGL32.dll" (ByVal I As Integer)
Public Declare Sub glEvalPoint2 Lib "OpenGL32.dll" (ByVal I As Integer, ByVal J As Integer)
Public Declare Sub glFeedbackBuffer Lib "OpenGL32.dll" (ByVal Size As Integer, ByVal iType As Integer, ByRef Buffer As Single)
Public Declare Sub glFinish Lib "OpenGL32.dll" ()
Public Declare Sub glFlush Lib "OpenGL32.dll" ()
Public Declare Sub glFogf Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glFogfv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glFogi Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glFogiv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glFrontFace Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glFrustum Lib "OpenGL32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal Znear As Double, ByVal Zfar As Double)
Public Declare Function glGenLists Lib "OpenGL32.dll" (ByVal Range As Integer) As Integer
Public Declare Sub glGenTextures Lib "OpenGL32.dll" (ByVal N As Integer, ByRef Textures As Integer)
Public Declare Sub glGetBooleanv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Byte)
Public Declare Sub glGetClipPlane Lib "OpenGL32.dll" (ByVal Plane As Integer, ByRef Equation As Double)
Public Declare Sub glGetDoublev Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Double)
Public Declare Function glGetError Lib "OpenGL32.dll" () As Integer
Public Declare Sub glGetFloatv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetIntegerv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetLightfv Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetLightiv Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetMapdv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Query As Integer, ByRef V As Double)
Public Declare Sub glGetMapfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Query As Integer, ByRef V As Single)
Public Declare Sub glGetMapiv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Query As Integer, ByRef V As Integer)
Public Declare Sub glGetMaterialfv Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetMaterialiv Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetPixelMapfv Lib "OpenGL32.dll" (ByVal Map As Integer, ByRef Values As Single)
Public Declare Sub glGetPixelMapuiv Lib "OpenGL32.dll" (ByVal Map As Integer, ByRef Values As Integer)
Public Declare Sub glGetPixelMapusv Lib "OpenGL32.dll" (ByVal Map As Integer, ByRef Values As Integer)
Public Declare Sub glGetPointerv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Any)
Public Declare Sub glGetPolygonStipple Lib "OpenGL32.dll" (ByRef Mask As Byte)
Public Declare Function glGetString Lib "OpenGL32.dll" (ByVal Name As Integer) As String
Public Declare Sub glGetTexEnvfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetTexEnviv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetTexGendv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Double)
Public Declare Sub glGetTexGenfv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetTexGeniv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetTexImage Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glGetTexLevelParameterfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetTexLevelParameteriv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glGetTexParameterfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glGetTexParameteriv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glHint Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Mode As Integer)
Public Declare Sub glIndexMask Lib "OpenGL32.dll" (ByVal Mask As Integer)
Public Declare Sub glIndexPointer Lib "OpenGL32.dll" (ByVal iType As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glIndexd Lib "OpenGL32.dll" (ByVal C As Double)
Public Declare Sub glIndexdv Lib "OpenGL32.dll" (ByRef C As Double)
Public Declare Sub glIndexf Lib "OpenGL32.dll" (ByVal C As Single)
Public Declare Sub glIndexfv Lib "OpenGL32.dll" (ByRef C As Single)
Public Declare Sub glIndexi Lib "OpenGL32.dll" (ByVal C As Integer)
Public Declare Sub glIndexiv Lib "OpenGL32.dll" (ByRef C As Integer)
Public Declare Sub glIndexs Lib "OpenGL32.dll" (ByVal C As Integer)
Public Declare Sub glIndexsv Lib "OpenGL32.dll" (ByRef C As Integer)
Public Declare Sub glIndexub Lib "OpenGL32.dll" (ByVal C As Byte)
Public Declare Sub glIndexubv Lib "OpenGL32.dll" (ByRef C As Byte)
Public Declare Sub glInitNames Lib "OpenGL32.dll" ()
Public Declare Sub glInterleavedArrays Lib "OpenGL32.dll" (ByVal Format As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Function glIsEnabled Lib "OpenGL32.dll" (ByVal Cap As Integer) As Byte
Public Declare Function glIsList Lib "OpenGL32.dll" (ByVal List As Integer) As Byte
Public Declare Function glIsTexture Lib "OpenGL32.dll" (ByVal Texture As Integer) As Byte
Public Declare Sub glLightModelf Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glLightModelfv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glLightModeli Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glLightModeliv Lib "OpenGL32.dll" (ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glLightf Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glLightfv Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glLighti Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glLightiv Lib "OpenGL32.dll" (ByVal Light As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glLineStipple Lib "OpenGL32.dll" (ByVal Factor As Integer, ByVal Pattern As Integer)
Public Declare Sub glLineWidth Lib "OpenGL32.dll" (ByVal Width As Single)
Public Declare Sub glListBase Lib "OpenGL32.dll" (ByVal Base As Integer)
Public Declare Sub glLoadIdentity Lib "OpenGL32.dll" ()
Public Declare Sub glLoadMatrixd Lib "OpenGL32.dll" (ByRef M As Double)
Public Declare Sub glLoadMatrixf Lib "OpenGL32.dll" (ByRef M As Single)
Public Declare Sub glLoadName Lib "OpenGL32.dll" (ByVal Name As Integer)
Public Declare Sub glLogicOp Lib "OpenGL32.dll" (ByVal Opcode As Integer)
Public Declare Sub glMap1d Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal U1 As Double, ByVal U2 As Double, ByVal Stride As Integer, ByVal Order As Integer, ByRef Points As Double)
Public Declare Sub glMap1f Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal U1 As Single, ByVal U2 As Single, ByVal Stride As Integer, ByVal Order As Integer, ByRef Points As Single)
Public Declare Sub glMap2d Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal U1 As Double, ByVal U2 As Double, ByVal Ustride As Integer, ByVal Uorder As Integer, ByVal V1 As Double, ByVal V2 As Double, ByVal Vstride As Integer, ByVal Vorder As Integer, ByRef Points As Double)
Public Declare Sub glMap2f Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal U1 As Single, ByVal U2 As Single, ByVal Ustride As Integer, ByVal Uorder As Integer, ByVal V1 As Single, ByVal V2 As Single, ByVal Vstride As Integer, ByVal Vorder As Integer, ByRef Points As Single)
Public Declare Sub glMapGrid1d Lib "OpenGL32.dll" (ByVal un As Integer, ByVal U1 As Double, ByVal U2 As Double)
Public Declare Sub glMapGrid1f Lib "OpenGL32.dll" (ByVal un As Integer, ByVal U1 As Single, ByVal U2 As Single)
Public Declare Sub glMapGrid2d Lib "OpenGL32.dll" (ByVal un As Integer, ByVal U1 As Double, ByVal U2 As Double, ByVal Vn As Integer, ByVal V1 As Double, ByVal V2 As Double)
Public Declare Sub glMapGrid2f Lib "OpenGL32.dll" (ByVal un As Integer, ByVal U1 As Single, ByVal U2 As Single, ByVal Vn As Integer, ByVal V1 As Single, ByVal V2 As Single)
Public Declare Sub glMaterialf Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glMaterialfv Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glMateriali Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glMaterialiv Lib "OpenGL32.dll" (ByVal face As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glMatrixMode Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glMultMatrixd Lib "OpenGL32.dll" (ByRef M As Double)
Public Declare Sub glMultMatrixf Lib "OpenGL32.dll" (ByRef M As Single)
Public Declare Sub glNewList Lib "OpenGL32.dll" (ByVal List As Integer, ByVal Mode As Integer)
Public Declare Sub glNormal3b Lib "OpenGL32.dll" (ByVal nX As Byte, ByVal nY As Byte, ByVal Nz As Byte)
Public Declare Sub glNormal3bv Lib "OpenGL32.dll" (ByRef V As Byte)
Public Declare Sub glNormal3d Lib "OpenGL32.dll" (ByVal nX As Double, ByVal nY As Double, ByVal Nz As Double)
Public Declare Sub glNormal3dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glNormal3f Lib "OpenGL32.dll" (ByVal nX As Single, ByVal nY As Single, ByVal Nz As Single)
Public Declare Sub glNormal3fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glNormal3i Lib "OpenGL32.dll" (ByVal nX As Integer, ByVal nY As Integer, ByVal Nz As Integer)
Public Declare Sub glNormal3iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glNormal3s Lib "OpenGL32.dll" (ByVal nX As Integer, ByVal nY As Integer, ByVal Nz As Integer)
Public Declare Sub glNormal3sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glNormalPointer Lib "OpenGL32.dll" (ByVal iType As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glOrtho Lib "OpenGL32.dll" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal Znear As Double, ByVal Zfar As Double)
Public Declare Sub glPassThrough Lib "OpenGL32.dll" (ByVal Token As Single)
Public Declare Sub glPixelMapfv Lib "OpenGL32.dll" (ByVal Map As Integer, ByVal Mapsize As Integer, ByRef Values As Single)
Public Declare Sub glPixelMapuiv Lib "OpenGL32.dll" (ByVal Map As Integer, ByVal Mapsize As Integer, ByRef Values As Integer)
Public Declare Sub glPixelMapusv Lib "OpenGL32.dll" (ByVal Map As Integer, ByVal Mapsize As Integer, ByRef Values As Integer)
Public Declare Sub glPixelStoref Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glPixelStorei Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glPixelTransferf Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glPixelTransferi Lib "OpenGL32.dll" (ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glPixelZoom Lib "OpenGL32.dll" (ByVal Xfactor As Single, ByVal Yfactor As Single)
Public Declare Sub glPointSize Lib "OpenGL32.dll" (ByVal Size As Single)
Public Declare Sub glPolygonMode Lib "OpenGL32.dll" (ByVal face As Integer, ByVal Mode As Integer)
Public Declare Sub glPolygonOffset Lib "OpenGL32.dll" (ByVal Factor As Single, ByVal Units As Single)
Public Declare Sub glPolygonStipple Lib "OpenGL32.dll" (ByRef Mask As Byte)
Public Declare Sub glPopAttrib Lib "OpenGL32.dll" ()
Public Declare Sub glPopClientAttrib Lib "OpenGL32.dll" ()
Public Declare Sub glPopMatrix Lib "OpenGL32.dll" ()
Public Declare Sub glPopName Lib "OpenGL32.dll" ()
Public Declare Sub glPrioritizeTextures Lib "OpenGL32.dll" (ByVal N As Integer, ByRef Textures As Integer, ByRef Priorities As Single)
Public Declare Sub glPushAttrib Lib "OpenGL32.dll" (ByVal Mask As Integer)
Public Declare Sub glPushClientAttrib Lib "OpenGL32.dll" (ByVal Mask As Integer)
Public Declare Sub glPushMatrix Lib "OpenGL32.dll" ()
Public Declare Sub glPushName Lib "OpenGL32.dll" (ByVal Name As Integer)
Public Declare Sub glRasterPos2d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double)
Public Declare Sub glRasterPos2dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glRasterPos2f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single)
Public Declare Sub glRasterPos2fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glRasterPos2i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer)
Public Declare Sub glRasterPos2iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glRasterPos2s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer)
Public Declare Sub glRasterPos2sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glRasterPos3d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Public Declare Sub glRasterPos3dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glRasterPos3f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Public Declare Sub glRasterPos3fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glRasterPos3i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)
Public Declare Sub glRasterPos3iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glRasterPos3s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)
Public Declare Sub glRasterPos3sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glRasterPos4d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal W As Double)
Public Declare Sub glRasterPos4dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glRasterPos4f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal W As Single)
Public Declare Sub glRasterPos4fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glRasterPos4i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer, ByVal W As Integer)
Public Declare Sub glRasterPos4iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glRasterPos4s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer, ByVal W As Integer)
Public Declare Sub glRasterPos4sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glReadBuffer Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glReadPixels Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glRectd Lib "OpenGL32.dll" (ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
Public Declare Sub glRectdv Lib "OpenGL32.dll" (ByRef V1 As Double, ByRef V2 As Double)
Public Declare Sub glRectf Lib "OpenGL32.dll" (ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
Public Declare Sub glRectfv Lib "OpenGL32.dll" (ByRef V1 As Single, ByRef V2 As Single)
Public Declare Sub glRecti Lib "OpenGL32.dll" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Public Declare Sub glRectiv Lib "OpenGL32.dll" (ByRef V1 As Integer, ByRef V2 As Integer)
Public Declare Sub glRects Lib "OpenGL32.dll" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Public Declare Sub glRectsv Lib "OpenGL32.dll" (ByRef V1 As Integer, ByRef V2 As Integer)
Public Declare Function glRenderMode Lib "OpenGL32.dll" (ByVal Mode As Integer) As Integer
Public Declare Sub glRotated Lib "OpenGL32.dll" (ByVal Angle As Double, ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Public Declare Sub glRotatef Lib "OpenGL32.dll" (ByVal Angle As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Public Declare Sub glScaled Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Public Declare Sub glScalef Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Public Declare Sub glScissor Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
Public Declare Sub glSelectBuffer Lib "OpenGL32.dll" (ByVal Size As Integer, ByRef Buffer As Integer)
Public Declare Sub glShadeModel Lib "OpenGL32.dll" (ByVal Mode As Integer)
Public Declare Sub glStencilFunc Lib "OpenGL32.dll" (ByVal Func As Integer, ByVal Ref As Integer, ByVal Mask As Integer)
Public Declare Sub glStencilMask Lib "OpenGL32.dll" (ByVal Mask As Integer)
Public Declare Sub glStencilOp Lib "OpenGL32.dll" (ByVal Fail As Integer, ByVal ZFail As Integer, ByVal ZPass As Integer)
Public Declare Sub glTexCoord1d Lib "OpenGL32.dll" (ByVal S As Double)
Public Declare Sub glTexCoord1dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glTexCoord1f Lib "OpenGL32.dll" (ByVal S As Single)
Public Declare Sub glTexCoord1fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glTexCoord1i Lib "OpenGL32.dll" (ByVal S As Integer)
Public Declare Sub glTexCoord1iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord1s Lib "OpenGL32.dll" (ByVal S As Integer)
Public Declare Sub glTexCoord1sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord2d Lib "OpenGL32.dll" (ByVal S As Double, ByVal T As Double)
Public Declare Sub glTexCoord2dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glTexCoord2f Lib "OpenGL32.dll" (ByVal S As Single, ByVal T As Single)
Public Declare Sub glTexCoord2fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glTexCoord2i Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer)
Public Declare Sub glTexCoord2iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord2s Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer)
Public Declare Sub glTexCoord2sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord3d Lib "OpenGL32.dll" (ByVal S As Double, ByVal T As Double, ByVal R As Double)
Public Declare Sub glTexCoord3dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glTexCoord3f Lib "OpenGL32.dll" (ByVal S As Single, ByVal T As Single, ByVal R As Single)
Public Declare Sub glTexCoord3fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glTexCoord3i Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer, ByVal R As Integer)
Public Declare Sub glTexCoord3iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord3s Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer, ByVal R As Integer)
Public Declare Sub glTexCoord3sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord4d Lib "OpenGL32.dll" (ByVal S As Double, ByVal T As Double, ByVal R As Double, ByVal Q As Double)
Public Declare Sub glTexCoord4dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glTexCoord4f Lib "OpenGL32.dll" (ByVal S As Single, ByVal T As Single, ByVal R As Single, ByVal Q As Single)
Public Declare Sub glTexCoord4fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glTexCoord4i Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer, ByVal R As Integer, ByVal Q As Integer)
Public Declare Sub glTexCoord4iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoord4s Lib "OpenGL32.dll" (ByVal S As Integer, ByVal T As Integer, ByVal R As Integer, ByVal Q As Integer)
Public Declare Sub glTexCoord4sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glTexCoordPointer Lib "OpenGL32.dll" (ByVal Size As Integer, ByVal iType As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glTexEnvf Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glTexEnvfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glTexEnvi Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glTexEnviv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glTexGend Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByVal Param As Double)
Public Declare Sub glTexGendv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Double)
Public Declare Sub glTexGenf Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glTexGenfv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glTexGeni Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glTexGeniv Lib "OpenGL32.dll" (ByVal Coord As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glTexImage1D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal Internalformat As Integer, ByVal Width As Integer, ByVal Border As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glTexImage2D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal Internalformat As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Border As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glTexParameterf Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByVal Param As Single)
Public Declare Sub glTexParameterfv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Single)
Public Declare Sub glTexParameteri Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByVal Param As Integer)
Public Declare Sub glTexParameteriv Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal PName As Integer, ByRef Params As Integer)
Public Declare Sub glTexSubImage1D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal XOffset As Integer, ByVal Width As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glTexSubImage2D Lib "OpenGL32.dll" (ByVal Target As Integer, ByVal Level As Integer, ByVal XOffset As Integer, ByVal YOffset As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Format As Integer, ByVal iType As Integer, ByRef Pixels As Any)
Public Declare Sub glTranslated Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Public Declare Sub glTranslatef Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Public Declare Sub glVertex2d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double)
Public Declare Sub glVertex2dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glVertex2f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single)
Public Declare Sub glVertex2fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glVertex2i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer)
Public Declare Sub glVertex2iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertex2s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer)
Public Declare Sub glVertex2sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertex3d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
Public Declare Sub glVertex3dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glVertex3f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Public Declare Sub glVertex3fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glVertex3i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)
Public Declare Sub glVertex3iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertex3s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer)
Public Declare Sub glVertex3sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertex4d Lib "OpenGL32.dll" (ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal W As Double)
Public Declare Sub glVertex4dv Lib "OpenGL32.dll" (ByRef V As Double)
Public Declare Sub glVertex4f Lib "OpenGL32.dll" (ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal W As Single)
Public Declare Sub glVertex4fv Lib "OpenGL32.dll" (ByRef V As Single)
Public Declare Sub glVertex4i Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer, ByVal W As Integer)
Public Declare Sub glVertex4iv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertex4s Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Z As Integer, ByVal W As Integer)
Public Declare Sub glVertex4sv Lib "OpenGL32.dll" (ByRef V As Integer)
Public Declare Sub glVertexPointer Lib "OpenGL32.dll" (ByVal Size As Integer, ByVal iType As Integer, ByVal Stride As Integer, ByRef Pointer As Any)
Public Declare Sub glViewport Lib "OpenGL32.dll" (ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)

'____________________________________________
'Temp GLU
Public Declare Sub gluPerspective Lib "glu32.dll" (ByVal Fovy As Double, ByVal Aspect As Double, ByVal Near As Double, ByVal Far As Double)
Public Declare Function gluNewQuadric Lib "glu32.dll" () As Long
Public Declare Sub gluQuadricDrawStyle Lib "glu32.dll" (QuadObj As Long, ByVal DrawStyle As Long)
Public Declare Sub gluQuadricNormals Lib "glu32.dll" (QuadObj As Long, ByVal Normals As Long)
Public Declare Sub gluDeleteQuadric Lib "glu32.dll" (QuadObj As Long)
Public Declare Sub gluCylinder Lib "glu32.dll" (QuadObj As Long, ByVal BaseRadius As Double, ByVal TopRadius As Double, ByVal Height As Double, ByVal Slices As Integer, ByVal Stacks As Integer)
Public Declare Sub gluDisk Lib "glu32.dll" (QuadObj As Long, ByVal InnerRadius As Double, ByVal OuterRadius As Double, ByVal Slices As Integer, ByVal Loops As Integer)

Public Declare Function gluNewTess Lib "glu32.dll" () As Long
Public Declare Sub gluTessCallback Lib "glu32.dll" (ByVal tessObj As Long, ByVal Parameters As Long, ByVal FunctionPointer As Long)
Public Declare Sub gluTessBeginPolygon Lib "glu32.dll" (ByVal tessObj As Long, Data As Any)
Public Declare Sub gluTessEndPolygon Lib "glu32.dll" (ByVal tessObj As Long)

Public Declare Sub gluTessVertex Lib "glu32.dll" (ByVal tessObj As Long, ByRef Coords As Double, Data As Any)
'Public Declare Sub gluTessVertex Lib "Glu32.dll" (ByVal tessObj As Long, Coords() As Double, ByVal Data As Long)

Public Declare Sub gluDeleteTess Lib "glu32.dll" (ByVal tessObj As Long)
Public Declare Sub gluTessBeginContour Lib "glu32.dll" (ByVal tessObj As Long)
Public Declare Sub gluTessEndContour Lib "glu32.dll" (ByVal tessObj As Long)

''''''''''''''Pixel Format
Public Const PFD_DEPTH_DONTCARE = &H20000000
Public Const PFD_DOUBLEBUFFER = &H1
Public Const PFD_DOUBLEBUFFER_DONTCARE = &H40000000
Public Const PFD_DRAW_TO_BITMAP = &H8
Public Const PFD_DRAW_TO_WINDOW = &H4
Public Const PFD_GENERIC_ACCELERATED = &H1000
Public Const PFD_GENERIC_FORMAT = &H40
Public Const PFD_MAIN_PLANE = 0
Public Const PFD_NEED_PALETTE = &H80
Public Const PFD_NEED_SYSTEM_PALETTE = &H100
Public Const PFD_OVERLAY_PLANE = 1
Public Const PFD_STEREO = &H2
Public Const PFD_STEREO_DONTCARE = &H80000000
Public Const PFD_SUPPORT_GDI = &H10
Public Const PFD_SUPPORT_OPENGL = &H20
Public Const PFD_SWAP_COPY = &H400
Public Const PFD_SWAP_EXCHANGE = &H200
Public Const PFD_SWAP_LAYER_BUFFERS = &H800
Public Const PFD_TYPE_COLORINDEX = 1
Public Const PFD_TYPE_RGBA = 0
Public Const PFD_UNDERLAY_PLANE = (-1)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const GWL_STYLE = (-16)

Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Integer
Public Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal N As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal N As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Integer

Public Declare Function wglCreateContext Lib "OpenGL32.dll" (ByVal hdc As Long) As Long
Public Declare Function wglDeleteContext Lib "OpenGL32.dll" (ByVal hRenderingContext As Long) As Boolean
Public Declare Function wglGetCurrentContext Lib "OpenGL32.dll" () As Long
Public Declare Function wglGetCurrentDC Lib "OpenGL32.dll" () As Long
Public Declare Function wglMakeCurrent Lib "OpenGL32.dll" (ByVal hdc As Long, ByVal hRenderingContext As Long) As Boolean
Public Declare Function wglChoosePixelFormat Lib "OpenGL32.dll" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Integer
Public Declare Function wglSetPixelFormat Lib "OpenGL32.dll" (ByVal hdc As Long, ByVal N As Integer, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Boolean
Public Declare Sub wglSwapBuffers Lib "OpenGL32.dll" (ByVal hdc As Long)

Public Function InitializeOpenGL(dData As Data3D) As Boolean
Dim Result As Boolean
Dim PixelFormat As Integer
Dim PFD As PIXELFORMATDESCRIPTOR

With dData
    
    .glDC = GetDC(.glWnd)
    
    PFD.nSize = Len(PFD)
    PFD.nVersion = 1
    PFD.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_SUPPORT_OPENGL
    PFD.iPixelType = PFD_TYPE_RGBA
    PFD.cColorBits = 16
    PFD.cDepthBits = 32
    PFD.iLayerType = PFD_MAIN_PLANE
    
    PixelFormat = wglChoosePixelFormat(.glDC, PFD)
    If PixelFormat = 0 Then
        InitializeOpenGL = False
        Exit Function
    End If
    
    Result = SetPixelFormat(.glDC, PixelFormat, PFD)
    If Not Result Then
        InitializeOpenGL = False
        Exit Function
    End If
    
    .glRC = wglCreateContext(.glDC)
    If .glRC = 0 Then
        InitializeOpenGL = False
        Exit Function
    End If
    
    wglMakeCurrent .glDC, .glRC
    
    InitializeOpenGL = True
End With

End Function

Public Sub TerminateOpenGL(dData As Data3D)
DeleteList dData.Index
DeleteLinesList dData.Index
wglMakeCurrent 0, 0
wglDeleteContext dData.glRC
ReleaseDC dData.glWnd, dData.glDC
End Sub

Public Function AddGLData(ctlCanvas As Object) As Long
'===================================================
' This adds a new record to the list of initialized GL contexts
'===================================================
GLDataCount = GLDataCount + 1
ReDim Preserve GLData(1 To GLDataCount)
Set GLData(GLDataCount).Viewer = ctlCanvas
ctlCanvas.DataIndex = GLDataCount
GLData(GLDataCount).Index = GLDataCount
AddGLData = GLDataCount
End Function

Public Sub RemoveGLData(ByVal Index As Long)
'===================================================
' This removes a specified GLData from the list; scrolling down any
' successing records
'===================================================
If Index < 1 Or Index > GLDataCount Then Exit Sub
Dim Z As Long

'Set GLData(Index).Viewer = Nothing
If Index < GLDataCount Then
    For Z = Index To GLDataCount - 1
        GLData(Z) = GLData(Z + 1)
        GLData(Z).Index = Z
        GLData(Z).Viewer.DataIndex = Z
    Next
End If

GLDataCount = GLDataCount - 1
If GLDataCount > 0 Then ReDim Preserve GLData(1 To GLDataCount)
End Sub

Public Sub ClearGLData(ByVal Index As Long)
'===================================================
'
'===================================================
With GLData(Index)
    ReDim .Points3D(1 To 1)
    ReDim .Lines3D(1 To 1)
    ReDim .Polygons3D(1 To 1)
    .Points3DCount = 0
    .Lines3DCount = 0
    .Polygons3DCount = 0
    Do While .PolygonOrder.Count > 0
        .PolygonOrder.Remove 1
    Loop
    DeleteList Index
    DeleteLinesList Index
End With
End Sub

Public Function GenerateNewList() As Long
Dim Z As Long, Q As Long

For Z = 1 To GLDataCount
    If Q < GLData(Z).ListID Then Q = GLData(Z).ListID
    If Q < GLData(Z).ListLinesID Then Q = GLData(Z).ListLinesID
Next

GenerateNewList = Q + 1
End Function

Public Sub NewList(ByVal nDataIndex As Long)
DeleteList nDataIndex
GLData(nDataIndex).ListID = GenerateNewList
End Sub

Public Sub NewLinesList(ByVal nDataIndex As Long)
DeleteLinesList nDataIndex
GLData(nDataIndex).ListLinesID = GenerateNewList
End Sub

Public Sub DeleteList(ByVal nDataIndex As Long)
With GLData(nDataIndex)
    If glIsList(.ListID) <> GL_FALSE Then glDeleteLists .ListID, 1
    .ListID = 0
End With
End Sub

Public Sub DeleteLinesList(ByVal nDataIndex As Long)
With GLData(nDataIndex)
    If glIsList(.ListLinesID) <> GL_FALSE Then glDeleteLists .ListLinesID, 1
    .ListLinesID = 0
End With
End Sub

'===================================================
'
'===================================================
'
'Public Sub ClearStructure(ByVal Index As Long)
'With GLData(Index)
'    ReDim .Points3D(1 To 1)
'    ReDim .Lines3D(1 To 1)
'    ReDim .Polygons3D(1 To 1)
'    .Points3DCount = 0
'    .Lines3DCount = 0
'    .Polygons3DCount = 0
'    Do While .PolygonOrder.Count > 0
'        .PolygonOrder.Remove 1
'    Loop
'End With
'End Sub

'===================================================
'
'===================================================

Public Function StringToByteArray&(A() As Byte, S$)
Dim I&
If Len(S) Then
For I = 1 To Len(S)
    A(I - 1) = Asc(Mid$(S, I, 1))
Next
A(I) = 0
End If
End Function

'return the 255 values for the color components
Public Function GetBValue(X As Long) As Integer
GetBValue = (X \ &H10000) And &HFF
End Function

Public Function GetGValue(X As Long) As Integer
GetGValue = (X \ &H100) And &HFF
End Function

Public Function GetRValue(X As Long) As Integer
GetRValue = X And &HFF
End Function

Public Function Blue(X As Long) As Integer
Blue = (X \ &H10000) And &HFF
End Function

Public Function Green(X As Long) As Integer
Green = (X \ &H100) And &HFF
End Function

Public Function Red(X As Long) As Integer
Red = X And &HFF
End Function

'set current color from RGB255 values
Public Sub glRGB(R%, G%, B%)
    glColor3f R / 255, G / 255, B / 255
End Sub

'set current color from RGB255 values
Public Sub glColor32(Color&)
Dim R%, G%, B%
    R = GetRValue(Color)
    G = GetGValue(Color)
    B = GetBValue(Color)
    glColor3f R / 255, G / 255, B / 255
End Sub

Public Sub Create3dFont(hdc&, ID&, Typeface$, _
    Height&, weight&, Italic&, Color&)
Dim hFont&, R& ', base&
Dim agmf(256) As GLYPHMETRICSFLOAT 'agmf[256] ' Throw away
Dim face As String * 32
Dim LF As LOGFONT
    LF.lfHeight = Height '-10
    LF.lfWidth = 0
    LF.lfEscapement = 0
    LF.lfOrientation = 0
    LF.lfWeight = weight 'FW_BOLD
    LF.lfItalic = Italic 'FALSE
    LF.lfUnderline = False
    LF.lfStrikeOut = False
    LF.lfCharSet = 0
    LF.lfOutPrecision = 0
    LF.lfClipPrecision = 0
    LF.lfQuality = 0
    LF.lfPitchAndFamily = 0
    'lf.lfFaceName = Typeface
    'StringToByteArray LF.lfFaceName, Typeface
    LF.lfFaceName = Typeface
    '
    glColor32 Color
    hFont = CreateFontIndirect(LF)
    SelectObject hdc, hFont
    'create display lists for glyphs 0 through 255 with 0.1 extrusion
    ' and default deviation. The display list numbering starts at 1000
    ' (it could be any number).
    ID = glGenLists(96)
    If ID = 0 Then
        MsgBox "Unable to create font"
    Else
        R = wglUseFontOutlines(hdc, 32, 96, ID, 0#, 0.3, WGL_FONT_POLYGONS, agmf(0))
        If R = 0 Then
            MsgBox "Unable to create font"
            glDeleteLists ID, 96: ID = 0
        End If
    End If
    DeleteObject hFont
End Sub

Public Sub DrawGLText(Font&, S$, X!, Y!, Z!, Optional ByVal Color As Long)
Dim R As Long, I&
Dim A() As Byte
    'glPushAttrib amListBit 'don't know, what's this ?????
    glPushMatrix
    'set font size
    'glScalef m_FontSize*10.0f, m_FontSize*10.0f, m_FontSize*10.0f);
    'set font
    glListBase Font - 32 'ARIAL36 'TIMES36
    glTranslatef X, Y, Z
    'set forecolor
    If Not IsMissing(Color) Then
        glColor32 CLng(Color)
    End If
    ReDim A(0 To Len(S) + 1)
    StringToByteArray A, S
    glCallLists Len(S), GL_BYTE, A(0) 'ByVal s
    'GL_UNSIGNED_BYTE
    'GL_UNSIGNED_SHORT
'    For i = font To font + 96
'        glCallList i
'    Next
    glPopMatrix
    R = glGetError()
    If R <> GL_NO_ERROR Then MsgBox R
    glPopAttrib
End Sub

Public Function TesselatePolygon(X() As Double, Y() As Double, Z() As Double, Count As Long, P As Polygon3D)
Dim Tess As Long, V As Point3D, Q As Long, V3(0 To 2) As Double

ReDim TessPoly.Fans(1 To 1)
ReDim TessPoly.Strips(1 To 1)
ReDim TessPoly.Triangles(1 To 1)
TessPoly.FanCount = 0
TessPoly.StripCount = 0
TessPoly.TriangleCount = 0
TessPoly.NeedsTesselation = True
TessState = 0

Tess = gluNewTess
gluTessCallback Tess, GLU_TESS_BEGIN, AddressOf GLCallbackTessBegin
gluTessCallback Tess, GLU_TESS_VERTEX, AddressOf GLCallbackTessVertex
gluTessCallback Tess, GLU_TESS_END, AddressOf GLCallbackTessEnd

gluTessBeginPolygon Tess, 0
    gluTessBeginContour Tess
        For Q = 1 To Count
            V3(0) = X(Q)
            V3(1) = Y(Q)
            V3(2) = Z(Q)
            gluTessVertex Tess, V3(0), ByVal P.P(Q)
        Next Q
    gluTessEndContour Tess
gluTessEndPolygon Tess

gluDeleteTess Tess

P.NeedsTesselation = True
P.FanCount = TessPoly.FanCount
P.StripCount = TessPoly.StripCount
P.TriangleCount = TessPoly.TriangleCount
P.Fans = TessPoly.Fans
P.Strips = TessPoly.Strips
P.Triangles = TessPoly.Triangles
End Function

Public Sub GLCallbackTessBegin(ByVal P As Long)
TessState = P
With TessPoly
    If TessState <> 0 Then
        Select Case TessState
        Case GL_TRIANGLE_FAN
            .FanCount = .FanCount + 1
            ReDim Preserve .Fans(1 To .FanCount)
        Case GL_TRIANGLE_STRIP
            .StripCount = .StripCount + 1
            ReDim Preserve .Strips(1 To .StripCount)
        Case GL_TRIANGLES
            .TriangleCount = .TriangleCount + 1
            ReDim Preserve .Triangles(1 To .TriangleCount)
        End Select
    End If
End With

End Sub

Public Sub GLCallbackTessVertex(ByVal V As Long)  'Point3D
With TessPoly
    If TessState <> 0 Then
        Select Case TessState
        Case GL_TRIANGLE_FAN
            .Fans(.FanCount).Count = .Fans(.FanCount).Count + 1
            ReDim Preserve .Fans(.FanCount).P(1 To .Fans(.FanCount).Count)
            .Fans(.FanCount).P(.Fans(.FanCount).Count) = GLData(TessIndex).Points3D(V)
        Case GL_TRIANGLE_STRIP
            .Strips(.StripCount).Count = .Strips(.StripCount).Count + 1
            ReDim Preserve .Strips(.StripCount).P(1 To .Strips(.StripCount).Count)
            .Strips(.StripCount).P(.Strips(.StripCount).Count) = GLData(TessIndex).Points3D(V)
        Case GL_TRIANGLES
            .Triangles(.TriangleCount).Count = .Triangles(.TriangleCount).Count + 1
            ReDim Preserve .Triangles(.TriangleCount).P(1 To .Triangles(.TriangleCount).Count)
            .Triangles(.TriangleCount).P(.Triangles(.TriangleCount).Count) = GLData(TessIndex).Points3D(V)
        End Select
    End If
End With
End Sub

Public Sub GLCallbackTessEnd()
TessState = 0
End Sub

Public Function InterpolateColor(ByVal C1 As Long, ByVal C2 As Long, ByVal Percentage As Double) As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
R1 = Red(C1)
G1 = Green(C1)
B1 = Blue(C1)
R2 = Red(C2)
G2 = Green(C2)
B2 = Blue(C2)
InterpolateColor = RGB(R1 + (R2 - R1) * Percentage, G1 + (G2 - G1) * Percentage, B1 + (B2 - B1) * Percentage)
End Function

Public Function GetCurrentMatrix() As Single()
On Local Error Resume Next
Dim T() As Single, M As Integer
ReDim T(0 To 15)
glGetFloatv GL_MODELVIEW_MATRIX, T(0)
GetCurrentMatrix = T
End Function

Public Function MultVectorByCurrentMatrix(V As CVector) As CVector
Dim T() As Single
ReDim T(0 To 15)
Set MultVectorByCurrentMatrix = New CVector
glGetFloatv GL_MODELVIEW_MATRIX, T(0)
'glGetFloatv GL_PROJECTION_MATRIX, T(0)
With MultVectorByCurrentMatrix
    .X = T(0) * V.X + T(4) * V.Y + T(8) * V.Z '+ T(12)
    .Y = T(1) * V.X + T(5) * V.Y + T(9) * V.Z '+ T(13)
    .Z = T(2) * V.X + T(6) * V.Y + T(10) * V.Z '+ T(14)
End With
End Function
