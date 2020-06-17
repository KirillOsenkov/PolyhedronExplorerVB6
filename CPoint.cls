VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This defines one point in three-dimensional space."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'A class to represent a point in the three-dimensional space
'Contains simply three coordinates of a points as public data members

Public X As Double 'represent the coordinates of a point
Public Y As Double
Public Z As Double

Public Sub Move(ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double)
Attribute Move.VB_Description = "Change point coordinates."
'move the point to a specified new location
X = dX
Y = dY
Z = dZ
End Sub
