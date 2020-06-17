Attribute VB_Name = "modMath"
Option Explicit

Public Enum OrientationResultType
    oreOK
    oreNotClosed
    oreTooManyFaces
    oreNotConnected
    oreNonOrientable
End Enum

Public Function ToDouble(ByVal Expression As String) As Double
If Expression = "" Then ToDouble = 0: Exit Function
If IsNumeric(Expression) Then ToDouble = CDbl(Expression): Exit Function
If Val(Expression) <> 0 Then ToDouble = Val(Expression): Exit Function
End Function

Public Function IsNumber(ByVal Expression As String) As Boolean
If Expression = "" Then IsNumber = False: Exit Function
If IsNumeric(Expression) Or Val(Expression) <> 0 Then IsNumber = True
End Function
