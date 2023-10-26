Attribute VB_Name = "CodeMATH2"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Function MathArcSin(x As Double) As Double
    
ierror = False
On Error GoTo MathArcsinError
    
    If x# = 1 Then
        MathArcSin# = PID# / 2
    Else
        MathArcSin# = Atn(x# / Sqr(-x# * x# + 1))
    End If

Exit Function

' Errors
MathArcsinError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArcsin"
ierror = True
Exit Function

End Function

Function MathArcCos(x As Double) As Double
    
ierror = False
On Error GoTo MathArcCosError
    
    If x# = 1 Then
        MathArcCos# = 0
    Else
        MathArcCos# = Atn(-x# / Sqr(-x# * x# + 1)) + 2 * Atn(1)
    End If

Exit Function

' Errors
MathArcCosError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArcCos"
ierror = True
Exit Function

End Function

Function MathTruncate(x As Double, Optional digit As Integer) As Double
' Return truncated number up to decimal digit

ierror = False
On Error GoTo MathTruncateError

Dim q As Double
    
    If IsMissing(digit%) Then digit% = 2
    q# = 10 ^ digit%
    
    MathTruncate# = (Int(x# * q#)) / q#
    
Exit Function

' Errors
MathTruncateError:
MsgBox Error$, vbOKOnly + vbCritical, "MathTruncate"
ierror = True
Exit Function

End Function

Function MathDegreesToRadians(ByVal degrees As Double) As Double
' Convert from degrees to radians

ierror = False
On Error GoTo MathDegreesToRadiansError

    MathDegreesToRadians# = degrees# / 57.29578

Exit Function

' Errors
MathDegreesToRadiansError:
MsgBox Error$, vbOKOnly + vbCritical, "MathDegreesToRadians"
ierror = True
Exit Function

End Function

Function MathRadiansToDegrees(ByVal radians As Double) As Double
' Convert from radians to degrees

ierror = False
On Error GoTo MathRadiansToDegreesError

    MathRadiansToDegrees# = radians# * 57.29578

Exit Function

' Errors
MathRadiansToDegreesError:
MsgBox Error$, vbOKOnly + vbCritical, "MathRadiansToDegrees"
ierror = True
Exit Function

End Function

Public Function MathArcCos2(x As Variant) As Variant
' Calculate the arc cosine in radians

ierror = False
On Error GoTo MathArcCos2Error
    
    Select Case x
        Case -1
            MathArcCos2 = 4 * Atn(1)
             
        Case 0:
            MathArcCos2 = 2 * Atn(1)
             
        Case 1:
            MathArcCos2 = 0
             
        Case Else:
            MathArcCos2 = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select

Exit Function

' Errors
MathArcCos2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathArcCos2"
ierror = True
Exit Function

End Function

