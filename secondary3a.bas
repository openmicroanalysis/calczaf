Attribute VB_Name = "CodeSecondary3"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Sub SecondaryBraggDefocusLoadAllImages(sample() As TypeSample)
' Load all Bragg defocus images to module level for all channels performing a Bragg defocus correction

ierror = False
On Error GoTo SecondaryBraggDefocusLoadAllImagesError

' Dummy procedure for CalcZAF, Standard and Matrix

Exit Sub

' Errors
SecondaryBraggDefocusLoadAllImagesError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryBraggDefocusLoadAllImages"
ierror = True
Exit Sub

End Sub

Sub SecondaryBraggDefocusCalculateFraction(sampleline As Integer, chan As Integer, sample() As TypeSample)
' Calculate the Bragg defocus correction fraction (assume analysis point is in center of Bragg defocus image)

ierror = False
On Error GoTo SecondaryBraggDefocusCalculateFractionError

' Dummy procedure for CalcZAF, Standard and Matrix

Exit Sub

' Errors
SecondaryBraggDefocusCalculateFractionError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryBraggDefocusCalculateFraction"
ierror = True
Exit Sub

End Sub



