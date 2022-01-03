Attribute VB_Name = "CodeGLOBAL3"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Import/Export files
Global Const ImportDataFileNumber% = 40          ' #40 ImportDataFile$ (STANDARD.DAT or *.DAT)
Global Const ImportDataFileNumber2% = 41         ' #41 ImportDataFile$ (*.DAT)
Global Const ExportDataFileNumber% = 42          ' #42 ExportDataFile$ (*.OUT)
Global Const ExportDataFileNumber2% = 43         ' #43 ExportDataFile$ (*.OUT)

' Constants for CalcZAF and Standard (and UserWin) only
Global Const HistogramDataFileNumber% = 44       ' #44 HistogramDataFile$ (*.TXT)

Global Const ModalInputDataFileNumber% = 49      ' #49 Modal analysis input ASCII file
Global Const ModalOuputDataFileNumber% = 50      ' #50 Modal analysis output ASCII file

Global CalculateAlternativeZbarsFlag As Integer
Global CalculateContinuumAbsorptionFlag As Integer

Global Const GSR_USELEG% = 3            ' Graphics Server legend flag

Global Const MAXCOEFF10% = 10           ' maximum number of fit coefficients
Global Const MAXCOEFF13% = 13           ' maximum number of fit coefficients

Global Const MILLIWATTSPERWATT& = 1000&  ' milliwatts per watt

Global ImportDataFile2 As String
Global HistogramDataFile As String

Global ImportDataFile As String
Global ExportDataFile As String

Global MaterialFcb As Double, MaterialWcb As Double

Global GetCmpFlag As Integer     ' 1 = new, 2 = modified, 3 = duplicate, flag for GETCMP (Standard.exe only)
Global DisplayZAFCalculationFlag As Integer ' see FormMAIN (Standard.exe only)

Global DisplayBiotiteCalculationFlag
Global DisplayAmphiboleCalculationFlag

Global UseTotalCationsCalculationFlag As Integer    ' (Standard.exe only)

Global TotalAcquisitionTimeString As String


