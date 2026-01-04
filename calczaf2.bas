Attribute VB_Name = "CodeCalcZAF2"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Horizontal field width for non-graphical methods (in microns)
Dim ImageHFW As Single

' Array of filenames for output
Dim arraysize As Integer
Dim filenamearray() As String

Dim tlabel(1 To MAXCHAN% + 2) As String
Dim tData(1 To MAXCHAN% + 2) As Double

Dim InputLineCount As Integer
Dim OutputLineCount As Long

Dim CalcZAFTmpSample(1 To 1) As TypeSample

Sub CalcZAFElementalToOxideFactors(mode As Integer)
' Calcualte and display elemental and oxide factors
' mode = 1 elemental to oxide factors
' mode = 2 oxide to elemental factors

ierror = False
On Error GoTo CalcZAFElementalToOxideFactorsError

Dim num As Integer, i As Integer
Dim oxup As String, elup As String
Dim temp As Single
Dim temp1 As Single, temp2 As Single

' Loop on each element
Call IOWriteLog(vbNullString)
For i% = 1 To MAXELM%

' Calculate oxide weight
temp1! = AllAtomicWts!(i%) * AllCat%(i%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * AllOxd%(i%)

' Calculate elemental weight
temp2! = AllAtomicWts!(i%) * AllCat%(i%)

' Get oxide symbols
Call ElementGetSymbols(Symlo$(i%), AllCat%(i%), AllOxd%(i%), num%, oxup$, elup$)
If ierror Then Exit Sub

' Elemental to oxide
If mode% = 1 Then
temp! = temp1! / temp2!
msg$ = Format$(Symup$(i%) & " to " & oxup$, a16$) & MiscAutoFormat$(temp!)
Call IOWriteLog(msg$)

' Oxide to elemental
Else
temp! = temp2! / temp1!
msg$ = Format$(oxup$ & " to " & Symup$(i%), a16$) & MiscAutoFormat$(temp!)
Call IOWriteLog(msg$)
End If

Next i%

Exit Sub

' Errors
CalcZAFElementalToOxideFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementalToOxideFactors"
ierror = True
Exit Sub

End Sub

Sub CalcZAFElementDelete()
' Deletes an element

ierror = False
On Error GoTo CalcZAFElementDeleteError

' Blank fields
FormZAFELM.ComboElement.Text = vbNullString
FormZAFELM.ComboXRay.Text = vbNullString
FormZAFELM.ComboCations.Text = vbNullString
FormZAFELM.ComboOxygens.Text = vbNullString

FormZAFELM.TextWeight.Text = vbNullString
FormZAFELM.TextIntensity.Text = vbNullString
FormZAFELM.TextIntensityStd.Text = vbNullString

Exit Sub

' Errors
CalcZAFElementDeleteError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementDelete"
ierror = True
Exit Sub

End Sub

Sub CalcZAFRunStandard()
' This menu item runs the STANDARD.EXE program from CalcZAF for Window

ierror = False
On Error GoTo CalcZAFRunStandardError

Dim taskID As Long

' Check that file exists
msg$ = Dir$(ProgramPath$ & "STANDARD.EXE")
If msg$ = vbNullString Then
msg$ = "File " & ProgramPath$ & "STANDARD.EXE was not found. The file may have been deleted."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFRunStandard"
ierror = True
Exit Sub
End If

' Run the application
msg$ = ProgramPath$ & "STANDARD.EXE"
taskID& = Shell(msg$, 1)

Exit Sub

' Errors
CalcZAFRunStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFRunStandard"
ierror = True
Exit Sub

End Sub

Sub CalcZAFGetExcel2(analysis As TypeAnalysis, sample() As TypeSample)
' Routine to get calculated data and send to Excel

ierror = False
On Error GoTo CalcZAFGetExcel2Error

Dim chan As Integer

Dim nCol As Integer
Dim astring As String, bstring As String

' Check for new run
If NumberofSamples% < 1 Then
msg$ = "No samples loaded yet"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFGetExcel2"
ierror = True
Exit Sub
End If

If sample(1).LastChan% < 1 Then
msg$ = "No elements loaded yet"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFGetExcel2"
ierror = True
Exit Sub
End If

' Load sample name
astring$ = SampleGetString2$(sample())

' Load line number
nCol% = 1
tlabel$(nCol%) = "Line"
OutputLineCount& = OutputLineCount& + 1
tData#(nCol%) = CDbl(OutputLineCount&)

' Load elemental data
For chan% = 1 To sample(1).LastChan%
nCol% = nCol% + 1
tlabel$(nCol%) = vbNullString
tData#(nCol%) = CDbl(0#)

If ExcelMethodOption% = 0 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tData#(nCol%) = CDbl(analysis.WtPercents!(chan%))   ' load weight percents

ElseIf ExcelMethodOption% = 1 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%) & " " & sample(1).Xrsyms$(chan%)

If CalcZAFMode% = 0 Then    ' intensities from concentrations
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
tData#(nCol%) = CDbl(analysis.StdAssignsKfactors!(chan%))   ' load k-ratios
Else
tData#(nCol%) = CDbl(analysis.StdAssignsBetas!(chan%))   ' load beta factors
End If

Else                        ' concentrations from intensities
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
tData#(nCol%) = CDbl(analysis.UnkKrats!(chan%))   ' load k-ratios
Else
tData#(nCol%) = CDbl(analysis.UnkBetas!(chan%))   ' load beta factors
End If
End If

ElseIf ExcelMethodOption% = 2 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tData#(nCol%) = CDbl(analysis.AtPercents!(chan%))   ' load atomic percents

ElseIf ExcelMethodOption% = 3 Then
tlabel$(nCol%) = sample(1).Oxsyup$(chan%)
tData#(nCol%) = CDbl(analysis.OxPercents!(chan%))   ' load oxide percents

ElseIf ExcelMethodOption% = 4 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tData#(nCol%) = CDbl(analysis.Formulas!(chan%))     ' load formulas

ElseIf ExcelMethodOption% = 5 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tData#(nCol%) = CDbl(analysis.NormElPercents!(chan%))     ' load normalized elemental

ElseIf ExcelMethodOption% = 6 Then
tlabel$(nCol%) = sample(1).Oxsyup$(chan%)
tData#(nCol%) = CDbl(analysis.NormOxPercents!(chan%))     ' load normalized oxide
End If

Next chan%

' Load total
nCol% = nCol% + 1
tlabel$(nCol%) = vbNullString
tData#(nCol%) = CDbl(0#)

If ExcelMethodOption% = 0 Then
tlabel$(nCol%) = "Total"
tData#(nCol%) = CDbl(analysis.TotalPercent!)   ' load weight percents
bstring$ = astring$ & ", Elemental weight percents"

ElseIf ExcelMethodOption% = 1 Then
tlabel$(nCol%) = vbNullString
tData#(nCol%) = CDbl(0#)                        ' load k-ratios
bstring$ = astring$ & ", K-ratios"

ElseIf ExcelMethodOption% = 2 Then
tlabel$(nCol%) = "Total"
tData#(nCol%) = CDbl(100#)   ' load atomic percents
bstring$ = astring$ & ", Atomic Percents"

ElseIf ExcelMethodOption% = 3 Then
tlabel$(nCol%) = "Total"
tData#(nCol%) = CDbl(analysis.TotalPercent!)   ' load oxide percents
bstring$ = astring$ & ", Oxide Percents"

ElseIf ExcelMethodOption% = 4 Then
tlabel$(nCol%) = "Sum"
tData#(nCol%) = CDbl(analysis.TotalCations!)     ' load formulas
If sample(1).FormulaElement <> vbNullString Then
bstring$ = astring$ & ", Formula Atoms based on " & Str$(sample(1).FormulaRatio!) & " Atoms of " & MiscAutoUcase$(sample(1).FormulaElement$)
Else
bstring$ = astring$ & ", Formula Atoms based on the Sum of Cations"
End If

ElseIf ExcelMethodOption% = 5 Then
tlabel$(nCol%) = "Total"
tData#(nCol%) = CDbl(100#)     ' load normalized elemental
bstring$ = astring$ & ", Normalized Elemental Percents"

ElseIf ExcelMethodOption% = 6 Then
tlabel$(nCol%) = "Total"
tData#(nCol%) = CDbl(100#)     ' load normalized oxides
bstring$ = astring$ & ", Normalized Oxide Percents"
End If

' Send labels if indicated
Call ExcelSendLabelToSpreadsheet(Int(0), nCol%, bstring$, tlabel$())
If ierror Then Exit Sub

' Send data to Excel
Call ExcelSendDataToSpreadsheet(nCol%, tData#())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFGetExcel2Error:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFGetExcel2"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExcelOptionsLoad()
' Load Excel output option

ierror = False
On Error GoTo CalcZAFExcelOptionsLoadError

' Load Excel output option
FormEXCELOPTIONS.OptionExcelOutputOption(ExcelMethodOption%).Value = True

' Get the current sample
Call CalcZAFReturnSample(CalcZAFTmpSample())
If ierror Then Exit Sub

' Disable oxide output if not stoichiometric oxygen and not display as oxide
If CalcZAFTmpSample(1).OxideOrElemental% = 2 And Not CalcZAFTmpSample(1).DisplayAsOxideFlag Then
FormEXCELOPTIONS.OptionExcelOutputOption(3).Enabled = False
FormEXCELOPTIONS.OptionExcelOutputOption(6).Enabled = False
End If

Exit Sub

' Errors
CalcZAFExcelOptionsLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExcelOptionsLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExcelOptionsSave()
' Load Excel output option

ierror = False
On Error GoTo CalcZAFExcelOptionsSaveError

Dim i As Integer

' Save Excel output option
For i% = 0 To 6
If FormEXCELOPTIONS.OptionExcelOutputOption(i%).Value = True Then
ExcelMethodOption% = i%
End If
Next i%

Exit Sub

' Errors
CalcZAFExcelOptionsSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExcelOptionsSave"
ierror = True
Exit Sub

End Sub

Sub CalcZAFTypeAnalysis2(analysis As TypeAnalysis, sample() As TypeSample)
' Print last analysis parameters (CalcZAF only)

ierror = False
On Error GoTo CalcZAFTypeAnalysis2Error

Dim i As Integer
Dim temp As Single

msg$ = vbNullString
Call IOWriteLog(msg$)
msg$ = "Last Analysis Parameters..."
Call IOWriteLog(msg$)

' Type elements
msg$ = "ELEM: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "XRAY: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ATNUM:"
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(sample(1).AtomicNums%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ATWT: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(sample(1).AtomicWts!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ANGSTR"
For i% = 1 To sample(1).LastChan%
temp! = 0#
If sample(1).LineEnergy!(i%) > 0# Then temp! = ANGEV! / sample(1).LineEnergy!(i%)
msg$ = msg$ & MiscAutoFormat$(temp!)
Next i%
Call IOWriteLog(msg$)

msg$ = "Eo(kV)"
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & MiscAutoFormat$(sample(1).KilovoltsArray!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "ENERGY"
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & MiscAutoFormat$(sample(1).LineEnergy!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "EDGE: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & MiscAutoFormat$(sample(1).LineEdge!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "Eo/Ec:"
For i% = 1 To sample(1).LastChan%
temp! = 0#
If sample(1).LineEnergy!(i%) > 0# Then temp! = EVPERKEV# * sample(1).KilovoltsArray!(i%) / sample(1).LineEdge!(i%)
msg$ = msg$ & MiscAutoFormat$(temp!)
Next i%
Call IOWriteLog(msg$)

' Type out weight percents
msg$ = vbCrLf & "ELWT: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out standard assignments
msg$ = "STDS: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssigns%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbNullString
Call IOWriteLog(msg$)

' Standard counts
msg$ = "STCT: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsCounts!(i%))
Next i%
Call IOWriteLog(msg$)

' Type out unknown factors
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
msg$ = "UNKR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkKrats!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STKF: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKfactors!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
msg$ = "UNBE: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkKrats!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STBE: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsBetas!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
CalcZAFTypeAnalysis2Error:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFTypeAnalysis2"
ierror = True
Exit Sub

End Sub

Sub CalcZAFCalculateExportAll(tForm As Form)
' Opens a file, then calculates and exports all data

ierror = False
On Error GoTo CalcZAFCalculateExportAllError

Dim laststring As String
Dim response As Integer

Dim firstline As Boolean
Dim xydist As Single

' Set firstline flag
firstline = True
icancelauto = False

' Open the input file
Call CalcZAFImportOpen(tForm)
If ierror Then Exit Sub

' Open the export file (loads first data point)
ExportDataFile$ = MiscGetFileNameNoExtension(ImportDataFile$) & "_Export.dat"
Call CalcZAFExportOpen(tForm)
If ierror Then Exit Sub

' Check CalcZAFMode from import file
If CalcZAFMode% = 0 Then
Call IOStatusAuto("CalcZAFCalculateExportAll: skipping intensity calculation [mode=0]...")

' Calculate first point
Else
Call CalcZAFCalculate
If ierror Then Exit Sub

' Export first point
Call CalcZAFExportSend2(firstline, xydist!, laststring$)
If ierror Then Exit Sub
End If

' Loop on remaining points
Do Until EOF(ImportDataFileNumber%)

Call CalcZAFImportNext
If ierror Then Exit Sub

' Check CalcZAFMode
If CalcZAFMode% = 0 Then
Call IOStatusAuto("CalcZAFCalculateExportAll: skipping intensity calculation [mode=0]...")

Else
Call CalcZAFCalculate
If ierror Then Exit Sub

Call CalcZAFExportSend2(firstline, xydist!, laststring$)
If ierror Then Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
ierror = True
Exit Sub
End If
Loop

' Reset laststring
laststring$ = vbNullString

' Close the input file
Call CalcZAFImportClose
If ierror Then Exit Sub

' Close the export file
Call CalcZAFExportClose(Int(1))
If ierror Then Exit Sub

' Check if user wants to send k-ratio data file to Excel
msg$ = "Do you want to send the k-ratio data files to Excel?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "CalcZAFCalculateExportAll")

' Send k-ratio files to excel
If response% = vbYes Then
arraysize% = 1
ReDim Preserve filenamearray$(1 To arraysize%)
filenamearray$(arraysize%) = ExportDataFile$

Call ExcelSendFileListToExcel(arraysize%, filenamearray$(), tForm)
If ierror Then Exit Sub
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
CalcZAFCalculateExportAllError:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFCalculateExportAll"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportColumnString(laststring As String, sample() As TypeSample)
' Create a column label string and output it

ierror = False
On Error GoTo CalcZAFExportColumnStringError

Dim i As Integer
Dim astring As String

' Sample name
astring$ = VbDquote$ & Format$("SAMPLE", a80$) & VbDquote$

' Line number
astring$ = astring$ & vbTab & VbDquote$ & Format$("LINE", a80$) & VbDquote$

' Linear distance from boundary or relative distance from first point
If UseSecondaryBoundaryFluorescenceCorrectionFlag Then
astring$ = astring$ & vbTab & VbDquote$ & Format$("Boundary Dist(um)", a80$) & VbDquote$
Else
astring$ = astring$ & vbTab & VbDquote$ & Format$("Relative Dist(um)", a80$) & VbDquote$
End If

' Elemental labels
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " WT%", a80$) & VbDquote$
Next i%

' Add total label
astring$ = astring$ & vbTab & VbDquote$ & Format$("TOTAL", a80$) & VbDquote$

' Coordinates
astring$ = astring$ & vbTab & VbDquote$ & Format$("X-POS", a80$) & VbDquote$
astring$ = astring$ & vbTab & VbDquote$ & Format$("Y-POS", a80$) & VbDquote$
astring$ = astring$ & vbTab & VbDquote$ & Format$("Z-POS", a80$) & VbDquote$

' Atomic labels
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " AT%", a80$) & VbDquote$
Next i%

' K ratios
For i% = 1 To sample(1).LastElm%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " K-RAT", a80$) & VbDquote$
Next i%

' Standard Assignments
For i% = 1 To sample(1).LastElm%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " STD_NUM", a80$) & VbDquote$
Next i%

For i% = 1 To sample(1).LastElm%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " STD_NAM", a80$) & VbDquote$
Next i%

' Print column string if different from last string
If Trim$(astring$) <> Trim$(laststring$) Then Print #ExportDataFileNumber%, astring$
laststring$ = astring$

Exit Sub

' Errors
CalcZAFExportColumnStringError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportColumnString"
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportDataString(xydist As Single, analysis As TypeAnalysis, sample() As TypeSample)
' Create a data line string for each line

ierror = False
On Error GoTo CalcZAFExportDataStringError

Dim i As Integer, ip As Integer
Dim sampleline As Integer
Dim astring As String

' Sampleline is always 1
sampleline% = 1

' Print sample name
astring$ = VbDquote$ & SampleGetString2$(sample()) & VbDquote$

' Line number
astring$ = astring$ & vbTab & Format$(sample(1).Linenumber&(sampleline%), a80$)

' Linear distance to boundary or relative distance from first point
If UseSecondaryBoundaryFluorescenceCorrectionFlag Then
astring$ = astring$ & vbTab & Format$(sample(1).SecondaryFluorescenceBoundaryDistance!(sampleline%), a80$)
Else
astring$ = astring$ & vbTab & Format$(xydist!, a80$)
End If

' Elemental percents
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.WtPercents!(i%))
Next i%

' Add total
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.TotalPercent!)

' Stage Coordinates
astring$ = astring$ & vbTab & MiscAutoFormat$(sample(1).StagePositions!(sampleline%, 1))
astring$ = astring$ & vbTab & MiscAutoFormat$(sample(1).StagePositions!(sampleline%, 2))
astring$ = astring$ & vbTab & MiscAutoFormat$(sample(1).StagePositions!(sampleline%, 3))

' Convert to atomic percents
Call ConvertWeightToAtomic(sample(1).LastChan%, analysis.AtomicWts!(), analysis.WtPercents!(), analysis.AtPercents!())
If ierror Then Exit Sub

' Atomic percents
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.AtPercents!(i%))
Next i%

' K-ratios
For i% = 1 To sample(1).LastElm%
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.UnkKrats!(i%))
Next i%

For i% = 1 To sample(1).LastElm%
astring$ = astring$ & vbTab & Format$(sample(1).StdAssigns%(i%), i80$)
Next i%

' Standard assignments
For i% = 1 To sample(1).LastElm%
ip% = IPOS2%(NumberofStandards%, sample(1).StdAssigns%(i%), StandardNumbers%())
If ip% > 0 Then
astring$ = astring$ & vbTab & VbDquote$ & StandardNames$(ip%) & VbDquote$
Else
astring$ = astring$ & vbTab & VbDquote$ & vbNullString & VbDquote$
End If
Next i%

' Output data string
Print #ExportDataFileNumber%, astring$

Exit Sub

' Errors
CalcZAFExportDataStringError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportDataString"
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFConvertPouchouCSV()
' Convert the Pouchou k-ratios in kRatioConditions.csv to Pouchou2.dat for binary calculations
'
' Data file output format assumes one line for each binary. The first two
' columns are the atomic numbers of the two binary components
' to be calculated. The second two columns are the xray lines to use.
' ( 1 = Ka, 2 = Kb, 3 = La, 4 = Lb, 5 = Ma, 6 = Mb, 7 = by difference). The next
' two columns are the operating voltage and take-off angle. The next
' two columns are the wt. fractions of the binary components. The
' last two columns contains the k-exp values for calculation of k-calc/k-exp.
'
'       79     29     5    7    15.     52.5    .8015   .1983   .7400   .0
'       79     29     5    7    15.     52.5    .6036   .3964   .5110   .0
'       79     29     5    7    15.     52.5    .4010   .5992   .3120   .0
'       79     29     5    7    15.     52.5    .2012   .7985   .1450   .0

ierror = False
On Error GoTo CalcZAFConvertPouchouCSVError

Dim ip As Integer
Dim astring As String
Dim eO As Single, TOA As Single

' Input
Dim id As Integer
Dim symA As String, symB As String
Dim Xray As String, xray1 As String, xray2 As String

' Output
ReDim isym(1 To 2) As Integer
ReDim iray(1 To 2) As Integer
ReDim conc(1 To 2) As Single
ReDim kexp(1 To 2) As Single

icancelauto = False

' Get import filename
ImportDataFile2$ = ApplicationCommonAppData$ & "kRatioConditions.csv"
ExportDataFile$ = ApplicationCommonAppData$ & "Pouchou2.dat"

' Open normal files
Open ImportDataFile2$ For Input As #ImportDataFileNumber2%
Open ExportDataFile$ For Output As #ExportDataFileNumber2%

' Read column labels for input file
Line Input #ImportDataFileNumber2%, astring$
InputLineCount% = 0

' Check for end of file
Do While Not EOF(ImportDataFileNumber2%)
InputLineCount% = InputLineCount% + 1

msg$ = "Converting binary " & Str$(InputLineCount%) & "..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
ierror = True
Exit Sub
End If

' Read input file
Input #ImportDataFileNumber2%, id%, symA$, xray1$, xray2$, symB$, eO!, conc!(1), kexp!(1), TOA!
conc!(2) = 0#   ' not used
kexp!(2) = 0#   ' not used

If eO! < 1# Or eO! > 100# Then GoTo CalcZAFConvertPouchouCSVOutofLimits
If TOA! < 1# Or TOA! > 90# Then GoTo CalcZAFConvertPouchouCSVOutofLimits
If conc!(1) < 0# Or conc!(1) > 1# Then GoTo CalcZAFConvertPouchouCSVOutofLimits
If conc!(2) < 0# Or conc!(2) > 1# Then GoTo CalcZAFConvertPouchouCSVOutofLimits
If kexp!(1) < 0# Or kexp!(1) > 1# Then GoTo CalcZAFConvertPouchouCSVOutofLimits
If kexp!(2) < 0# Or kexp!(2) > 1# Then GoTo CalcZAFConvertPouchouCSVOutofLimits

' Convert element and x-ray symbols
ip% = IPOS1%(MAXELM%, symA$, Symlo$())
If ip% = 0 Then GoTo CalcZAFConvertPouchouCSVBadSymbol
isym%(1) = ip%
ip% = IPOS1%(MAXELM%, symB$, Symlo$())
If ip% = 0 Then GoTo CalcZAFConvertPouchouCSVBadSymbol
isym%(2) = ip%

Xray$ = xray1$ & xray2$
ip% = IPOS1%(MAXRAY% - 1, Xray$, Xraylo$())
If ip% = 0 Then GoTo CalcZAFConvertPouchouCSVBadXray
iray%(1) = ip%
iray%(2) = 7    ' unanalyzed element

' Check that both elements are not by difference
If iray%(1) = MAXRAY% And iray%(2) = MAXRAY% Then GoTo CalcZAFConvertPouchouCSVBothByDifference

' Check that at least one concentration is entered
If conc!(1) = 0# And conc!(2) = 0# Then GoTo CalcZAFConvertPouchouCSVNoConcData

' Check for valid kexp data if x-ray used
If iray%(1) <= MAXRAY% - 1 And kexp!(1) = 0# Then GoTo CalcZAFConvertPouchouCSVNoKexpData
If iray%(2) <= MAXRAY% - 1 And kexp!(2) = 0# Then GoTo CalcZAFConvertPouchouCSVNoKexpData

' Write binary elements, kilovolts and takeoff
Print #ExportDataFileNumber2%, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2)
Loop

' Close file
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%

Exit Sub

' Errors
CalcZAFConvertPouchouCSVError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVBadSymbol:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Bad element symbol on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVBadXray:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Bad x-ray symbol on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVOutofLimits:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Bad data on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVBothByDifference:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Both elements are by difference on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVNoConcData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No Conc data on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFConvertPouchouCSVNoKexpData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No K-exp data on line " & Str$(InputLineCount%) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFConvertPouchouCSV"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFSecondaryLoad()
' Load FormSECONDARY for secondary boundary fluorescence corrections

ierror = False
On Error GoTo CalcZAFSecondaryLoadError

' Get the current sample
Call CalcZAFReturnSample(CalcZAFTmpSample())
If ierror Then Exit Sub

' Load sample to FormSECONDARY
Call SecondarySampleLoadFrom(Int(1), CalcZAFTmpSample())
If ierror Then Exit Sub

' Load the element grid
Call CalcZAFSecondaryLoadList(CalcZAFTmpSample())
If ierror Then Exit Sub

' Load the boundary parameters in FormSECONDARY
If Penepma08CheckPenepmaVersion%() = 12 Then
Call SecondaryLoad(CalcZAFTmpSample())
If ierror Then Exit Sub
FormSECONDARY.Show vbModeless
Else
msg$ = "Penepma12 application files were not found. Please download the PENEPMA12.ZIP file and extract the files to the " & UserDataDirectory$ & " folder and check that the PENEPMA_Path, PENDBASE_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSecondaryLoad"
End If

Exit Sub

' Errors
CalcZAFSecondaryLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSecondaryLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSecondaryLoadList(sample() As TypeSample)
' Load the FormSECONDARY element grid (CalcZAF only)

ierror = False
On Error GoTo CalcZAFSecondaryLoadListError

Dim i As Integer, itemp As Integer
Dim tWidth As Single

' Blank the element grid
FormSECONDARY.GridElementList.Clear

' Set fixed columns
FormSECONDARY.GridElementList.FixedCols = 5

' Load the Grid Column labels
FormSECONDARY.GridElementList.row = 0
FormSECONDARY.GridElementList.col = 0
FormSECONDARY.GridElementList.Text = "Channel"
FormSECONDARY.GridElementList.col = 1
FormSECONDARY.GridElementList.Text = "Element"
FormSECONDARY.GridElementList.col = 2
FormSECONDARY.GridElementList.Text = "X-Ray"

' Load motor/crystal assignments
FormSECONDARY.GridElementList.col = 3
FormSECONDARY.GridElementList.Text = "Spectro"
FormSECONDARY.GridElementList.col = 4
FormSECONDARY.GridElementList.Text = "Crystal"

' Load enabled flag
FormSECONDARY.GridElementList.col = 5
FormSECONDARY.GridElementList.Text = "Enabled"

' Load Kratio DAT file or method
FormSECONDARY.GridElementList.col = 6
FormSECONDARY.GridElementList.Text = "PENFLUOR/FANAL Method/File"

FormSECONDARY.GridElementList.rows = MAXCHAN% + 1

' Initialize the Element List Grid Width
If FormSECONDARY.GridElementList.rows > 14 Then
itemp% = SCROLLBARWIDTH%
Else
itemp% = 0
End If

' Do not scale last col
tWidth! = 0#
For i% = 0 To FormSECONDARY.GridElementList.cols - 2
FormSECONDARY.GridElementList.ColWidth(i%) = (FormSECONDARY.GridElementList.Width - itemp%) / 15#
tWidth! = tWidth! + FormSECONDARY.GridElementList.ColWidth(i%)
Next i%

' Size last column
i% = FormSECONDARY.GridElementList.cols - 1
FormSECONDARY.GridElementList.ColWidth(i%) = (FormSECONDARY.GridElementList.Width - tWidth!)

' Load the element data into the grid
For i% = 1 To CalcZAFTmpSample(1).LastElm%
Call CalcZAFSecondaryUpdateList(i%)
If ierror Then Exit Sub
Next i%

Exit Sub

' Errors
CalcZAFSecondaryLoadListError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSecondaryLoadList"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSecondaryUpdateList(elementrow As Integer)
' This routine updates the element list grid based on the sample arrays (CalcZAF only)

ierror = False
On Error GoTo CalcZAFSecondaryUpdateListError

' Save sample from FormSECONDARY
Call SecondarySampleSaveTo(elementrow%, ImageHFW!, CalcZAFTmpSample())
If ierror Then Exit Sub

FormSECONDARY.GridElementList.row = elementrow%
FormSECONDARY.GridElementList.col = 0
FormSECONDARY.GridElementList.Text = Format$(elementrow%)

FormSECONDARY.GridElementList.col = 1
FormSECONDARY.GridElementList.Text = CalcZAFTmpSample(1).Elsyms$(elementrow%)
FormSECONDARY.GridElementList.col = 2
FormSECONDARY.GridElementList.Text = CalcZAFTmpSample(1).Xrsyms$(elementrow%)

' Motor/crystal assignments
FormSECONDARY.GridElementList.col = 3
FormSECONDARY.GridElementList.Text = Format$(CalcZAFTmpSample(1).MotorNumbers%(elementrow%))
FormSECONDARY.GridElementList.col = 4
FormSECONDARY.GridElementList.Text = CalcZAFTmpSample(1).CrystalNames$(elementrow%)

' Load secondary fluorescence flag
FormSECONDARY.GridElementList.col = 5
If CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryFlag(elementrow%) Then
FormSECONDARY.GridElementList.Text = "Yes"
Else
FormSECONDARY.GridElementList.Text = "No"
End If

' Load method ("Boundary.mdb" or K-ratio DAT filename)
FormSECONDARY.GridElementList.col = 6
If CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(elementrow%) = MiscGetFileNameOnly$(BoundaryMDBFile$) Then
FormSECONDARY.GridElementList.Text = MiscGetFileNameOnly$(BoundaryMDBFile$)
Else
FormSECONDARY.GridElementList.Text = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(elementrow%)
End If

Exit Sub

' Errors
CalcZAFSecondaryUpdateListError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSecondaryUpdateList"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSecondaryKratiosLoadForm(chan%)
' Load FormSecondaryKratios

ierror = False
On Error GoTo CalcZAFSecondaryKratiosLoadFormError

Dim tak As Single, keV As Single
Dim esym As String, xsym As String, tMatA As String, tMatB As String, tMatBStd As String

Dim t_npts As Long
Dim t_string1 As String, t_string2 As String, t_string3 As String
Dim t_eV() As Double, t_dist() As Double
Dim t_total() As Double, t_fluor() As Double
Dim t_flach() As Double, t_flabr() As Double
Dim t_flbch() As Double, t_flbbr() As Double
Dim t_pri_int() As Double, t_std_int() As Double

' Write K-ratio data (if specified) to disk file
'If Trim$(CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%)) <> vbNullString Then

' Load parameters for reading from probe database
'tak! = CalcZAFTmpSample(1).TakeoffArray!(chan%)
'keV! = CalcZAFTmpSample(1).KilovoltsArray!(chan%)
'esym$ = CalcZAFTmpSample(1).Elsyms$(chan%)
'xsym$ = CalcZAFTmpSample(1).Xrsyms$(chan%)
'tMatA$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%)
'tMatB$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%)
'tMatBStd$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%)
't_string1$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%)
't_string2$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%)
't_string3$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%)

' Read raw data from probe database
'Call DataBoundaryGetDATKratios(tak!, keV!, esym$, xsym$, tMatA$, tMatB$, tMatBStd$, t_string1$, t_string2$, t_string3$, t_eV#(), t_dist#(), t_total#(), t_fluor#(), t_flach#(), t_flabr#(), t_flbch#(), t_flbbr#(), t_pri_int#(), t_std_int#(), t_npts&)
'If ierror Then Exit Sub

' Write raw data to disk file if not present (original kratios.DAT file was deleted or probe MDB file was moved to another computer?)
'If Dir$(CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%)) = vbNullString Then

' Restore raw data values from probe database to module level
'Call SecondaryRestoreKratiosDAT(t_string1$, t_string2$, t_string3$, t_eV#(), t_dist#(), t_total#(), t_fluor#(), t_flach#(), t_flabr#(), t_flbch#(), t_flbbr#(), t_pri_int#(), t_std_int#(), t_npts&)
'If ierror Then Exit Sub

' Write the k-ratios.dat data file using module level variables
'Call SecondaryWriteKratiosDATFile(CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%))
'If ierror Then Exit Sub
'End If
'End If

' Load form with current sample parameters
Call SecondaryKratiosLoad(chan%, CalcZAFTmpSample())
If ierror Then Exit Sub '

' Load form
FormSECONDARYKratios.Show vbModal

' Save the saved form parameters to CalcZAFTmpSample
Call SecondarySampleSaveTo(chan%, ImageHFW!, CalcZAFTmpSample)
If ierror Then Exit Sub

' Save K-ratio data (if specified) to probe database
If Trim$(CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%)) <> vbNullString Then

' Read the k-ratios.dat data file and load into module level variables (in case user updated the kratios.dat file on disk)
Call SecondaryReadKratiosDATFile(CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%), CalcZAFTmpSample())
If ierror Then Exit Sub

' Load parameters for saving to probe database
tak! = CalcZAFTmpSample(1).TakeoffArray!(chan%)
keV! = CalcZAFTmpSample(1).KilovoltsArray!(chan%)
esym$ = CalcZAFTmpSample(1).Elsyms$(chan%)
xsym$ = CalcZAFTmpSample(1).Xrsyms$(chan%)
tMatA$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%)
tMatB$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%)
tMatBStd$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%)
t_string1$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%)
t_string2$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%)
t_string3$ = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%)

' Return raw data k-ratios from module level
Call SecondaryReturnKratiosDAT(t_string1$, t_string2$, t_string3$, t_eV#(), t_dist#(), t_total#(), t_fluor#(), t_flach#(), t_flabr#(), t_flbch#(), t_flbbr#(), t_pri_int#(), t_std_int#(), t_npts&)
If ierror Then Exit Sub

' Perform sanity check
If t_string1$ <> CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%) Then GoTo CalcZAFSecondaryKratiosLoadFormLinesDifferent1
If t_string2$ <> CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%) Then GoTo CalcZAFSecondaryKratiosLoadFormLinesDifferent2
If t_string3$ <> CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%) Then GoTo CalcZAFSecondaryKratiosLoadFormLinesDifferent3
End If

' Load element grid
Call CalcZAFSecondaryLoadList(CalcZAFTmpSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFSecondaryKratiosLoadFormError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSecondaryKratiosLoadForm"
ierror = True
Exit Sub

CalcZAFSecondaryKratiosLoadFormLinesDifferent1:
msg$ = "First line of Kratios.DAT file does not match saved string. This error should not occur, please contact probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSecondaryKratiosLoadForm"
ierror = True
Exit Sub

CalcZAFSecondaryKratiosLoadFormLinesDifferent2:
msg$ = "Second line of Kratios.DAT file does not match saved string. This error should not occur, please contact probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSecondaryKratiosLoadForm"
ierror = True
Exit Sub

CalcZAFSecondaryKratiosLoadFormLinesDifferent3:
msg$ = "Third line of Kratios.DAT file does not match saved string. This error should not occur, please contact probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSecondaryKratiosLoadForm"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSecondaryUpdateSample(sample() As TypeSample)
' Updates the paased sample with the secondary boundary parameters

ierror = False
On Error GoTo CalcZAFSecondaryUpdateSampleError

Dim chan As Integer

' Update the passed sample for the secondary boundary parameters
For chan% = 1 To sample(1).LastElm%
sample(1).SecondaryFluorescenceBoundaryFlag%(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryFlag%(chan%)
    
sample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%)
    
sample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%)
sample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%)
sample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%) = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%)
Next chan%
    
sample(1).SecondaryFluorescenceBoundaryDistanceMethod% = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryDistanceMethod%
sample(1).SecondaryFluorescenceBoundarySpecifiedDistance! = CalcZAFTmpSample(1).SecondaryFluorescenceBoundarySpecifiedDistance!
sample(1).SecondaryFluorescenceBoundaryCoordinateX1! = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryCoordinateX1!
sample(1).SecondaryFluorescenceBoundaryCoordinateY1! = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryCoordinateY1!
sample(1).SecondaryFluorescenceBoundaryCoordinateX2! = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryCoordinateX2!
sample(1).SecondaryFluorescenceBoundaryCoordinateY2! = CalcZAFTmpSample(1).SecondaryFluorescenceBoundaryCoordinateY2!
    
'    SecondaryFluorescenceBoundaryImageNumber As Integer     ' image number in BIM file (not used in matrix corrections)
'    SecondaryFluorescenceBoundaryImageFileName As String    ' original image file name (not used in matrix corrections)
       
'    SecondaryFluorescenceBoundaryDistance() As Single       ' (calculated in um) allocated in InitSample (1 To MAXROW%) (calculated)
'    SecondaryFluorescenceBoundaryKratios() As Single        ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%) (calculated)

Exit Sub

' Errors
CalcZAFSecondaryUpdateSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSecondaryUpdateSample"
ierror = True
Exit Sub

End Sub
