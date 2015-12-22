Attribute VB_Name = "CodeCalcZAF2"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Array of filenames for output
Dim arraysize As Integer
Dim filenamearray() As String

Dim tlabel(1 To MAXCHAN% + 2) As String
Dim tdata(1 To MAXCHAN% + 2) As Double

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
temp1! = AllAtomicWts!(i%) * AllCat%(i%) + AllAtomicWts!(8) * AllOxd%(i%)

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
tdata#(nCol%) = CDbl(OutputLineCount&)

' Load elemental data
For chan% = 1 To sample(1).LastChan%
nCol% = nCol% + 1
tlabel$(nCol%) = vbNullString
tdata#(nCol%) = CDbl(0#)

If ExcelMethodOption% = 0 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.WtPercents!(chan%))   ' load weight percents

ElseIf ExcelMethodOption% = 1 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%) & " " & sample(1).Xrsyms$(chan%)

If CalcZAFMode% = 0 Then    ' intensities from concentrations
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
tdata#(nCol%) = CDbl(analysis.StdAssignsKfactors!(chan%))   ' load k-ratios
Else
tdata#(nCol%) = CDbl(analysis.StdAssignsBetas!(chan%))   ' load beta factors
End If

Else                        ' concentrations from intensities
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
tdata#(nCol%) = CDbl(analysis.UnkKrats!(chan%))   ' load k-ratios
Else
tdata#(nCol%) = CDbl(analysis.UnkBetas!(chan%))   ' load beta factors
End If
End If

ElseIf ExcelMethodOption% = 2 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.AtPercents!(chan%))   ' load atomic percents

ElseIf ExcelMethodOption% = 3 Then
tlabel$(nCol%) = sample(1).Oxsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.OxPercents!(chan%))   ' load oxide percents

ElseIf ExcelMethodOption% = 4 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.Formulas!(chan%))     ' load formulas

ElseIf ExcelMethodOption% = 5 Then
tlabel$(nCol%) = sample(1).Elsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.NormElPercents!(chan%))     ' load normalized elemental

ElseIf ExcelMethodOption% = 6 Then
tlabel$(nCol%) = sample(1).Oxsyup$(chan%)
tdata#(nCol%) = CDbl(analysis.NormOxPercents!(chan%))     ' load normalized oxide
End If

Next chan%

' Load total
nCol% = nCol% + 1
tlabel$(nCol%) = vbNullString
tdata#(nCol%) = CDbl(0#)

If ExcelMethodOption% = 0 Then
tlabel$(nCol%) = "Total"
tdata#(nCol%) = CDbl(analysis.TotalPercent!)   ' load weight percents
bstring$ = astring$ & ", Elemental weight percents"

ElseIf ExcelMethodOption% = 1 Then
tlabel$(nCol%) = vbNullString
tdata#(nCol%) = CDbl(0#)                        ' load k-ratios
bstring$ = astring$ & ", K-ratios"

ElseIf ExcelMethodOption% = 2 Then
tlabel$(nCol%) = "Total"
tdata#(nCol%) = CDbl(100#)   ' load atomic percents
bstring$ = astring$ & ", Atomic Percents"

ElseIf ExcelMethodOption% = 3 Then
tlabel$(nCol%) = "Total"
tdata#(nCol%) = CDbl(analysis.TotalPercent!)   ' load oxide percents
bstring$ = astring$ & ", Oxide Percents"

ElseIf ExcelMethodOption% = 4 Then
tlabel$(nCol%) = "Sum"
tdata#(nCol%) = CDbl(analysis.TotalCations!)     ' load formulas
If sample(1).FormulaElement <> vbNullString Then
bstring$ = astring$ & ", Formula Atoms based on " & Str$(sample(1).FormulaRatio!) & " Atoms of " & MiscAutoUcase$(sample(1).FormulaElement$)
Else
bstring$ = astring$ & ", Formula Atoms based on the Sum of Cations"
End If

ElseIf ExcelMethodOption% = 5 Then
tlabel$(nCol%) = "Total"
tdata#(nCol%) = CDbl(100#)     ' load normalized elemental
bstring$ = astring$ & ", Normalized Elemental Percents"

ElseIf ExcelMethodOption% = 6 Then
tlabel$(nCol%) = "Total"
tdata#(nCol%) = CDbl(100#)     ' load normalized oxides
bstring$ = astring$ & ", Normalized Oxide Percents"
End If

' Send labels if indicated
Call ExcelSendLabelToSpreadsheet(Int(0), nCol%, bstring$, tlabel$())
If ierror Then Exit Sub

' Send data to Excel
Call ExcelSendDataToSpreadsheet(nCol%, tdata#())
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

' Open the export file
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
Call ConvertWeightToAtomic(sample(1).LastChan%, analysis.AtomicWeights!(), analysis.WtPercents!(), analysis.AtPercents!())
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
