Attribute VB_Name = "CodeEMSA"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Dim stdsample(1 To 1) As TypeSample

Sub EMSAWriteSpectrum(mode As Integer, datarow As Integer, sample() As TypeSample, tfilename As String)
' This subroutine WRITES the EMSA/MAS Spectrum File Format. The data to be written
' to the file is stored in a sample structure. This structure is passed
' directly from the main program and can be used to send data to the file as needed.
'
' mode = 0 = EDS spectra
' mode = 1 = CL spectra
' datarow is the data row number containing the spectrum (1 to MAXROW%)
' sample() is a sample structure array (always dimensioned 1 to 1)
' tfilename is the string containing the file name to output
'
' Definitions of Arrays used in this Subroutine
' XADATA  = Real Array <= 4096 values containing X-Axis Data
' YADATA  = Real Array <= 4096 values containing Y-Axis Data
' EXPARSPT = Real Array <=20 values containing Expt Parameters of Spectrum
' EXPARMSC = Real Array <=20 values containing Expt Parameters of Microscope
' EXPARSAM = Real Array <=20 values containing Expt Parameters of Sample
' EXPAREDS = Real Array <=20 values containing Expt Parameters of EDS
' EXPARELS = Real Array <=20 values containing Expt Parameters of ELS
' EXPARAES = Real Array <=20 values containing Expt Parameters of AES
' EXPARWDS = Real Array <=20 values containing Expt Parameters of WDS
' EXPARPES = Real Array <=20 values containing Expt Parameters of PES
' EXPARXRF = Real Array <=20 values containing Expt Parameters of XRF
' EXPARCLS = Real Array <=20 values containing Expt Parameters of CLS
' EXPARGAM = Real Array <=20 values containing Expt Parameters of GAM
' SIGNALTY = Byte Array of 2 characters describing the signal type
' DATATYPE = Byte Array of 3 characters describing the format
' OPERMODE = Byte Array of 5 characters describing
'           the instrument operating mode
' EDSDET  = Byte Array of 6 characters describing the EDS detector type
' ELSDET  = Byte Array of 6 characters describing the ELS detector type
' XUNITS  = Byte Array <= 64 characters describing the X-Axis Data units
' YUNITS  = Byte Array <= 64 characters describing the Y-Axis Data units
' XLABEL  = Byte Array <= 64 characters describing the X-Axis Data label
' YLABEL  = Byte Array <= 64 characters describing the Y-Axis Data label
' FFORMAT = Byte Array <= 64 characters with the File Format Title
' TITLE   = Byte Array <= 64 characters with a spectrum Title
' DATE    = Byte Array of 12 characters with the date of acquisition
'           in the form DD-MMM-YYYY
' TIME    = Byte Array of 5 characters with the time of acquisition
'           in the form HH:MM
' OWNER   = Byte Array <= 64 characters with the owner/analysts name
' COMMENT = Byte Array <= 64 Character*1 used to hold a text comment string
'
'-------------------------------------------------------------------------
'
' Current Definitions of EXPAR (Experimental Parameter Array) Values
'
'---------------------------------------------------------------------------
' Parameters relating to the Spectrum Characteristics = EXPARSPT
'---------------------------------------------------------------------------
'
' Real EXPARSPT(20),VERSION,NPOINTS,NCOLUMNS,XPERCHAN,OFFSET,CHOFFSET
' Character*1 SIGNALTY(3),XUNITS(64),YUNITS(64),XLABEL(64),YLABEL(64),DATATYPE(2)
'
' EXPARSPT(1) = c#VERSION = File Format Version Number
' EXPARSPT(2) = c#NPOINTS = Total Number of Data Points in X &Y Data Arrays
' 1 <= NPOINTS <= 4096
' EXPARSPT(3) = c#NCOLUMNS = Number of columns of data
' 1 <= NCOLUMNS <= 5
' EXPARSPT(4) = c#XPERCHAN = Increment of X-axis units per channel (eV/Channel)
'           This is only useful if (x,y) paired data are not provided
' EXPARSPT(5) = c#OFFSET  = Energy value of first data point in eV
'           This is only useful if (x,y) paired data are not provided
' EXPARSPT(6) = c#CHOFFSET = Channel number which corresponds to zero units along
'            x-axis, this may be either a positive or negative value
' SIGNALTY(.) = c#SIGNALTYPE = Type of Spectroscopy
'       EDS = Energy Dispersive Spectroscopy
'       WDS = Wavelength Dispersive Spectroscopy
'       ELS = Energy Loss Spectroscopy
'       AES = Auger Electron Spectroscopy
'       PES = Photo Electron Spectroscopy
'       XRF = X-ray Fluorescence Spectroscopy
'       CLS = Cathodoluminescence Spectroscopy
'       GAM = Gamma Ray Spectroscopy
' XUNITS(.)  = c#XUNITS = Up to 64 characters describing the X-Axis Data units
' YUNITS(.)  = c#YUNITS = Up to 64 characters describing the Y-Axis Data units
' XLABEL(.)  = c#XLABEL = Up to 64 characters describing the X-Axis Data label
' YLABEL(.)  = c#YLABEL = Up to 64 characters describing the Y-Axis Data label
' DATATYPE(.) = c#DATATYPE = Type of data format
'          Y = Spectrum Y axis data only, X-axis data to be calculated
'           using XPERCHAN and OFFSET and the following formulae
' X = offset + CHANNEL * XPERCHAN
'          XY = Spectral data is in the form of XY pairs
'
'
'---------------------------------------------------------------------------
' Microscope/Microanalysis Instrument Parameters = EXPARMSC
'---------------------------------------------------------------------------
'
' Real EXPARMSC(20),BEAMKV,EMISSION,PROBECUR,BEAMDIA,MAGCAM,
' CONVANGLE, COLLANGLE
' Character*1 OPERMODE(5)
'
' EXPARMSC(1) = c#BEAMKV   = Accelerating Voltage of Instrument in kV
' EXPARMSC(2) = c#EMISSION  = Gun Emission current in microAmps
' EXPARMSC(3) = c#PROBECUR  = Probe current in nanoAmps
' EXPARMSC(4) = c#BEAMDIAM  = Diameter of incident probe in nanometers
' EXPARMSC(5) = c#MAGCAM  = Magnification or Camera Length (Mag in x, Cl in mm)
' EXPARMSC(6) = c#CONVANGLE = Convergence semi-angle of incident beam in milliRadians
' EXPARMSC(7) = c#COLLANGLE = Collection semi-angle of scattered beam in milliRad
' OPERMODE(.) = c#OPERMODE  = Operating Mode of Instrument
'                            IMAGE = Imaging Mode
'                            DIFFR = Diffraction Mode
'                            SCIMG = Scanning Imaging Mode
'                            SCDIF = Scanning Diffraction Mode
'
'---------------------------------------------------------------------------
' Experimental Parameters relating to the Sample = EXPARSAM
'---------------------------------------------------------------------------
'
' Real EXPARSAM(20),THICKNESS,XTILTSTAGE,YTILTSTAGE,XPOSITION,YPOSITION,
' ZPOSITION
'
' EXPARSAM(1) = c#THICKNESS = Specimen thickness in nanometers
' EXPARSAM(2) = c#XTILTSTAGE = Specimen stage tilt X-axis in degrees
' EXPARSAM(3) = c#YTILTSTAGE = Specimen stage tilt Y-axis in degrees
' EXPARSAM(4) = c#XPOSITION = X location of beam or specimen
' EXPARSAM(5) = c#YPOSITION = Y location of beam or specimen
' EXPARSAM(6) = c#ZPOSITION = Z location of beam or specimen
'
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to ELS = EXPARELS
'---------------------------------------------------------------------------
'
' Real EXPARELS(20),DWELLTIME,INTEGTIME
' Character*1 ELSDET(6)
'
' EXPARELS(1) = c#DWELLTIME = Dwell time per channel for serial data collection in msec
' EXPARELS(2) = c#INTEGTIME = Integration time per spectrum for parallel data collection
'            in milliseconds
' ELSDET(.) = c#ELSDET  = Type of ELS Detector
'            Serial = Serial ELS Detector
'            Parall = Parallel ELS Detector
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to EDS = EXPAREDS
'---------------------------------------------------------------------------
'
' Real EXPAREDS(20),ELEVANGLE,AZIMANGLE,SOLIDANGLE,LIVETIME,REALTIME
' Real TBEWIND,TAUWIND,TDEADLYR,TACTLYR,TALWIND,TPYWIND,TDIWIND,THCWIND
' Character*1 EDSDET(6)
'
' EXPAREDS(1) = c#ELEVANGLE = Elevation angle of EDS,WDS detector in degrees
' EXPAREDS(2) = c#AZIMANGLE = Azimuthal angle of EDS,WDS detector in degrees
' EXPAREDS(3) = c#SOLIDANGLE = Collection solid angle of detector in sR
' EXPAREDS(4) = c#LIVETIME = Signal Processor Active (Live) time in seconds
' EXPAREDS(5) = c#REALTIME = Total clock time used to record the spectrum in seconds
' EXPAREDS(6) = c#TBEWIND = Thickness of Be Window on detector in cm
' EXPAREDS(7) = c#TAUWIND = Thickness of Au Window/Electrical Contact in cm
' EXPAREDS(8) = c#TDEADLYR = Thickness of Dead Layer in cm
' EXPAREDS(9) = c#TACTLYR = Thickness of Active Layer in cm
' EXPAREDS(10) = c#TALWIND = Thickness of Aluminium Window in cm
' EXPAREDS(11) = c#TPYWIND = Thickness of Pyrolene Window in cm
' EXPAREDS(12) = c#TBNWIND = Thickness of Boron-Nitride Window in cm
' EXPAREDS(13) = c#TDIWIND = Thickness of Diamond Window in cm
' EXPAREDS(14) = c#THCWIND = Thickness of HydroCarbon Window in cm
' EDSDET(.)   = c#EDSDET = Type of X-ray Detector
'        SIBEW = Si(Li) with Be Window
'        SIUTW = Si(Li) with Ultra Thin Window
'        SIWLS = Si(Li) Windowless
'        GEBEW = Ge with Be Window
'        GEUTW = Ge with Ultra Thin Window
'        GEWLS = Ge Windowless
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to WDS = EXPARWDS
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to XRF = EXPARXRF
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to AES = EXPARAES
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to PES = EXPARPES
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to CLS = EXPARCLS
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to GAM = EXPARGAM
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'===========================================================================
'        END OF DEFINITIONS
'============================================================================

ierror = False
On Error GoTo EMSAWriteSpectrumError

Dim n As Integer, i As Integer
Dim astring As String
Dim temp1 As Single, temp2 As Single, temp3 As Single

' Check some EDS parameters
If mode% = 0 Then
If sample(1).EDSSpectraNumberofChannels%(datarow%) <= 0 Then GoTo EMSAWriteSpectrumNoSpectrumData
End If

' Check some CL parameters
If mode% = 1 Then
If sample(1).CLSpectraNumberofChannels%(datarow%) <= 0 Then GoTo EMSAWriteSpectrumNoSpectrumData
End If
     
' Write the required parameters as defined by EMSA/MAS
Open tfilename$ For Output As #EMSASpectrumFileNumber%

Print #EMSASpectrumFileNumber%, "#FORMAT      : EMSA/MAS Spectral Data File"
Print #EMSASpectrumFileNumber%, "#VERSION     : 1.0"
Print #EMSASpectrumFileNumber%, "#TITLE       : " & Left$(Trim$(sample(1).Name$ & "_" & Format$(sample(1).Linenumber&(datarow%))), 64)

' Write comment after converting crlf to commas
astring$ = sample(1).Description$
astring$ = Replace$(astring$, vbCrLf, VbComma$)
Print #EMSASpectrumFileNumber%, "#COMMENT     : " & Left$(Trim$(astring$), 64)

Print #EMSASpectrumFileNumber%, "#DATE        : " & Day(Now) & "-" & MonthSyms$(Month(Now)) & "-" & Year(Now)
Print #EMSASpectrumFileNumber%, "#TIME        : " & Hour(Now) & ":" & Minute(Now)

' Write the EPMA and EDS interface types
If InterfaceTypeStored% > 0 Then
Print #EMSASpectrumFileNumber%, "#OWNER       : " & Left$(Trim$(InterfaceString$(InterfaceTypeStored%) & ", " & MDBUserName$ & ", " & MDBFileTitle$), 64)
Else
Print #EMSASpectrumFileNumber%, "#OWNER       : " & Left$(Trim$(InterfaceString$(InterfaceType%) & ", " & MDBUserName$ & ", " & MDBFileTitle$), 64)
End If

' EDS spectra
If mode% = 0 Then
Print #EMSASpectrumFileNumber%, "#NCOLUMNS    : 1"
Print #EMSASpectrumFileNumber%, "#NPOINTS     : " & Format$(sample(1).EDSSpectraNumberofChannels(datarow%))
Print #EMSASpectrumFileNumber%, "#OFFSET      : " & Format$(sample(1).EDSSpectraStartEnergy!(datarow%) * EVPERKEV#)
Print #EMSASpectrumFileNumber%, "#XPERCHAN    : " & Format$(sample(1).EDSSpectraEVPerChannel!(datarow%))

Print #EMSASpectrumFileNumber%, "#XUNITS      : eV"
Print #EMSASpectrumFileNumber%, "#YUNITS      : cps"
Print #EMSASpectrumFileNumber%, "#DATATYPE    : Y"

Print #EMSASpectrumFileNumber%, "#SIGNALTYPE  : EDS"
Print #EMSASpectrumFileNumber%, "#XLABEL      : Energy (eV)"
Print #EMSASpectrumFileNumber%, "#YLABEL      : Cps"
End If

' CL spectra
If mode% = 1 Then
Print #EMSASpectrumFileNumber%, "#NCOLUMNS    : 1"
Print #EMSASpectrumFileNumber%, "#NPOINTS     : " & Format$(sample(1).CLSpectraNumberofChannels(datarow%))
Print #EMSASpectrumFileNumber%, "#OFFSET      : " & Format$(sample(1).CLSpectraStartEnergy!(datarow%))
Print #EMSASpectrumFileNumber%, "#XPERCHAN    : " & Format$((sample(1).CLSpectraEndEnergy!(datarow%) - sample(1).CLSpectraStartEnergy!(datarow%)) / (sample(1).CLSpectraNumberofChannels(datarow%) - 1))

Print #EMSASpectrumFileNumber%, "#XUNITS      : nm"
Print #EMSASpectrumFileNumber%, "#YUNITS      : cps"
Print #EMSASpectrumFileNumber%, "#DATATYPE    : XY"

Print #EMSASpectrumFileNumber%, "#SIGNALTYPE  : CLS"
Print #EMSASpectrumFileNumber%, "#XLABEL      : Wavelength (nm)"
Print #EMSASpectrumFileNumber%, "#YLABEL      : Cps"
End If

Print #EMSASpectrumFileNumber%, "#BEAMKV   -kV: " & Format$(sample(1).kilovolts!)
'Print #EMSASpectrumFileNumber%, "#EMISSION -uA: "

' Output measured beam current if available, otherwise just output requested beam current
If sample(1).OnBeamCounts!(datarow%) > 0# Then
Print #EMSASpectrumFileNumber%, "#PROBECUR -nA: " & Format$(sample(1).OnBeamCounts!(datarow%))
Else
Print #EMSASpectrumFileNumber%, "#PROBECUR -nA: " & Format$(sample(1).beamcurrent!)
End If

Print #EMSASpectrumFileNumber%, "#BEAMDIAM -nm: " & Format$(sample(1).beamsize! * NMPERMICRON&)
Print #EMSASpectrumFileNumber%, "#MAGCAM      : " & Format$(sample(1).magnificationanalytical!)
'Print #EMSASpectrumFileNumber%, "#CONVANGLE-mR: "
'Print #EMSASpectrumFileNumber%, "#COLLANGLE-mR: "
'Print #EMSASpectrumFileNumber%, "#OPERMODE    : "
'Print #EMSASpectrumFileNumber%, "#THICKNESS-nm: "
'Print #EMSASpectrumFileNumber%, "#XTILTSTAGE  : "
'Print #EMSASpectrumFileNumber%, "#YTILTSTAGE  : "
Print #EMSASpectrumFileNumber%, "#XPOSITION   : " & Format$(sample(1).StagePositions!(datarow%, 1))
Print #EMSASpectrumFileNumber%, "#YPOSITION   : " & Format$(sample(1).StagePositions!(datarow%, 2))
Print #EMSASpectrumFileNumber%, "#ZPOSITION   : " & Format$(sample(1).StagePositions!(datarow%, 3))

' EDS INFORMATION
If mode% = 0 Then
If EDSSpectraInterfaceTypeStored% > 0 Then
Print #EMSASpectrumFileNumber%, "#EDSDET      : " & InterfaceStringEDS$(EDSSpectraInterfaceTypeStored%)
Else
Print #EMSASpectrumFileNumber%, "#EDSDET      : " & InterfaceStringEDS$(EDSSpectraInterfaceType%)
End If
End If

' CL INFORMATION
If mode% = 1 Then
If CLSpectraInterfaceTypeStored% > 0 Then
Print #EMSASpectrumFileNumber%, "#EDSDET      : " & InterfaceStringCL$(CLSpectraInterfaceTypeStored%) & " (XY light, XY dark)"
Else
Print #EMSASpectrumFileNumber%, "#EDSDET      : " & InterfaceStringCL$(CLSpectraInterfaceType%) & " (XY light, XY dark)"
End If
End If

Print #EMSASpectrumFileNumber%, "#ELEVANGLE-dg: " & Format$(sample(1).takeoff!)
'Print #EMSASpectrumFileNumber%, "#AZIMANGLE-dg: "
'Print #EMSASpectrumFileNumber%, "#SOLIDANGL-sR: "

' EDS count time
If mode% = 0 Then
Print #EMSASpectrumFileNumber%, "#LIVETIME  -s: " & Format$(sample(1).EDSSpectraLiveTime!(datarow%))
Print #EMSASpectrumFileNumber%, "#REALTIME  -s: " & Format$(sample(1).EDSSpectraElapsedTime!(datarow%))
End If

' CL count time
If mode% = 1 Then
Print #EMSASpectrumFileNumber%, "#LIVETIME  -s: " & Format$(sample(1).CLAcquisitionCountTime!(datarow%))
End If

'Print #EMSASpectrumFileNumber%, "#TBEWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#TAUWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#TDEADLYR -cm: "
'Print #EMSASpectrumFileNumber%, "#TACTLYR  -cm: "
'Print #EMSASpectrumFileNumber%, "#TALWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#TPYWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#TBNWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#TDIWIND  -cm: "
'Print #EMSASpectrumFileNumber%, "#THCWIND  -cm: "

' Custom definitions (Nicholas Ritchie)
If sample(1).Type% = 1 Then                 ' add standard composition to custom field if sample is a standard

' Get standard from database
Call StandardGetMDBStandard(sample(1).number%, stdsample())
If ierror Then Exit Sub

' Load standard composition string
astring$ = stdsample(1).Name$
For i% = 1 To stdsample(1).LastChan%
astring$ = astring$ & VbComma$ & "(" & Trim$(MiscAutoUcase$(stdsample(1).Elsyms$(i%))) & ":" & Trim$(Format$(stdsample(1).ElmPercents!(i%))) & ")"
Next i%

Print #EMSASpectrumFileNumber%, "##D2STDCMP   : " & astring$
End If

' Write spectrum data start
Print #EMSASpectrumFileNumber%, "#SPECTRUM    : "

' Write EDS spectrum data (always write cps)
If mode% = 0 Then
For n% = 1 To sample(1).EDSSpectraNumberofChannels%(datarow%)
temp1! = sample(1).EDSSpectraIntensities&(datarow%, n%) / sample(1).EDSSpectraLiveTime!(datarow%)
Print #EMSASpectrumFileNumber%, MiscAutoFormat$(temp1!)
Next n%
End If

' Write CL spectrum data (always write in cps)
If mode% = 1 Then
For n% = 1 To sample(1).CLSpectraNumberofChannels%(datarow%)
temp1! = sample(1).CLSpectraIntensities&(datarow%, n%) / sample(1).CLAcquisitionCountTime!(datarow%)
temp2! = sample(1).CLSpectraDarkIntensities(datarow%, n%) / (sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%))
temp3! = sample(1).CLSpectraNanometers(datarow%, n%)        ' in nanometers
Print #EMSASpectrumFileNumber%, MiscAutoFormat$(temp3!) & VbComma$ & MiscAutoFormat$(temp1! - temp2!)
Next n%
End If

' All the data should now be in the arrays, write end of data marker
Print #EMSASpectrumFileNumber%, "#ENDOFDATA   : "

Close #EMSASpectrumFileNumber%
Exit Sub

' Errors
EMSAWriteSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "EMSAWriteSpectrum"
Close #EMSASpectrumFileNumber%
ierror = True
Exit Sub

EMSAWriteSpectrumNoSpectrumData:
If mode% = 0 Then msg$ = "Number of EDS spectrum points in the sample " & SampleGetString2$(sample()) & ", line " & Format$(sample(1).Linenumber&(datarow%)) & " is ZERO"
If mode% = 1 Then msg$ = "Number of CL spectrum points in the sample " & SampleGetString2$(sample()) & ", line " & Format$(sample(1).Linenumber&(datarow%)) & " is ZERO"
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAWriteSpectrum"
ierror = True
Exit Sub

End Sub

Sub EMSAReadSpectrum(mode As Integer, datarow As Integer, sample() As TypeSample, tfilename As String)
' Routine to read EMSA format for EDS and CL spectra
' mode = 0 = EDS spectra (single column of intensity data)
' mode = 1 = CL spectra (two columns of X/Y data: CL and dark)
' datarow is the data row number containing the spectrum (1 to MAXROW%)
' sample() is a sample structure array (always dimensioned 1 to 1)
' tfilename is the string containing the file name to output
'
' Definitions of Arrays used in this Subroutine
' XADATA  = Real Array <= 4096 values containing X-Axis Data
' YADATA  = Real Array <= 4096 values containing Y-Axis Data
' EXPARSPT = Real Array <=20 values containing Expt Parameters of Spectrum
' EXPARMSC = Real Array <=20 values containing Expt Parameters of Microscope
' EXPARSAM = Real Array <=20 values containing Expt Parameters of Sample
' EXPAREDS = Real Array <=20 values containing Expt Parameters of EDS
' EXPARELS = Real Array <=20 values containing Expt Parameters of ELS
' EXPARAES = Real Array <=20 values containing Expt Parameters of AES
' EXPARWDS = Real Array <=20 values containing Expt Parameters of WDS
' EXPARPES = Real Array <=20 values containing Expt Parameters of PES
' EXPARXRF = Real Array <=20 values containing Expt Parameters of XRF
' EXPARCLS = Real Array <=20 values containing Expt Parameters of CLS
' EXPARGAM = Real Array <=20 values containing Expt Parameters of GAM
' SIGNALTY = Byte Array of 2 characters describing the signal type
' DATATYPE = Byte Array of 3 characters describing the format
' OPERMODE = Byte Array of 5 characters describing
'           the instrument operating mode
' EDSDET  = Byte Array of 6 characters describing the EDS detector type
' ELSDET  = Byte Array of 6 characters describing the ELS detector type
' XUNITS  = Byte Array <= 64 characters describing the X-Axis Data units
' YUNITS  = Byte Array <= 64 characters describing the Y-Axis Data units
' XLABEL  = Byte Array <= 64 characters describing the X-Axis Data label
' YLABEL  = Byte Array <= 64 characters describing the Y-Axis Data label
' FFORMAT = Byte Array <= 64 characters with the File Format Title
' TITLE   = Byte Array <= 64 characters with a spectrum Title
' DATE    = Byte Array of 12 characters with the date of acquisition
'           in the form DD-MMM-YYYY
' TIME    = Byte Array of 5 characters with the time of acquisition
'           in the form HH:MM
' OWNER   = Byte Array <= 64 characters with the owner/analysts name
' COMMENT = Byte Array <= 64 Character*1 used to hold a text comment string
'
'-------------------------------------------------------------------------
'
' Current Definitions of EXPAR (Experimental Parameter Array) Values
'
'---------------------------------------------------------------------------
' Parameters relating to the Spectrum Characteristics = EXPARSPT
'---------------------------------------------------------------------------
'
' Real EXPARSPT(20),VERSION,NPOINTS,NCOLUMNS,XPERCHAN,OFFSET,CHOFFSET
' Character*1 SIGNALTY(3),XUNITS(64),YUNITS(64),XLABEL(64),YLABEL(64),DATATYPE(2)
'
' EXPARSPT(1) = c#VERSION = File Format Version Number
' EXPARSPT(2) = c#NPOINTS = Total Number of Data Points in X &Y Data Arrays
' 1 <= NPOINTS <= 4096
' EXPARSPT(3) = c#NCOLUMNS = Number of columns of data
' 1 <= NCOLUMNS <= 5
' EXPARSPT(4) = c#XPERCHAN = Increment of X-axis units per channel (eV/Channel)
'           This is only useful if (x,y) paired data are not provided
' EXPARSPT(5) = c#OFFSET  = Energy value of first data point in eV
'           This is only useful if (x,y) paired data are not provided
' EXPARSPT(6) = c#CHOFFSET = Channel number which corresponds to zero units along
'            x-axis, this may be either a positive or negative value
' SIGNALTY(.) = c#SIGNALTYPE = Type of Spectroscopy
'       EDS = Energy Dispersive Spectroscopy
'       WDS = Wavelength Dispersive Spectroscopy
'       ELS = Energy Loss Spectroscopy
'       AES = Auger Electron Spectroscopy
'       PES = Photo Electron Spectroscopy
'       XRF = X-ray Fluorescence Spectroscopy
'       CLS = Cathodoluminescence Spectroscopy
'       GAM = Gamma Ray Spectroscopy
' XUNITS(.)  = c#XUNITS = Up to 64 characters describing the X-Axis Data units
' YUNITS(.)  = c#YUNITS = Up to 64 characters describing the Y-Axis Data units
' XLABEL(.)  = c#XLABEL = Up to 64 characters describing the X-Axis Data label
' YLABEL(.)  = c#YLABEL = Up to 64 characters describing the Y-Axis Data label
' DATATYPE(.) = c#DATATYPE = Type of data format
'          Y = Spectrum Y axis data only, X-axis data to be calculated
'           using XPERCHAN and OFFSET and the following formulae
' X = offset + CHANNEL * XPERCHAN
'          XY = Spectral data is in the form of XY pairs
'
'
'---------------------------------------------------------------------------
' Microscope/Microanalysis Instrument Parameters = EXPARMSC
'---------------------------------------------------------------------------
'
' Real EXPARMSC(20),BEAMKV,EMISSION,PROBECUR,BEAMDIA,MAGCAM,
' CONVANGLE, COLLANGLE
' Character*1 OPERMODE(5)
'
' EXPARMSC(1) = c#BEAMKV   = Accelerating Voltage of Instrument in kV
' EXPARMSC(2) = c#EMISSION  = Gun Emission current in microAmps
' EXPARMSC(3) = c#PROBECUR  = Probe current in nanoAmps
' EXPARMSC(4) = c#BEAMDIAM  = Diameter of incident probe in nanometers
' EXPARMSC(5) = c#MAGCAM  = Magnification or Camera Length (Mag in x, Cl in mm)
' EXPARMSC(6) = c#CONVANGLE = Convergence semi-angle of incident beam in milliRadians
' EXPARMSC(7) = c#COLLANGLE = Collection semi-angle of scattered beam in milliRad
' OPERMODE(.) = c#OPERMODE  = Operating Mode of Instrument
'                            IMAGE = Imaging Mode
'                            DIFFR = Diffraction Mode
'                            SCIMG = Scanning Imaging Mode
'                            SCDIF = Scanning Diffraction Mode
'
'---------------------------------------------------------------------------
' Experimental Parameters relating to the Sample = EXPARSAM
'---------------------------------------------------------------------------
'
' Real EXPARSAM(20),THICKNESS,XTILTSTAGE,YTILTSTAGE,XPOSITION,YPOSITION,
' ZPOSITION
'
' EXPARSAM(1) = c#THICKNESS = Specimen thickness in nanometers
' EXPARSAM(2) = c#XTILTSTAGE = Specimen stage tilt X-axis in degrees
' EXPARSAM(3) = c#YTILTSTAGE = Specimen stage tilt Y-axis in degrees
' EXPARSAM(4) = c#XPOSITION = X location of beam or specimen
' EXPARSAM(5) = c#YPOSITION = Y location of beam or specimen
' EXPARSAM(6) = c#ZPOSITION = Z location of beam or specimen
'
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to ELS = EXPARELS
'---------------------------------------------------------------------------
'
' Real EXPARELS(20),DWELLTIME,INTEGTIME
' Character*1 ELSDET(6)
'
' EXPARELS(1) = c#DWELLTIME = Dwell time per channel for serial data collection in msec
' EXPARELS(2) = c#INTEGTIME = Integration time per spectrum for parallel data collection
'            in milliseconds
' ELSDET(.) = c#ELSDET  = Type of ELS Detector
'            Serial = Serial ELS Detector
'            Parall = Parallel ELS Detector
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to EDS = EXPAREDS
'---------------------------------------------------------------------------
'
' Real EXPAREDS(20),ELEVANGLE,AZIMANGLE,SOLIDANGLE,LIVETIME,REALTIME
' Real TBEWIND,TAUWIND,TDEADLYR,TACTLYR,TALWIND,TPYWIND,TDIWIND,THCWIND
' Character*1 EDSDET(6)
'
' EXPAREDS(1) = c#ELEVANGLE = Elevation angle of EDS,WDS detector in degrees
' EXPAREDS(2) = c#AZIMANGLE = Azimuthal angle of EDS,WDS detector in degrees
' EXPAREDS(3) = c#SOLIDANGLE = Collection solid angle of detector in sR
' EXPAREDS(4) = c#LIVETIME = Signal Processor Active (Live) time in seconds
' EXPAREDS(5) = c#REALTIME = Total clock time used to record the spectrum in seconds
' EXPAREDS(6) = c#TBEWIND = Thickness of Be Window on detector in cm
' EXPAREDS(7) = c#TAUWIND = Thickness of Au Window/Electrical Contact in cm
' EXPAREDS(8) = c#TDEADLYR = Thickness of Dead Layer in cm
' EXPAREDS(9) = c#TACTLYR = Thickness of Active Layer in cm
' EXPAREDS(10) = c#TALWIND = Thickness of Aluminium Window in cm
' EXPAREDS(11) = c#TPYWIND = Thickness of Pyrolene Window in cm
' EXPAREDS(12) = c#TBNWIND = Thickness of Boron-Nitride Window in cm
' EXPAREDS(13) = c#TDIWIND = Thickness of Diamond Window in cm
' EXPAREDS(14) = c#THCWIND = Thickness of HydroCarbon Window in cm
' EDSDET(.)   = c#EDSDET = Type of X-ray Detector
'        SIBEW = Si(Li) with Be Window
'        SIUTW = Si(Li) with Ultra Thin Window
'        SIWLS = Si(Li) Windowless
'        GEBEW = Ge with Be Window
'        GEUTW = Ge with Ultra Thin Window
'        GEWLS = Ge Windowless
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to WDS = EXPARWDS
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to XRF = EXPARXRF
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to AES = EXPARAES
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to PES = EXPARPES
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to CLS = EXPARCLS
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'---------------------------------------------------------------------------
' Experimental Parameters relating mainly to GAM = EXPARGAM
'---------------------------------------------------------------------------
'
'     Nothing currently defined
'
'===========================================================================
'        END OF DEFINITIONS
'============================================================================

ierror = False
On Error GoTo EMSAReadSpectrumError

Dim n As Integer, nc As Integer
Dim astring As String, bstring As String
Dim ystring As String
Dim temp As Single
Dim temp1 As Single, temp2 As Single, temp3 As Single
Dim tIntensityOption As Integer

Close (EMSASpectrumFileNumber%)
DoEvents

' Write the required parameters as defined by EMSA/MAS
Open tfilename$ For Input As #EMSASpectrumFileNumber%

Do Until EOF(EMSASpectrumFileNumber%)
Line Input #EMSASpectrumFileNumber%, astring$

' Load some throw away items
If InStr(astring$, "#FORMAT      : EMSA/MAS SPECTRAL DATA STANDARD") > 0 Then
End If
If InStr(astring$, "#VERSION     :") > 0 Then
End If
If InStr(astring$, "#TITLE       :") > 0 Then       ' standard already has a name
End If
If InStr(astring$, "#COMMENT     :") > 0 Then       ' standard already has a description
End If

If InStr(astring$, "#DATE        :") > 0 Then
End If
If InStr(astring$, "#TIME        :") > 0 Then
End If
If InStr(astring$, "#OWNER       :") > 0 Then
End If

' EDS spectra
If mode% = 0 Then
If InStr(astring$, "#NCOLUMNS    :") > 0 Then
nc% = Val(Mid$(astring$, Len("#NCOLUMNS    :") + 1))
If nc% <> 1 Then GoTo EMSAReadSpectrumNotSingleColumn           ' EDS is one spectrum (raw) one column of Y values only
End If

' Load number of points
If InStr(astring$, "#NPOINTS     :") > 0 Then
sample(1).EDSSpectraNumberofChannels%(datarow%) = Val(Mid$(astring$, Len("#NPOINTS     :") + 1))
End If

' Load x axis parameters
If InStr(astring$, "#OFFSET      :") > 0 Then
sample(1).EDSSpectraStartEnergy!(datarow%) = Val(Mid$(astring$, Len("#OFFSET      :") + 1)) / EVPERKEV#
End If

If InStr(astring$, "#XPERCHAN    :") > 0 Then
sample(1).EDSSpectraEVPerChannel!(datarow%) = Val(Mid$(astring$, Len("#XPERCHAN    :") + 1))
End If

If InStr(astring$, "#XUNITS      :") > 0 Then
End If
If InStr(astring$, "#YUNITS      :") > 0 Then
End If
If InStr(astring$, "#DATATYPE    :") > 0 Then
End If

' Check for proper spectral type
If InStr(astring$, "#SIGNALTYPE  :") > 0 Then
If Trim$(Mid$(astring$, Len("#SIGNALTYPE  :") + 1)) <> "EDS" Then GoTo EMSAReadSpectrumNotEDS
End If

If InStr(astring$, "#XLABEL      :") > 0 Then
End If
If InStr(astring$, "#YLABEL      :") > 0 Then
End If
End If

' CL spectra
If mode% = 1 Then
If InStr(astring$, "#NCOLUMNS    :") > 0 Then
nc% = Val(Mid$(astring$, Len("#NCOLUMNS    :") + 1))
If nc% <> 1 Then GoTo EMSAReadSpectrumNotSingleColumn           ' CL is one spectrum (dark corrected CL), 2 columns of X/Y data
End If

If InStr(astring$, "#NPOINTS     :") > 0 Then
sample(1).CLSpectraNumberofChannels%(datarow%) = Val(Mid$(astring$, Len("#NPOINTS     :") + 1))
End If

' Load x axis parameters
If InStr(astring$, "#OFFSET      :") > 0 Then
sample(1).CLSpectraStartEnergy!(datarow%) = Val(Mid$(astring$, Len("#OFFSET      :") + 1))
End If

If InStr(astring$, "#XPERCHAN    :") > 0 Then
temp! = Val(Mid$(astring$, Len("#XPERCHAN    :") + 1))
sample(1).CLSpectraEndEnergy!(datarow%) = sample(1).CLSpectraStartEnergy!(datarow%) + temp! * (sample(1).CLSpectraNumberofChannels(datarow%) - 1)
End If

If InStr(astring$, "#XUNITS      :") > 0 Then
End If

' Determine y units (assume raw counts)
If InStr(astring$, "#YUNITS      :") > 0 Then
ystring$ = Mid$(astring$, Len("#YUNITS      :") + 1)
If MiscStringsAreSame(ystring$, "cps") Then
tIntensityOption% = 1
Else
tIntensityOption% = 0
End If
End If

If InStr(astring$, "#DATATYPE    :") > 0 Then
End If

' Check for proper spectral type
If InStr(astring$, "#SIGNALTYPE  :") > 0 Then
If Trim$(Mid$(astring$, Len("#SIGNALTYPE  :") + 1)) <> "CL" And Trim$(Mid$(astring$, Len("#SIGNALTYPE  :") + 1)) <> "CLS" Then GoTo EMSAReadSpectrumNotCL   ' use "CL" for backward compatibility
End If

If InStr(astring$, "#XLABEL      :") > 0 Then
End If
If InStr(astring$, "#YLABEL      :") > 0 Then
End If
End If

' Load beam conditions
If InStr(astring$, "#BEAMKV   -kV:") > 0 Then
If mode% = 0 Then sample(1).EDSSpectraAcceleratingVoltage!(datarow%) = Val(Mid$(astring$, Len("#BEAMKV   -kV:") + 1))
If mode% = 1 Then sample(1).CLSpectraKilovolts!(datarow%) = Val(Mid$(astring$, Len("#BEAMKV   -kV:") + 1))
End If

If InStr(astring$, "#ELEVANGLE-dg:") > 0 Then
If mode% = 0 Then sample(1).EDSSpectraTakeOff!(datarow%) = Val(Mid$(astring$, Len("#ELEVANGLE-dg:") + 1))
End If

' EDS INFORMATION
If mode% = 0 Then
If InStr(astring$, "#LIVETIME  -s:") > 0 Then
sample(1).EDSSpectraLiveTime!(datarow%) = Val(Mid$(astring$, Len("#LIVETIME  -s:") + 1))
End If
If InStr(astring$, "#REALTIME  -s:") > 0 Then
sample(1).EDSSpectraElapsedTime!(datarow%) = Val(Mid$(astring$, Len("#REALTIME  -s:") + 1))
End If
End If

' CL INFORMATION
If mode% = 1 Then
If InStr(astring$, "#LIVETIME  -s:") > 0 Then
sample(1).CLAcquisitionCountTime!(datarow%) = Val(Mid$(astring$, Len("#LIVETIME  -s:") + 1))
End If
End If

' Check for spectra intensity start line
If InStr(astring$, "#SPECTRUM    : ") > 0 Then Exit Do
Loop

' Specify other parameter as needed
If mode% = 0 Then
sample(1).EDSSpectraEndEnergy!(datarow%) = sample(1).EDSSpectraStartEnergy!(datarow%) + sample(1).EDSSpectraEVPerChannel!(datarow%) / EVPERKEV# * (sample(1).EDSSpectraNumberofChannels(datarow%) - 1)
End If

' Check that a spectrum was indeed found in case just EOF() was found
If InStr(astring$, "#SPECTRUM    : ") = 0 Then GoTo EMSAReadSpectrumNotFound

' Loop on spectrum intensity lines
n% = 0
Do Until EOF(EMSASpectrumFileNumber%)
n% = n% + 1

' Input each intensity
Line Input #EMSASpectrumFileNumber%, astring$
If InStr(astring$, "#ENDOFDATA   : ") > 0 Then Exit Do

' Read EDS spectrum data (one column of Y data)
If mode% = 0 Then
temp! = Val(astring$)
If tIntensityOption% > 0 Then temp! = temp! * sample(1).EDSSpectraLiveTime!(datarow%)           ' de-normalize to actual counts
sample(1).EDSSpectraIntensities&(datarow%, n%) = temp!
End If

' Read CL spectrum data (one column of X/Y data, comma delimited)
If mode% = 1 Then
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
temp3! = Val(bstring$)                                          ' nanometers

Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
temp1! = Val(bstring$)                                          ' CL spectra intensities
If tIntensityOption% > 0 Then temp1! = temp1! * sample(1).CLAcquisitionCountTime!(datarow%)     ' de-normalize to actual counts

sample(1).CLSpectraIntensities&(datarow%, n%) = temp1!
sample(1).CLSpectraDarkIntensities&(datarow%, n%) = 0#      ' dark corrected
sample(1).CLSpectraNanometers!(datarow%, n%) = temp3!
End If

Loop

' Check for correct number of data points
If mode% = 0 Then
If sample(1).EDSSpectraNumberofChannels%(datarow%) <> n% - 1 Then GoTo EMSAReadSpectrumWrongNumberOfPoints
End If
If mode% = 1 Then
If sample(1).CLSpectraNumberofChannels%(datarow%) <> n% - 1 Then GoTo EMSAReadSpectrumWrongNumberOfPoints
End If

' Store number of intensity channels loaded
Close #EMSASpectrumFileNumber%
Exit Sub

' Errors
EMSAReadSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "EMSAReadSpectrum"
Close #EMSASpectrumFileNumber%
ierror = True
Exit Sub

EMSAReadSpectrumNotFound:
If mode% = 0 Then msg$ = "No EDS spectrum intensities were found in file " & tfilename$
If mode% = 1 Then msg$ = "No CL spectrum intensities were found in file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAReadSpectrum"
ierror = True
Exit Sub

EMSAReadSpectrumNotSingleColumn:
msg$ = "More than one data column found in " & tfilename$ & vbCrLf & vbCrLf
msg$ = msg$ & "Multiple data columns are not supported in EMSA EDS files at this time."
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAReadSpectrum"
ierror = True
Exit Sub

EMSAReadSpectrumNotEDS:
msg$ = tfilename$ & " is not an EMSA EDS spectrum file."
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAReadSpectrum"
ierror = True
Exit Sub

EMSAReadSpectrumNotCL:
msg$ = tfilename$ & " is not an EMSA CL spectrum file."
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAReadSpectrum"
ierror = True
Exit Sub

EMSAReadSpectrumWrongNumberOfPoints:
msg$ = "The number of spectral intensities does not match the number of intensites found in file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "EMSAReadSpectrum"
ierror = True
Exit Sub

End Sub

