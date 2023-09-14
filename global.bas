Attribute VB_Name = "CodeGLOBAL"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Global ierror As Integer                           ' global error for backing out gracefully from error events

' New integrated intensity background fit parameter (on each side of scan)
Global Const MAX_INTEGRATED_BGD_FIT% = 50
Global Const MAX_ENERGY_ARRAY_SIZE% = 10
Global Const MAX_THROUGHPUT_ARRAY_SIZE% = 20

' Based on Cameca SX100/SXFive set times
Global Const KILOVOLT_SET_TIME! = 6#
Global Const BEAMCURRENT_SET_TIME_CAMECA! = 10#
Global Const BEAMCURRENT_SET_TIME_JEOL! = 12#
Global Const BEAMSIZE_SET_TIME! = 2#
Global Const HYSTERESIS_SET_TIME! = 12#
Global Const COLUMNCONDITION_SET_TIME! = 25#
Global Const BEAMMODE_SET_TIME! = 1#

Global Const PENEPMA_MATERIAL_TIME# = 4#         ' in seconds to init Penepma for demo EDS (running material.exe, per compositional element)
Global Const PENEPMA_STARTUP_TIME# = 15#         ' in seconds to init Penepma for demo EDS (starting penepma.exe)
Global Const PENEPMA_WDS_SYNTHESIS_TIME# = 6#    ' in seconds for demo WDS spectrum synthesis (per analyzed element wo polygonization modeling)
Global Const BOUNDARYNUMBEROFPOINTS& = 100       ' number of boaundary points for cluster digitize

' Special folders for system
Global Const SpecialFolder_CommonAppData = &H23  ' for all Windows users on this computer [Windows 2000 or later]
Global Const SpecialFolder_AppData = &H1A        ' for the current Windows user (roaming), on any computer on the network [Windows 98 or later]
Global Const SpecialFolder_LocalAppData = &H1C   ' for the current Windows user (non roaming), on this computer only [Windows 2000 or later]
Global Const SpecialFolder_Documents = &H5       ' the Documents folder for the current Windows user
Global Const SpecialFolder_Program_Files_CommonX86 = &H2C   ' for the Program Files (x86)\Common Files folder

' VB trappable errors
Global Const VB_OutOfMemory& = 7
Global Const VB_FileNotFound& = 53
Global Const VB_FileAlreadyOpen& = 55
Global Const VB_UnrecognizedDatabaseFormat& = 3343

'Global Const SW_HIDE& = 0
Global Const SW_SHOWNORMAL& = 1
'Global Const SW_SHOWMINIMIZED& = 2
'Global Const SW_SHOWMAXIMIZED& = 3
'Global Const SW_SHOWNOACTIVATE& = 4
'Global Const SW_SHOW& = 5
'Global Const SW_MINIMIZE& = 6
'Global Const SW_SHOWMINNOACTIVE& = 7
'Global Const SW_SHOWNA& = 8
'Global Const SW_RESTORE& = 9
'Global Const SW_SHOWDEFAULT& = 10

' Declare Windows API Functions
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, _
    ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' New constants for non-traditional emission lines
Global Const MAXRAY_OLD% = 6            ' maximum xray symbols (ka,kb,la,lb,ma,mb)
Global Const MAXRAY% = 13               ' maximum xray symbols (ka,kb,la,lb,ma,mb,ln,lg,lv,ll,mg,mz," ") including blank for non-analyzed

' Constants for array declarations
Global Const MAXINTERFACE% = 5          ' maximum number of instrument interfaces (0 to MAXINTERFACE%)
Global Const MAXINTERFACE_EDS% = 6      ' maximum number of EDS spectra interfaces (0 to MAXINTERFACE_EDS)
Global Const MAXINTERFACE_IMAGE% = 10   ' maximum number of imaging interfaces (0 to MAXINTERFACE_IMAGE)
Global Const MAXINTERFACE_CL% = 4       ' maximum number of CL spectra interfaces (0 to MAXINTERFACE_CL)

Global Const MAXCHAN% = 72              ' maximum elements per sample
Global Const MAXCHAN1% = MAXCHAN% + 1   ' maximum elements plus 1 (stoichiometric oxygen)
Global Const MAXSTD% = 132              ' maximum standards per run (changed from 128 to 132 08-23-2017)
Global Const MAXROW% = 500              ' maximum lines per sample (changed to 500 rows 08-11-2012)
Global Const MAXVOLATILE% = 400         ' maximum volatile/alternating intensities per chan per line per sample
Global Const MAXEDG% = 9                ' maximum emission edges
Global Const MAXELM% = 100              ' maximum elements (do not change due to data restrictions in AbsorbGetMAC)
Global Const MAXEMP% = 20               ' maximum empirical MAC/APFs
Global Const MAXSAMPLE% = 19999         ' maximum samples per run
Global Const MAXINTF% = 6               ' maximum interferences per element
Global Const MAXINDEX% = 10000          ' maximum standards per standard database   (changed to 10000 as of 2-27-2007)
Global Const MAXMAN% = 112              ' maximum MAN assignments per element   (changed from 36 to 112 08-23-2017)
Global Const MAXSET% = 30               ' maximum sets for drift correction
Global Const MAXCRYSTYPE% = 60          ' maximum crystal types
Global Const MAXCRYS% = 6               ' maximum crystals per spectrometer

Global Const MAXCOND% = 64              ' maximum number of analytical or column conditions per sample

Global Const MAXSPEC% = 6                       ' maximum spectrometers per run (spectro 0 = EDS)
Global Const MAXAXES% = 3                       ' maximum stage axes (x, y, z)
Global Const MAXMOT% = MAXSPEC% + MAXAXES%      ' maximum motors (spectrometer + stage)

Global Const MAXSPECTRA_CL% = 4096      ' maximum number of CL channels

Global Const MAXDIM% = 3                ' maximum number of dimensions for matrix transformation
Global Const MAXCOEFF% = 3              ' maximum number of linear fit coefficients
Global Const MAXCOEFF4% = 4             ' maximum number of linear fit coefficients
Global Const MAXCOEFF9% = 9             ' maximum number of fit coefficients
Global Const MAXBITMAP% = 12            ' maximum number of stage bit maps
Global Const MAXLINE& = 99999           ' maximum number of data lines per run
Global Const MAXDET% = 12               ' maximum number of detector parameters
Global Const MAXMULTI_OLD% = 12         ' (old) maximum number of points on each side for multi point background acquisition
Global Const MAXMULTI% = 18             ' maximum number of points on each side for multi point background acquisition

Global Const MAXCATION% = 100           ' maximum number of formula cations (1 to MAXCATION% - 1) and oxygens (0 to MAXCATION% - 1)
Global Const MAXCI% = 5                 ' maximum number of t-test confidence intervals
Global Const MAXPLOTPOINTS% = 10000     ' maximum number of points in plot graph control
Global Const MAXORDER% = 72             ' maximum number of elements per spectrometer (sample().OrderNumbers)
Global Const MAXEND% = 4                ' maximum number of mineral end members per mineral type
Global Const MAXPHASCAN% = 256          ' maximum number of PHA scan data points (see Cameca MCA PHA code)
Global Const MAXBRAGG% = 9              ' maximum analytical Bragg order
Global Const MAXELEMENTSORTMETHODS% = 3        ' 0 to MAXELEMENTSORTMETHODS% (0 = none)

Global Const MAXCORRECTION% = 6          ' maximum number of matrix correction types (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
Global Const MAXBINARY% = 11             ' maximum binary pairs for alpha factor calculations
Global Const MAXDENSITY# = 30#           ' maximum density (gm/cm3)
Global Const BLANKINGVALUE! = 1.70141E+38   ' Surfer blanking grid value
Global Const MAXPATHLENGTH% = 255        ' maximum file path length
Global Const MINCPSPERNA! = 0.00000000001 ' default minimum bgd count rate for P/B calculations
Global Const SIMULATION_ZERO! = 0.0000000001    ' zero intensity

Global Const MAXEMPFAC% = 100            ' maximum empirical alpha factors (from EMPFAC.DAT)
Global Const MAXCALIBRATE% = 5           ' maximum number of elements for multiple peak calibration
Global Const LIF2D! = 4.0267             ' 2d spacing of LIF crystal

Global Const SCROLLBARWIDTH% = 325       ' scroll bar width for sizing grid column widths (twips)
Global Const WINDOWBORDERWIDTH% = 175    ' window border width for sizing image width
Global Const FRAMEBORDERWIDTH% = 120     ' window border width for sizing controls
Global Const FRAMEBORDERWIDTH2% = 45     ' window border width for sizing controls
Global Const GRIDCOLUMNWIDTH% = 1100     ' width of a single grid column

Global Const PPMPERWTPERCENT# = 10000#   ' PPM per weight percent
Global Const DEFAULTDENSITY! = 5#        ' default density for sample (needs user input)

Global Const ANGPERMICRON& = 10000       ' angstroms per micron
Global Const NMPERMICRON& = 1000         ' nano-meters per micron
Global Const UMPERMICRON& = 1            ' micro-meters per micron
Global Const MMPERMICRON# = 0.001        ' milli-meters per micron
Global Const CMPERMICRON# = 0.0001       ' centi-meters per micron
Global Const MPERMICRON# = 0.000001      ' meters per micron
Global Const MICRONSPERCM& = 10000       ' microns per centimeter
Global Const MICRONSPERMM& = 1000        ' microns per millimeter

Global Const ANGPERNM& = 10              ' angstroms per nanometer
Global Const NMPERANG# = 0.1             ' nanometers per angstrom

Global Const CMPERANGSTROM# = 0.00000001 ' centi-meters per angstrom

Global Const MICROINCHESPERMICRON# = 39.37007874     ' micro-inches per micron
Global Const MILLIINCHESPERMICRON# = 0.03937007874   ' milli-inches per micron
Global Const INCHESPERMICRON# = 0.00003937           ' inches per micron

Global Const NAPA# = 1000000000#         ' nano-amps per amp
Global Const APNA# = 0.000000001         ' amps per nano-amps
Global Const PAPA# = 1000000000000#      ' pico-amps per amp
Global Const PAPERNA# = 1000#            ' pico-amps per nano-amp
Global Const NAPERMA# = 1000#            ' nano-amps per milli-amp

Global Const ANGKEV! = 12.39854          ' angstrom per KeV (and visa versa)
Global Const ANGEV! = 12398.54           ' angstrom per eV (and visa versa)
Global Const EVPERKEV# = 1000#           ' eV per keV
Global Const MILLIVOLTPERVOLT# = 1000#   ' millivolt per volt

Global Const MICROSECPERMILLSEC& = 1000# ' micro-secs per milli-sec
Global Const MSPS! = 1000000#            ' micro-secs per second
Global Const TENTHMSECPERSEC# = 10000#   ' 1/10th millsecs per second
Global Const MSECPERSEC# = 1000#         ' milli-seconds per second
Global Const SECPERMIN# = 60#            ' seconds per min
Global Const SECPERHOUR# = 3600#         ' seconds per hour
Global Const SECPERDAY# = 86400#         ' seconds per day
Global Const HOURPERDAY# = 24#           ' hours per day

Global Const CPSPERKCPS# = 1000#         ' cps per kcps

Global Const MILLIGMPERGRAM# = 1000      ' milligrams per gram

Global Const PASCALSPERTORR# = 131.578   ' Pascals per Torr conversion
Global Const PASCALSPERMBAR# = 100#      ' Pascals per mBar conversion

Global Const MAXCOUNT& = 100000000       ' default maximum counts for statistics based counting
Global Const MAXMINIMUM! = 10000000000#  ' (reversed) for calculating minimums
Global Const MAXMAXIMUM! = -10000000000# ' (reversed) for calculating maximums
Global Const MAXMINIMUM2& = 2147483647   ' (reversed) for calculating minimums
Global Const MAXMAXIMUM2& = -2147483648# ' (reversed) for calculating maximums
Global Const MAXMINIMUM3% = 32767        ' (reversed) for calculating minimums
Global Const MAXMAXIMUM3% = -32768       ' (reversed) for calculating maximums

Global Const MAXCRYSTAL2D_NOT_LDE! = 30# ' maximum 2d for non LDE crystal
Global Const MAXCRYSTAL2D_LARGE_LDE! = 100# ' maximum 2d for non large LDE crystal
Global Const ANALYTICALMAGTHRESHOLD# = 50000#    ' minimum mag for analytical warning

Global Const PI! = 3.14159               ' close enough!
Global Const PID# = 3.141592653          ' closer!
Global Const PIDD# = 3.14159265358979    ' even closer!

Global Const NATURALE# = 2.718281828     ' natural log constant
Global Const MAXLOGEXPD! = 709.6         ' maximum exponent for double precision natural log (e^MAXLOGEXPD!)
Global Const MAXLOGEXPS! = 88.721        ' maximum exponent for single precision natural log (e^MAXLOGEXPS!)

Global Const MININTEGER% = -32768        ' minimum integer value
Global Const MAXINTEGER% = 32767         ' maximum integer value
Global Const MINLONG& = -2147483648#     ' minimum long
Global Const MAXLONG& = 2147483647       ' maximum long
Global Const MINSINGLE! = -3.402823E+38  ' minimum single precision
Global Const MAXSINGLE! = 3.402823E+38   ' maximum single precision
Global Const MINDOUBLE# = -1.79769E+308  ' minimum double precision
Global Const MAXDOUBLE# = 1.79769E+308   ' maximum double precision

Global Const MAXOFFBGDTYPES% = 8         ' maximum off-peak background correction types (0 to 8, 0 = default linear) (***do not change this value***)
Global Const MAXMINTYPES% = 6            ' maximum mineral end-member types (including zero for none)
Global Const MAXZAF% = 11                ' maximum number of ZAF correction options
Global Const MAXZAFCOR% = 8              ' maximum number of stored ZAF corrections (analysis structure)
Global Const MAXLISTBOXSIZE% = 3999      ' maximum number of listbox items
Global Const MAXMACTYPE% = 7             ' maximum number of mass absorption files
Global Const MAXRELDEV% = 9999           ' maxmimum size for printout of % relative standard deviations
Global Const MAXGRIDROWS% = 4096         ' maximum number of rows allowed in grid control

Global Const ROWLAND_JEOL# = 140#        ' default JEOL focal circle
Global Const ROWLAND_CAMECA# = 160#      ' default Cameca focal circle

Global Const MAXEDS% = 48                ' maximum number of eds spectrum elements
Global Const MAXSPECTRA% = 8192          ' maximum number of eds spectrum channels
Global Const MAXSTROBE% = 128            ' maximum number of strobes (resolution calibrations)

Global Const BIT24& = 16777215           ' maximum 24 bit depth 0-16777215
Global Const BIT23& = 8388607            ' maximum 23 bit depth 0-8388607
Global Const BIT16& = 65535              ' maximum 16 bit depth 0-65535
Global Const BIT15& = 32767              ' maximum 15 bit depth 0-32767
Global Const BIT14& = 16383              ' maximum 14 bit depth 0-16383
Global Const BIT13& = 8191               ' maximum 13 bit depth 0-8191
Global Const BIT12& = 4095               ' maximum 12 bit depth 0-4095
Global Const BIT11& = 2047               ' maximum 11 bit depth 0-2047
Global Const BIT10& = 1023               ' maximum 10 bit depth 0-1023
Global Const BIT8& = 255                 ' maximum 8 bit color palette 0-255
Global Const BIT7& = 127                 ' maximum 7 bit depth 0-127
Global Const BIT6& = 63                  ' maximum 6 bit depth 0-63
Global Const BIT5& = 31                  ' maximum 5 bit depth 0-31
Global Const BIT4& = 15                  ' maximum 4 bit depth 0-15
Global Const BIT3& = 7                   ' maximum 3 bit depth 0-7
Global Const BIT2& = 3                   ' maximum 2 bit depth 0-3
Global Const BIT1& = 1                   ' maximum 1 bit depth 0-1

Global Const SIZE_65536_BYTES& = 65536   ' 65536 byte constant
Global Const SIZE_32768_BYTES& = 32768   ' 32768 byte constant
Global Const SIZE_16384_BYTES& = 16384   ' 16384 byte constant
Global Const SIZE_8192_BYTES& = 8192     ' 8192 byte constant
Global Const SIZE_4096_BYTES& = 4096     ' 4096 byte constant
Global Const SIZE_2048_BYTES& = 2048     ' 2048 byte constant
Global Const SIZE_1024_BYTES& = 1024     ' 1024 byte constant
Global Const SIZE_512_BYTES& = 512       ' 512 byte constant
Global Const SIZE_256_BYTES& = 256       ' 256 byte constant
Global Const SIZE_128_BYTES& = 128       ' 128 byte constant
Global Const SIZE_64_BYTES& = 64         ' 64 byte constant
Global Const SIZE_32_BYTES& = 32         ' 32 byte constant
Global Const SIZE_16_BYTES& = 16         ' 16 byte constant
Global Const SIZE_8_BYTES& = 8           ' 8 byte constant
Global Const SIZE_4_BYTES& = 4           ' 4 byte constant
Global Const SIZE_2_BYTES& = 2           ' 2 byte constant

Global Const IMAGESIZE128& = 128         ' image size 128 x 128 (96 @ 4/3)
Global Const IMAGESIZE256& = 256         ' image size 256 x 256 (192 @ 4/3)
Global Const IMAGESIZE512& = 512         ' image size 512 x 512 (384 @ 4/3)
Global Const IMAGESIZE1024& = 1024       ' image size 1024 x 1024 (768 @ 4/3)
Global Const IMAGESIZE2048& = 2048       ' image size 2048 x 2048 (1536 @ 4/3)

Global Const INT_ZERO% = 0               ' for parameter passing
Global Const INT_ONE% = 1                ' for parameter passing
Global Const INT_TWO% = 2                ' for parameter passing

Global Const MAXMODELS% = 6              ' maxmimum number of PTC geometric models
Global Const MAXDIAMS% = 6               ' maxmimum number of PTC particle diameters
Global Const MAXMONITOR% = 4             ' maximum number of Monitor app lists
Global Const MAXMONITORLIST% = 13        ' maximum number of Monitor app list items
Global Const MAXSTDIMAGE% = 5            ' maximum number of standard images displayed per standard number

Global Const MAXAUTOFOCUSPOINTS& = 1010  ' maximum autofocus data points (1010 + 14 = 1024 longs)
Global Const MAXAUTOFOCUSSCANS% = 3      ' maximum autofocus scans (fine, coarse, 2nd fine)
Global Const MAXIMAGES% = 2000           ' maximum number of sample images in run
Global Const MAXIMAGEIX% = 2048          ' maximum number of x pixels (SX100/SXFive = 2048, all others 1024)
Global Const MAXIMAGEIY% = 2048          ' maximum number of y pixles (SX100/SXFive = 2048, all others 1024)
Global Const MAXROMSCAN% = 1000          ' maximum points per ROM scan
Global Const MAXIMAGESIZES& = 4          ' dimensioned 0 to MAXIMAGESIZES&
Global Const MAXPALETTE% = 4             ' maximum number of color palettes

Global Const MINTAKEOFF! = 10#           ' minimum takeoff in degrees
Global Const MINKILOVOLTS! = 1#          ' minimum beam energy in kilovolts
Global Const MINBEAMCURRENT! = 0.01      ' minimum beam current in nA
Global Const MINBEAMSIZE! = 0#           ' minimum beam size in microns

Global Const MAXTAKEOFF! = 75            ' maximum takeoff in degrees
Global Const MAXKILOVOLTS! = 40#         ' maximum beam energy in kilovolts
Global Const MAXBEAMCURRENT! = 1000#     ' maximum beam current in nA
Global Const MAXBEAMSIZE! = 50#          ' maximum beam size in microns

Global Const MINREPLICATES% = 1          ' minimum number of replicate samples
Global Const MAXREPLICATES% = 500        ' maximum number of replicates samples

Global Const MAXBEAMCALIBRATIONS% = 32   ' maximum number of beam scan calibration array members

' Integrated intensity constants
Global Const MININTEGRATEDINTENSITYINITIALSTEP! = 800#
Global Const MAXINTEGRATEDINTENSITYINITIALSTEP! = 10#
Global Const MININTEGRATEDINTENSITYMINIMUMSTEP! = 10000#
Global Const MAXINTEGRATEDINTENSITYMINIMUMSTEP! = 20#
Global Const INTEGRATEDINTENSITYINITIALSTEP! = 100#
Global Const INTEGRATEDINTENSITYMINIMUMSTEP! = 400#

Global Const SCALERSDATLINESNEW% = 90               ' number of lines in new SCALERS.DAT file

Global Const MAXSCANS& = 99999                      ' maximum number of PHA, bias, gain or peaking scans
Global Const MAXAPERTURES% = 4                      ' maximum number of apertures in system
Global Const SMALLAMOUNTFRACTION_HALF! = 0.0005     ' small amount fraction for motion limits (times 0.5)
Global Const SMALLAMOUNTFRACTION! = 0.001           ' small amount fraction for motion limits
Global Const SMALLAMOUNTFRACTIONx5! = 0.005         ' small amount fraction for motion limits (times 5)
Global Const SMALLAMOUNTFRACTIONx10! = 0.01         ' small amount fraction for motion limits (times 10)
Global Const MAXROMPEAKTYPES% = 6                   ' maximum number of ROM scan peaking types
Global Const DEFAULTVOLATILEINTERVALS% = 5          ' default number of volatile intervals

Global Const JEOL_SPECTRO_JOG_SIZE# = 0.6           ' additional room for JEOL spectro low limit jog (in mm)
Global Const JEOL_IMAGE_SHIFT_FACTOR! = 163.8       ' image_shift = um_shift * 163.8

Global Const MAXKLMORDER% = 20           ' maximum number of KLM higher order lines to display
Global Const MAXKLMORDERCHAR% = 5        ' maximum number of KLM higher order Roman characters to display
Global Const MAXSAMPLETYPES% = 3         ' maximumn types of samples (standard, unknown, wavescan)
Global Const MAXFORBIDDEN% = 20          ' maximum number of forbidden elements

Global Const NOT_ANALYZED_VALUE_SINGLE! = 0.00000001       ' 10^-8
'Global Const NOT_ANALYZED_VALUE_DOUBLE# = 0.00000001       ' 10^-8     (do not utilize as it causes a problem for Pro Essentials with .NullDataValueX property)

Global Const FeO_OXIDE_PROPORTION! = 0.286497     ' Fe to FeO stoichiometry (see parameter z.p1!() in ZAF.BAS)

Global Const MAX_EXCEL_2003_COLS% = 256  ' maximum number of columns supported by Excel 2003 (version 11)

Global Const LOTSOFGRIDPOINTS% = 1000    ' lots of polygon points (hide FormAUTOMATE)
Global Const TOOMANYGRIDSTEPS% = 2000    ' too many grid steps
Global Const MAXTITLELENGTH% = 80        ' maximum graph title length
Global Const MAX_PATH% = 259             ' maximum file and path length for Dir$ command

Global Const FONT_REGULAR% = 0           ' regular format
Global Const FONT_BOLD% = 1              ' bold
Global Const FONT_ITALIC% = 2            ' italic
Global Const FONT_UNDERLINE% = 4         ' underline
Global Const FONT_STRIKETHRU% = 8        ' strikethru

' Database Field lengths
Global Const DbTextDescriptionLength% = 255
Global Const DbTextDescriptionOldLength% = 64

Global Const DbTextFilenameLengthNew% = 255
Global Const DbTextFilenameLength% = 128
Global Const DbTextFilenameOldLength% = 64

Global Const DbTextAcquireStringLength% = 128
Global Const DbTextAcquireStringOldLength% = 64

Global Const DbTextTitleStringLength% = 128
Global Const DbTextUserNameLength% = 128
Global Const DbTextNameLength% = 64
Global Const DbTextEmpiricalStringLength% = 48
Global Const DbTextIPAddressLength% = 32
Global Const DbTextDetectorStringLength% = 32
Global Const DbTextCrystalStringLength% = 12
Global Const DbTextFaradayUnitsLength% = 8
Global Const DbTextMotorStringLength% = 8
Global Const DbTextImageChannelLength% = 8
Global Const DbTextXrayStringLength% = 2
Global Const DbTextElementStringLength% = 2
Global Const DbTextDefaultImageAnalogUnitsLength% = 32
Global Const DbTextKratiosDATLineLength% = 164
Global Const DbTextFormulaStringLength% = 128
Global Const DbTextInstrumentAcknowledgementLength% = 128

Global Const DbTextMemoStringLength& = 65535

' Default format specifiers
Global Const a20$ = "@@"
Global Const a30$ = "@@@"
Global Const a40$ = "@@@@"
Global Const a50$ = "@@@@@"
Global Const a60$ = "@@@@@@"
Global Const a70$ = "@@@@@@@"
Global Const a80$ = "@@@@@@@@"
Global Const a90$ = "@@@@@@@@@"
Global Const a100$ = "@@@@@@@@@@"
Global Const a120$ = "@@@@@@@@@@@@"
Global Const a140$ = "@@@@@@@@@@@@@@"

Global Const a4x$ = "    "
Global Const a6x$ = "      "
Global Const a8x$ = "        "

' Floating point formats
Global Const f40$ = "###."
Global Const f41$ = "##.0"
Global Const f42$ = "#.00"
Global Const f43$ = ".000"

Global Const f50$ = "####."
Global Const f51$ = "###.0"
Global Const f52$ = "##.00"
Global Const f60$ = "#####."
Global Const f61$ = "####.0"
Global Const f62$ = "###.00"
Global Const f63$ = "##.000"
Global Const f64$ = "#.0000"
Global Const f65$ = ".00000"

Global Const f80$ = "#######."
Global Const f81$ = "######.0"
Global Const f82$ = "#####.00"
Global Const f83$ = "####.000"
Global Const f84$ = "###.0000"
Global Const f85$ = "##.00000"
Global Const f86$ = "#.000000"
Global Const f87$ = ".0000000"

Global Const f94$ = "####.0000"
Global Const f95$ = "###.00000"
Global Const f96$ = "##.000000"

Global Const f100$ = "#########."
Global Const f101$ = "########.0"
Global Const f102$ = "#######.00"
Global Const f103$ = "######.000"
Global Const f104$ = "#####.0000"
Global Const f105$ = "####.00000"
Global Const f106$ = "###.000000"
Global Const f107$ = "##.0000000"
Global Const f108$ = "#.00000000"
Global Const f109$ = ".000000000"

Global Const f120$ = "###########."
Global Const f121$ = "##########.0"
Global Const f122$ = "#########.00"
Global Const f123$ = "########.000"
Global Const f124$ = "#######.0000"
Global Const f125$ = "######.00000"
Global Const f126$ = "#####.000000"
Global Const f127$ = "####.0000000"

Global Const i20$ = "#0"
Global Const i30$ = "##0"
Global Const i40$ = "###0"
Global Const i50$ = "####0"
Global Const i60$ = "#####0"
Global Const i80$ = "#######0"

' Exponential formats
Global Const e50$ = "#e+00"
Global Const e61$ = "#0e+00"
Global Const e71$ = "#.0e+00"
Global Const e82$ = "#.00e+00"
Global Const e104$ = "#.0000e+00"
Global Const e115$ = "#.00000e+00"
Global Const e125$ = "#.000000e+00"
Global Const e137$ = "+.0000000e+00;-.0000000e+00"

Global Const DASHED3$ = "---"
Global Const DASHED4$ = "----"
Global Const DASHED5$ = "-----"

Global Const EDS_CRYSTAL$ = "EDS"        ' EDS crystal (not WDS)
Global Const CONTINUED$ = "continued"

Global Const ATOMIC_NUM_HYDROGEN% = 1
Global Const ATOMIC_NUM_HELIUM% = 2
Global Const ATOMIC_NUM_LITHIUM% = 3
Global Const ATOMIC_NUM_BERYLLIUM% = 4
Global Const ATOMIC_NUM_BORON% = 5
Global Const ATOMIC_NUM_CARBON% = 6
Global Const ATOMIC_NUM_NITROGEN% = 7
Global Const ATOMIC_NUM_OXYGEN% = 8
Global Const ATOMIC_NUM_FLUORINE% = 9
Global Const ATOMIC_NUM_NEON% = 10
Global Const ATOMIC_NUM_SODIUM% = 11
Global Const ATOMIC_NUM_MAGNESIUM% = 12
Global Const ATOMIC_NUM_ALUMINUM% = 13
Global Const ATOMIC_NUM_SILICON% = 14
Global Const ATOMIC_NUM_PHOSPHORUS% = 15
Global Const ATOMIC_NUM_SULFUR% = 16
Global Const ATOMIC_NUM_CHLORINE% = 17
Global Const ATOMIC_NUM_ARGON% = 18
Global Const ATOMIC_NUM_POTASSIUM% = 19
Global Const ATOMIC_NUM_CALCIUM% = 20
Global Const ATOMIC_NUM_SCANDIUM% = 21
Global Const ATOMIC_NUM_TITANIUM% = 22
Global Const ATOMIC_NUM_VANADIUM% = 23
Global Const ATOMIC_NUM_CHROMIUM% = 24
Global Const ATOMIC_NUM_MANGANESE% = 25
Global Const ATOMIC_NUM_IRON% = 26
Global Const ATOMIC_NUM_COBALT% = 27
Global Const ATOMIC_NUM_NICKEL% = 28
Global Const ATOMIC_NUM_COPPER% = 29
Global Const ATOMIC_NUM_ZINC% = 30
Global Const ATOMIC_NUM_BROMINE% = 35
Global Const ATOMIC_NUM_KRYPTON% = 36
Global Const ATOMIC_NUM_STRONTIUM% = 38
Global Const ATOMIC_NUM_ZIRCONIUM% = 40
Global Const ATOMIC_NUM_MOLYBDENUM% = 42
Global Const ATOMIC_NUM_IODINE% = 53
Global Const ATOMIC_NUM_XENON% = 54
Global Const ATOMIC_NUM_LEAD% = 82
Global Const ATOMIC_NUM_RADON% = 86
Global Const ATOMIC_NUM_THORIUM% = 90
Global Const ATOMIC_NUM_URANIUM% = 92

Type TypeXray
    atnum As Integer
    syme As String * 2
    n As Integer
    xline As String * 8
    abedg As String * 3
    xwave As Double
    xints As Double
    refer As String * 5
End Type

' User defined data types
Type TypeEnergy
    energy(1 To MAXRAY_OLD%) As Single
End Type

Type TypeEdge
    energy(1 To MAXEDG%) As Single
End Type

Type TypeFlur
    fraction(1 To MAXRAY_OLD%) As Single
End Type

Type TypeMu
    mac(1 To MAXELM% * MAXRAY_OLD%) As Single
End Type

Type TypeXraySymbols
    syms(1 To MAXRAY_OLD%) As String * 2
End Type

Type TypeEdgeSymbols
    syms(1 To MAXEDG%) As String * 2
End Type

' Average arrays
Type TypeAverage
    averags(1 To MAXCHAN1%) As Single   ' average
    Stddevs(1 To MAXCHAN1%) As Single   ' standard deviation
    Sqroots(1 To MAXCHAN1%) As Single   ' square root
    Stderrs(1 To MAXCHAN1%) As Single   ' standard error
    Reldevs(1 To MAXCHAN1%) As Single   ' relative standard deviation
    Minimums(1 To MAXCHAN1%) As Single  ' minimum
    Maximums(1 To MAXCHAN1%) As Single  ' maximum
    AverDateTime As Double              ' average datetime (count arrays only)
End Type

Type TypeAverageMathSingle
    averags() As Single   ' average
    Stddevs() As Single   ' standard deviation
    Sqroots() As Single   ' square root
    Stderrs() As Single   ' standard error
    Reldevs() As Single   ' relative standard deviation
    Minimums() As Single  ' minimum
    Maximums() As Single  ' maximum
End Type

Type TypeAverageMathDouble
    averags() As Double   ' average
    Stddevs() As Double   ' standard deviation
    Sqroots() As Double   ' square root
    Stderrs() As Double   ' standard error
    Reldevs() As Double   ' relative standard deviation
    Minimums() As Double  ' minimum
    Maximums() As Double  ' maximum
End Type

' Position array
Type TypePosition
    samplerow As Integer
    sampletype As Integer
    samplenumber As Integer
    samplename As String
    sampledescription As String
    takeoff As Single
    kilovolts As Single
    beamcurrent As Single
    beamsize As Single
    
    FiducialSet As Integer
    FiducialDescription As String
    
    SampleSetupNumber As Integer
    FileSetupName As String
    FileSetupNumber As Integer
    
    Magnification As Single
    BeamCenterXYZ(1 To MAXAXES%) As Single     ' stage coordinate for center of image
    
    MultipleSetupNumber As Integer
    MultipleSetupNumbers() As Integer
    
    ColumnConditionMethod As Integer
    ColumnConditionString As String
    Replicates As Integer
    
    beammode As Integer ' 0 = analog spot, 1 = analog scan, 2 = digital spot
    magnificationanalytical As Single       ' new 10-28-2006
    magnificationimaging As Single          ' new 10-28-2006
    
    ImageShiftX As Single    ' change from integer for SX100/SXFive (10-29-2011)
    ImageShiftY As Single    ' change from integer for SX100/SXFive (10-29-2011)
    
    DriftCorrectionImageNumber As Integer   ' stored image number for drift correction (ImageNumber in Image table)
End Type

' TDI (volatile and alternating on/off peak) structure
Type TypeVolatile
    VolatilePoints As Integer
    VolatileDateTime() As Double       ' time of day (in days)
    VolatileIntensity() As Single      ' x-ray intensity
    VolatileInterval() As Single       ' time interval
    VolatileOnAbsorbed() As Single       ' absorbed current
    VolatileHiAbsorbed() As Single       ' absorbed current
    VolatileLoAbsorbed() As Single       ' absorbed current
End Type

' Position data array
Type TypePositionData
    xyz(1 To MAXAXES%) As Single
    grainnumber As Integer
    autofocus As Integer
End Type

' Realtime monitor array
Type TypeMonitorStructure
    xs As Integer
    ys As Integer
    xt As Integer
    yt As Integer
    
    CondenserCoarse As Long
    CondenserFine As Long
    ObjectiveCoarse As Long
    ObjectiveFine As Long
    Astigmation1 As Long
    Astigmation2 As Long
    ScanRotation As Single
    
    KilovoltStatus As String
    kilovolts As Single
    EmissionCurrent As Single
    FilamentCurrent As Single
    Magnification As Single
    takeoff As Single
    beamcurrent As Single
    beamsize As Single
    beamonstatus As Boolean  ' true = beam on
    
    scanmode As Integer ' 0 = scan, 1 = spot, 2 = digital
    reflectstatus As Integer        ' true = reflected light on
    transmitstatus As Integer       ' true = transmitted light on
    
    crystalpositions(1 To MAXSPEC%) As String
    counts(1 To MAXSPEC%) As Single
    motorpositions(1 To MAXMOT%) As Single
    
    phabaselines(1 To MAXSPEC%) As Single
    phawindows(1 To MAXSPEC%) As Single
    phagains(1 To MAXSPEC%) As Single
    phabiases(1 To MAXSPEC%) As Single
    phamodes(1 To MAXSPEC%) As Integer
End Type

' Analysis arrays
Type TypeAnalysis
    TotalPercent As Single       ' variable loaded from routine ZAFCalZbar
    totaloxygen As Single        ' variable loaded from routine ZAFCalZbar
    TotalCations As Single       ' variable loaded from routine ZAFCalZbar
    totalatoms As Single         ' variable loaded from routine ZAFCalZbar
    CalculatedOxygen As Single   ' variable loaded from routine ZAFCalZbar
    ExcessOxygen As Single       ' variable loaded from routine ZAFCalZbar
    zbar As Single               ' variable loaded from routine ZAFCalZbar
    AtomicWeight As Single             ' variable loaded from routine ZAFCalZbar
    OxygenFromHalogens As Single       ' variable loaded from routine ZAFCalZbar
    HalogenCorrectedOxygen As Single   ' variable loaded from routine ZAFCalZbar
    ChargeBalance As Single            ' variable loaded from routine ZAFCalZbar
    FeCharge As Single                 ' variable loaded from routine ZAFCalZbar
    
    OxygenFromSulfur As Single         ' variable loaded from routine ZAFCalZbar
    SulfurCorrectedOxygen As Single    ' variable loaded from routine ZAFCalZbar
    
    ZAFIter As Single
    MANIter As Single
    
    WtsData() As Single ' calculated elemental weight percents (allocated in InitStandards) (1 To MAXROW%, 1 To MAXCHAN1%)
    CalData() As Single ' calculated oxide/atomic/formula/etc (allocated in InitStandards) (1 To MAXROW%, 1 To MAXCHAN1%)
    
    UnkZAFCors() As Single  ' (allocated in InitStandards) 1 To MAXZAFCOR%, 1 To MAXCHAN%
    Elsyms(1 To MAXCHAN%) As String
    Xrsyms(1 To MAXCHAN%) As String
    MotorNumbers(1 To MAXCHAN%) As Integer
    CrystalNames(1 To MAXCHAN%) As String
    
    AtomicNumbers(1 To MAXCHAN%) As Single
    AtomicCharges(1 To MAXCHAN%) As Single
    AtomicWts(1 To MAXCHAN%) As Single
      
    WtPercents(1 To MAXCHAN%) As Single
    Formulas(1 To MAXCHAN%) As Single
    OxPercents(1 To MAXCHAN%) As Single
    AtPercents(1 To MAXCHAN%) As Single
    OxMolPercents(1 To MAXCHAN%) As Single
    ElPercents(1 To MAXCHAN%) As Single
    NormElPercents(1 To MAXCHAN%) As Single
    NormOxPercents(1 To MAXCHAN%) As Single
    
    UnkKrats(1 To MAXCHAN%) As Single
    UnkBetas(1 To MAXCHAN%) As Single   ' alpha factor calculations only
    UnkMACs(1 To MAXCHAN%) As Single

    StdZAFCors() As Single      ' allocated in InitStandards (1 To MAXZAFCOR%, 1 To MAXSTD%, 1 To MAXCHAN%)
    StdBetas() As Single        ' allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)
        
    StdPercents() As Single     ' allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)
    StdZbars(1 To MAXSTD%) As Single
    StdMACs() As Single         ' allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)
    
    StdAtomicCharges() As Single        ' allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)      ' v. 13.3.2
    StdAtomicWts() As Single            ' allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)      ' v. 13.3.2
    
    StdAssignsCounts(1 To MAXCHAN%) As Single
    StdAssignsTimes(1 To MAXCHAN%) As Single
    StdAssignsBeams(1 To MAXCHAN%) As Single
    StdAssignsKfactors(1 To MAXCHAN%) As Single
    StdAssignsZAFCors() As Single                   ' allocated in InitStandards (1 To MAXZAFCOR%, 1 To MAXCHAN%)
    StdAssignsBetas(1 To MAXCHAN%) As Single        ' alpha factor calculations only
    StdAssignsPercents(1 To MAXCHAN%) As Single     ' alpha factor calculations only
    StdAssignsRows(1 To MAXCHAN%) As Integer
    StdAssignsZbars(1 To MAXCHAN%) As Single
    StdAssignsBgdCounts(1 To MAXCHAN%) As Single
    
    StdAssignsIntfCounts() As Single    ' allocated in InitStandards (1 To MAXINTF%, 1 To MAXCHAN%)
    StdAssignsIntfRows() As Integer     ' allocated in InitStandards (1 To MAXINTF%, 1 To MAXCHAN%)

    StdAssignsActualKilovolts(1 To MAXCHAN%) As Single    ' in keV (includes beam energy loss from coating if specified)
    StdAssignsEdgeEnergies(1 To MAXCHAN%) As Single       ' in keV
    StdAssignsActualOvervoltages(1 To MAXCHAN%) As Single ' includes beam energy loss from coating if specified
       
    Coating_StdAssignsTrans(1 To MAXCHAN%) As Single      ' standard coating x-ray transmission
    Coating_StdAssignsAbsorbs(1 To MAXCHAN%) As Single    ' standard coating electron absorption
    
    ActualKilovolts(1 To MAXCHAN%) As Single              ' in keV (includes energy loss from coating correction if specified)
    EdgeEnergies(1 To MAXCHAN%) As Single                 ' in keV
    ActualOvervoltages(1 To MAXCHAN%) As Single           ' includes energy loss from coating correction if specified

    MANFitCoefficients() As Single      ' allocated in InitStandards (1 To MAXCOEFF%, 1 To MAXCHAN%)
    MANAssignsCounts() As Single        ' allocated in InitStandards (1 To MAXMAN%, 1 To MAXCHAN%)
    MANAssignsRows() As Integer         ' allocated in InitStandards (1 To MAXMAN%, 1 To MAXCHAN%)

    MANAssignsCountTimes() As Single     ' allocated in InitStandards (1 To MAXMAN%, 1 To MAXCHAN%)
    MANAssignsBeamCurrents() As Single   ' allocated in InitStandards (1 To MAXMAN%, 1 To MAXCHAN%)

    UnkContinuumCorrections(1 To MAXCHAN%) As Single      ' Heinrich/Myklebust continuum corrections for unknown
    StdContinuumCorrections() As Single     ' Heinrich/Myklebust continuum corrections for stds, allocated in InitStandards (1 To MAXSTD%, 1 To MAXCHAN%)

    SampleIsAnalyzed As Boolean
    
    FerricToTotalIronRatio As Single         ' ferric to total iron ratio
    FerricFerrousFeO As Single           ' total FeO
    FerricFerrousFe2O3 As Single         ' total Fe2O3
    FerricOxygen As Single               ' oxygen from Fe2O3 (not including specified excess oxygen)
End Type

' Sample arrays
Type TypeSample
    number As Integer
    Set As Integer
    Type As Integer
    Name As String
    
    SampleSetupNumber As Integer  ' for storing analytical sample setup number
    MultipleSetupNumber As Integer  ' for storing analytical multiple sample setup number (combined only)
    MultipleSetupNumbers() As Integer  ' for storing analytical multiple sample setup numbers (combined only)
    FileSetupName As String  ' for storing analytical file setup name
    FileSetupNumber As Integer  ' for storing analytical file setup number
    
    VolatileAcquisitionType As Integer  ' new for v. 4.84 (1 = self, 2 = assigned)
    WavescanAcquisitionType As Integer  ' new for v. 5.03 (1 = normal, 2 = quick, 3 = normal ROM, 4 = quick ROM)
    
    takeoff As Single
    kilovolts As Single
    beamcurrent As Single
    beamsize As Single
    ColumnConditionMethod As Integer    ' 0 = TKCS, 1 = condition string
    ColumnConditionString As String
    ApertureNumber As Integer           ' aperture number (JEOL only)
    
    PreAcquireString As String          ' column string applied during sample acquisition
    PostAcquireString As String         ' column string applied after sample acquisition
    
    CombinedConditionsFlag As Integer
    TakeoffArray(1 To MAXCHAN%) As Single                   ' for multiple conditions
    KilovoltsArray(1 To MAXCHAN%) As Single                 ' for multiple conditions
    BeamCurrentArray(1 To MAXCHAN%) As Single               ' for multiple conditions
    BeamSizeArray(1 To MAXCHAN%) As Single                  ' for multiple conditions
    ColumnConditionMethodArray(1 To MAXCHAN%) As Integer    ' for multiple conditions
    ColumnConditionStringArray(1 To MAXCHAN%) As String     ' for multiple conditions
    
    Description As String
    OxideOrElemental As Integer
    DisplayAsOxideFlag As Integer
    AtomicPercentFlag As Integer

    DifferenceElementFlag As Integer
    DifferenceElement As String
    DifferenceFormulaFlag As Integer
    DifferenceFormula As String
    StoichiometryElementFlag As Integer
    StoichiometryElement As String
    StoichiometryRatio As Single
    RelativeElementFlag As Integer
    RelativeElement As String
    RelativeToElement As String
    RelativeRatio As Single
    FormulaElementFlag As Integer
    FormulaElement As String
    FormulaRatio As Single
    MineralFlag As Integer
    DetectionLimitsFlag As Integer
    DetectionLimitsProjectedFlag As Integer
    HomogeneityFlag As Integer
    HomogeneityAlternateFlag As Integer
    CorrelationFlag As Integer
    DisplayAmphiboleCalculationFlag As Integer
    DisplayBiotiteCalculationFlag As Integer
    HydrogenStoichiometryFlag As Integer
    HydrogenStoichiometryRatio As Single
    
    CoatingFlag As Integer          ' 0 = not coated, 1 = coated
    CoatingElement As Integer
    CoatingDensity As Single
    CoatingThickness As Single      ' in angstroms
    CoatingSinThickness As Single   ' in angstroms (x-ray absorption path length)

    AlternatingOnAndOffPeakAcquisitionFlag As Integer
    
    EDSSpectraFlag As Integer                ' EDS spectrum data is stored in EDS Spectra table
    EDSSpectraUseFlag As Integer             ' use EDS spectrum data in quant calculations
    
    EDSUnknownCountFactors() As Single       ' real time, allocated in InitSample (1 to MAXROW%)
    LastEDSUnknownCountFactor As Single
    LastEDSSpecifiedCountTime As Single
    
    EDSSpectraIntensities() As Long          ' allocated in InitSample (1 to MAXROW%, 1 to MAXSPECTRA%)
    EDSSpectraStrobes() As Long              ' allocated in InitSample (1 to MAXROW%, 1 to MAXSTROBE%) (Oxford only)
    
    EDSSpectraElapsedTime() As Single        ' real time, allocated in InitSample (1 to MAXROW%)
    EDSSpectraDeadTime() As Single           ' dead time as percentage
    EDSSpectraSampleTime() As Single         ' sample counting time (estimated)
    EDSSpectraLiveTime() As Single           ' actual count integration time
    EDSSpectraNumberofChannels() As Integer  ' in spectrum
    EDSSpectraNumberofStrobes() As Integer   ' in strobe (new 11/06/04)
    
    EDSSpectraEVPerChannel() As Single
    EDSSpectraTakeOff() As Single
    EDSSpectraAcceleratingVoltage() As Single   ' in keV
    EDSSpectraStartEnergy() As Single           ' in keV
    EDSSpectraEndEnergy() As Single             ' in keV
    EDSSpectraMaxCounts() As Long
    EDSSpectraADCTimeConstant() As Single       ' pulse processing time (in vendor units)
    
    EDSSpectraDetectorSubtype() As Long           ' used by Thermo PF 2.10 and higher only for spectrum processing only
    EDSSpectraZeroWidth() As Double               ' used by Thermo PF 2.10 and higher only
    EDSSpectraTimeConstant() As Long              ' used by Thermo PF 2.10 and higher only
    
    EDSSpectraKLineBCoefficient() As Single     ' used by Bruker only
    EDSSpectraKLineCCoefficient() As Single     ' used by Bruker only
    
    EDSSpectraQuantMethodOrProject As String    ' only used by Bruker
   
    EDSSpectraEDSFileName() As String           ' only used by JEOL OEM EDS
    
    FiducialSetNumber As Integer        ' data is stored in Fiducial table
    FiducialSetDescription As String        ' data is stored in Fiducial table
    fiducialpositions(1 To MAXAXES%, 1 To MAXDIM%) As Single

    Linenumber(1 To MAXROW%) As Long        ' changed from integer 11/11/05 to handle more than 32K data points
    LineStatus(1 To MAXROW%) As Integer
    DateTimes(1 To MAXROW%) As Double
    StagePositions() As Single              ' allocated in InitSample (1 To MAXROW%, 1 To MAXAXES%)

    OnBeamCounts(1 To MAXROW%) As Single
    AbBeamCounts(1 To MAXROW%) As Single
    OnBeamCountsArray() As Single           ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    AbBeamCountsArray() As Single           ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)

    OnBeamCounts2(1 To MAXROW%) As Single
    AbBeamCounts2(1 To MAXROW%) As Single
    OnBeamCountsArray2() As Single          ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    AbBeamCountsArray2() As Single          ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)

    Elsyms(1 To MAXCHAN%) As String
    Xrsyms(1 To MAXCHAN%) As String
    BraggOrders(1 To MAXCHAN%) As Integer           ' x-ray Bragg order (I, II, II, IV, ect)
    numcat(1 To MAXCHAN%) As Integer
    numoxd(1 To MAXCHAN%) As Integer
    ElmPercents(1 To MAXCHAN%) As Single
    OffPeakCorrectionTypes(1 To MAXCHAN%) As Integer    ' 0=linear, 1=average, 2=high only, 3=low only, 4=exponential, 5=slope hi, 6=slope lo, 7=polynomial, 8=multi-point

    MotorNumbers(1 To MAXCHAN%) As Integer
    OrderNumbers(1 To MAXCHAN%) As Integer          ' spectrometer order number (acquisition order)
    CrystalNames(1 To MAXCHAN%) As String
    Crystal2ds(1 To MAXCHAN%) As Single
    CrystalKs(1 To MAXCHAN%) As Single
    OnPeaks(1 To MAXCHAN%) As Single
    HiPeaks(1 To MAXCHAN%) As Single
    LoPeaks(1 To MAXCHAN%) As Single
    StdAssigns(1 To MAXCHAN%) As Integer
    StdAssignsFlag(1 To MAXCHAN%) As Integer        ' 0 = normal, 1 = virtual
    DisableQuantFlag(1 To MAXCHAN%) As Integer      ' 0 = enabled, 1 = disabled
    DisableAcqFlag(1 To MAXCHAN%) As Integer        ' 0 = enabled, 1 = disabled

    BackgroundTypes(1 To MAXCHAN%) As Integer       ' 0 = off-peak, 1 = MAN, 2 = MPB

    Baselines(1 To MAXCHAN%) As Single
    Windows(1 To MAXCHAN%) As Single
    Gains(1 To MAXCHAN%) As Single
    Biases(1 To MAXCHAN%) As Single
    InteDiffModes(1 To MAXCHAN%) As Integer
    DeadTimes(1 To MAXCHAN%) As Single
    
    DetectorSlitSizes(1 To MAXCHAN%) As String      ' new in v. 5.15
    DetectorSlitPositions(1 To MAXCHAN%) As String
    DetectorModes(1 To MAXCHAN%) As String

    StdAssignsIntfElements()  As String     ' interfering element, allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    StdAssignsIntfXrays()  As String        ' interfering x-ray (for channel ID only), allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    StdAssignsIntfStds()  As Integer        ' interference standard, allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    StdAssignsIntfOrders()  As Integer      ' order of interfering line (for matrix correction adjustment), allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)

    VolatileCorrectionUnks(1 To MAXCHAN%) As Integer
    SpecifiedAreaPeakFactors(1 To MAXCHAN%) As Single   ' not calculated and on emitter basis

    LastCountFactors(1 To MAXCHAN%) As Single   ' change to single (4/24/02)
    LastMaxCounts(1 To MAXCHAN%) As Long
    LastOnCountTimes(1 To MAXCHAN%) As Single
    LastHiCountTimes(1 To MAXCHAN%) As Single
    LastLoCountTimes(1 To MAXCHAN%) As Single
    LastWaveCountTimes(1 To MAXCHAN%) As Single
    LastPeakCountTimes(1 To MAXCHAN%) As Single
    LastQuickCountTimes(1 To MAXCHAN%) As Single
    
    UnknownCountFactors() As Single     ' changed to single (4/24/02)
    UnknownMaxCounts() As Long          ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    
    OnPeakCounts_Raw_Cps() As Single    ' raw cps data, allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    HiPeakCounts_Raw_Cps() As Single    ' raw cps data, allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    LoPeakCounts_Raw_Cps() As Single    ' raw cps data, allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)

    OnPeakCounts() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    HiPeakCounts() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    LoPeakCounts() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    
    OnCountTimes() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    HiCountTimes() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    LoCountTimes() As Single    ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    
    VolCountTimesStart() As Variant     ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    VolCountTimesStop() As Variant      ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
    VolCountTimesDelay() As Single      ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%)
        
    LastElm As Integer          ' calculated in DataGetMDBSample
    LastChan As Integer
    Datarows As Integer
    GoodDataRows As Integer
    AllMANBgdFlag As Integer
    MANBgdFlag As Integer

    CorData() As Single     ' corrected/normalized on-peak count data (allocated in DataGetMDBSample)
    BgdData() As Single     ' corrected/normalized off-peak count data (allocated in DataGetMDBSample)
    ErrData() As Single     ' corrected/normalized off-peak count data (allocated in DataGetMDBSample)

    OnTimeData() As Single      ' on-peak count time data (allocated in DataGetMDBSample)
    HiTimeData() As Single      ' hi-peak count time data (allocated in DataGetMDBSample)
    LoTimeData() As Single      ' lo-peak count time data (allocated in DataGetMDBSample)
    
    OnBeamData() As Single          ' on beam data (allocated in DataGetMDBSample)
    OnBeamDataArray() As Single     ' on beam data array (allocated in DataGetMDBSample)
        
    AggregateNumChannels() As Integer       ' number of aggregate channels (allocated in DataGetMDBSample)
        
    Offsets(1 To MAXCHAN%) As Single        ' loaded in XrayGetOffsets
    LineEnergy(1 To MAXCHAN%) As Single     ' loaded in ElementCheckXray
    LineEdge(1 To MAXCHAN%) As Single       ' loaded in ElementCheckXray
    OxygenChannel As Integer                ' loaded in ZAFSetZAF for oxide calculations
    
    AtomicCharges(1 To MAXCHAN%) As Single      ' loaded in ElementLoadArrays
    AtomicWts(1 To MAXCHAN%) As Single          ' loaded in ElementLoadArrays (for natural/isotope enriched unknowns)
    
    AtomicNums(1 To MAXCHAN%) As Integer    ' loaded in ElementLoadArrays
    XrayNums(1 To MAXCHAN%) As Integer      ' loaded in ElementLoadArrays
    Oxsyup(1 To MAXCHAN%) As String         ' loaded in ElementLoadArrays
    Elsyup(1 To MAXCHAN%) As String         ' loaded in ElementLoadArrays
    
    ' Loaded from MAN database table
    MANStdAssigns() As Integer                  '  allocated in InitSample (1 To MAXMAN%, 1 To MAXCHAN%)
    MANLinearFitOrders() As Integer             '  allocated in InitSample (1 To MAXCHAN%)
    MANAbsCorFlags() As Integer                 '  allocated in InitSample (1 To MAXCHAN%)

    ' Calculated in VolatileCalculateFit- Time Dependent Intensity (TDI)
    VolatileFitIntercepts(1 To MAXCHAN%) As Single      ' intercepts (for display only)
    VolatileFitSlopes(1 To MAXCHAN%) As Single          ' slope coefficients
    VolatileFitCurvatures(1 To MAXCHAN%) As Single      ' curvature coefficients
    VolatileFitAvgDev(1 To MAXCHAN%) As Single          ' average relative deviation fit
    VolatileFitTypes(1 To MAXCHAN%) As Integer          ' linear or quadratic fit
    VolatileFitAvgTime(1 To MAXCHAN%) As Single         ' average elapsed time
    
    BackgroundExponentialBase() As Single                      '  allocated in InitSample (1 To MAXCHAN%)
    BackgroundSlopeCoefficients() As Single                    ' 1 = High, 2 = Low, allocated in InitSample (1 To 2, 1 To MAXCHAN%)
    BackgroundPolynomialPositions() As Single                  '  allocated in InitSample (1 To MAXCOEFF%, 1 To MAXCHAN%)
    BackgroundPolynomialCoefficients() As Single               '  allocated in InitSample (1 To MAXCOEFF%, 1 To MAXCHAN%)
    BackgroundPolynomialNominalBeam() As Single                '  allocated in InitSample (1 To MAXCHAN%)
    
    BeamDeflectionFlag As Integer
    Magnification As Single
    magnificationanalytical As Single           ' new 10-28-2006
    magnificationimaging As Single              ' new 10-28-2006
    BeamCenterXYZ(1 To MAXAXES%) As Single
    BeamXSize As Single
    BeamYSize As Single
    beammode As Integer     ' 0 = analog spot, 1 = analog scan, 2 = digital spot
    
    IntegratedIntensitiesUseIntegratedFlags(1 To MAXCHAN%) As Integer       ' channel flag to acquire or contains integrated data
    IntegratedIntensitiesIntegratedTypes(1 To MAXCHAN%) As Integer          ' channel integrated intensity type (wavescans only)
    IntegratedIntensitiesInitialStepSizes(1 To MAXCHAN%) As Single
    IntegratedIntensitiesMinimumStepSizes(1 To MAXCHAN%) As Single
    
    IntegratedIntensitiesFlag As Integer            ' sample flag to indicate integrated data is stored in Inte table
    IntegratedIntensitiesUseFlag As Integer         ' sample flag to indicate use integrated data in calculations
    IntegratedPoints() As Integer                   ' allocated in DataGetMDBSample (row, chan)
    IntegratedPositions() As Single                 ' allocated in DataGetMDBSample (row, chan, n)
    IntegratedIntensities() As Single               ' allocated in DataGetMDBSample (row, chan, n)
    IntegratedCountTimes() As Single                ' allocated in DataGetMDBSample (row, chan, n)
    IntegratedPeakIntensities() As Single           ' allocated in DataGetMDBSample (row, chan)
    
    Replicates As Integer           ' number of point replicates acquired
    
    iptc As Integer                 ' particle and thin film options (0=no, 1=yes)
    PTCModel As Integer
    PTCDiameter As Single
    PTCDensity As Single
    PTCThicknessFactor As Single
    PTCNumericalIntegrationStep As Single
    PTCDoNotNormalizeSpecifiedFlag As Boolean           ' default false
    
    PeakingBeforeAcquisitionElementFlags(1 To MAXCHAN%) As Integer
    
    BlankCorrectionUnks(1 To MAXCHAN%) As Integer   ' blank correction unknown row number
    BlankCorrectionLevels(1 To MAXCHAN%) As Single  ' blank correction level (usually zero)

    ImageShiftX As Single    ' change from integer for SX100/SXFive (10-29-2011)
    ImageShiftY As Single    ' change from integer for SX100/SXFive (10-29-2011)
    
    WDSWaveScanHiPeaks(1 To MAXCHAN%) As Single
    WDSWaveScanLoPeaks(1 To MAXCHAN%) As Single
    WDSWaveScanPoints(1 To MAXCHAN%) As Integer
    
    WDSQuickScanHiPeaks(1 To MAXCHAN%) As Single
    WDSQuickScanLoPeaks(1 To MAXCHAN%) As Single
    WDSQuickScanSpeeds(1 To MAXCHAN%) As Single
    
    NthPointAcquisitionFlag As Boolean
    NthPointAcquisitionFlags(1 To MAXCHAN%) As Integer
    NthPointAcquisitionIntervals(1 To MAXCHAN%) As Integer
    NthPointMonitorFlag As Boolean
    NthPointMonitorElement As String
    NthPointPercentChange As Single
    
    MultiPointNumberofPointsAcquireHi(1 To MAXCHAN%) As Integer     ' number of points to acquire on high side of peak
    MultiPointNumberofPointsAcquireLo(1 To MAXCHAN%) As Integer     ' number of points to acquire on low side of peak
    MultiPointNumberofPointsIterateHi(1 To MAXCHAN%) As Integer     ' number of points to iterate to on high side of peak
    MultiPointNumberofPointsIterateLo(1 To MAXCHAN%) As Integer     ' number of points to iterate to on low side of peak
    MultiPointBackgroundFitType(1 To MAXCHAN%) As Integer           ' 0 = linear, 1 = 2nd order polynomial
    
    MultiPointAcquirePositionsHi() As Single        ' allocated in InitSample: 1 To MAXCHAN%, 1 To MAXMULTI%
    MultiPointAcquirePositionsLo() As Single
    MultiPointAcquireLastCountTimesHi() As Single   ' allocated in InitSample: 1 To MAXCHAN%, 1 To MAXMULTI%
    MultiPointAcquireLastCountTimesLo() As Single

    ' New Multi table (for acquisition data) (allocated in InitSample: 1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%)
    MultiPointAcquireCountTimesHi() As Single  ' default to off-peak time divided by MultiPointNumberofPointsIterateHi
    MultiPointAcquireCountTimesLo() As Single  ' default to off-peak time divided by MultiPointNumberofPointsIterateLo
    
    MultiPointAcquireCountsHi() As Single  ' high peak side multi-point background x-ray intensities
    MultiPointAcquireCountsLo() As Single  ' low peak side multi-point background x-ray intensities
    
    MultiPointProcessManualFlagHi() As Integer  ' manual override flag (-1 = never use, 0 = automatic, 1 = always use)
    MultiPointProcessManualFlagLo() As Integer  ' manual override flag (-1 = never use, 0 = automatic, 1 = always use)
    
    MultiPointProcessLastManualFlagHi() As Integer  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)     ' v. 11.3.9
    MultiPointProcessLastManualFlagLo() As Integer  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)     ' v. 11.3.9
    
    SpecifyMatrixByAnalysisUnknownNumber As Integer     ' unknown sample number to be used for specifying matrix elements
    UnknownCountTimeForInterferenceStandardFlag As Boolean
    UnknownCountTimeForInterferenceStandardChanFlag(1 To MAXCHAN%) As Boolean
    
    SampleDensity As Single   ' sample density in gm/cm3
    
    SecondaryFluorescenceBoundaryFlag(1 To MAXCHAN%) As Integer     ' 0 = do not perform correction, 1 = perform correction (stored)
    
    SecondaryFluorescenceBoundaryKratiosDATFile() As String          ' allocated in InitSample (1 To MAXCHAN%) (stored) (for PAR couple)
    SecondaryFluorescenceBoundaryKratiosDATFileLine1() As String     ' allocated in InitSample (1 To MAXCHAN%) (stored)
    SecondaryFluorescenceBoundaryKratiosDATFileLine2() As String     ' allocated in InitSample (1 To MAXCHAN%) (stored)
    SecondaryFluorescenceBoundaryKratiosDATFileLine3() As String     ' allocated in InitSample (1 To MAXCHAN%) (stored)
    
    SecondaryFluorescenceBoundaryDistanceMethod As Integer      ' 0 = specified distance, 1 = calculated from boundary (stored)
    SecondaryFluorescenceBoundarySpecifiedDistance As Single    ' in microns (stored)
    SecondaryFluorescenceBoundaryCoordinateX1 As Single         ' in stage coordinates (stored)
    SecondaryFluorescenceBoundaryCoordinateY1 As Single         ' in stage coordinates (stored)
    SecondaryFluorescenceBoundaryCoordinateX2 As Single         ' in stage coordinates (stored)
    SecondaryFluorescenceBoundaryCoordinateY2 As Single         ' in stage coordinates (stored)
    
    SecondaryFluorescenceBoundaryImageNumber As Integer     ' image number for BIM file (stored)
    SecondaryFluorescenceBoundaryImageFileName As String    ' original image file name (stored)
       
    SecondaryFluorescenceBoundaryDistance() As Single                       ' (calculated in um) allocated in InitSample (1 To MAXROW%) (calculated)
    SecondaryFluorescenceBoundaryKratios() As Single                        ' allocated in InitSample (1 To MAXROW%, 1 To MAXCHAN%) (calculated)
    
    SecondaryFluorescenceBoundaryMatA_String(1 To MAXCHAN%) As String       ' for stored kratio DAT table
    SecondaryFluorescenceBoundaryMatB_String(1 To MAXCHAN%) As String       ' for stored kratio DAT table
    SecondaryFluorescenceBoundaryMatBStd_String(1 To MAXCHAN%) As String    ' for stored kratio DAT table
        
    OnPeakTimeFractionFlag As Boolean
    OnPeakTimeFractionValue As Single
    
    ChemicalAgeCalculationFlag As Boolean
      
    ConditionNumbers(1 To MAXCHAN%) As Integer     ' condition number for this element (stored in "AcqusitionOrders" table)
    ConditionOrders(1 To MAXCOND%) As Integer      ' sample condition acquisition order (stored in "AcqusitionOrders" table)
    
    ' CL spectrum intensities
    CLSpectraFlag As Boolean                          ' CL spectrum data is stored in CL Spectra table
    CLSpectraIntensities() As Long                    ' allocated in InitSample (1 to MAXROW%, 1 to MAXSPECTRA_CL%)
    CLSpectraDarkIntensities() As Long                ' allocated in InitSample (1 to MAXROW%, 1 to MAXSPECTRA_CL%)
    CLSpectraNanometers() As Single                   ' allocated in InitSample (1 to MAXROW%, 1 to MAXSPECTRA_CL%)
    
    CLSpectraNumberofChannels() As Integer            ' allocated in InitSample (1 to MAXROW%)
    CLAcquisitionCountTime() As Single                ' allocated in InitSample (1 to MAXROW%)
    CLSpectraStartEnergy() As Single                  ' allocated in InitSample (1 to MAXROW%)
    CLSpectraEndEnergy() As Single                    ' allocated in InitSample (1 to MAXROW%)
    CLSpectraKilovolts() As Single                    ' allocated in InitSample (1 to MAXROW%)
    CLDarkSpectraCountTimeFraction() As Single        ' allocated in InitSample (1 to MAXROW%)
    
    CLUnknownCountFactors() As Single                 ' allocated in InitSample (1 to MAXROW%)
    
    LastCLSpecifiedCountTime As Single
    LastCLUnknownCountFactor As Single
    LastCLDarkSpectraCountTimeFraction As Single
    
    MaterialType As String                            ' for standard database only
    
    FerrousFerricCalculationFlag As Boolean           ' flag to calculate ferrous/ferric ratio for excess oxygen in matrix corrections
    FerrousFerricTotalCations As Single               ' total number of cations in the ferrous/ferric mineral
    FerrousFerricTotalOxygens As Single               ' total number of oxygens in the ferrous/ferric mineral
    
    FerrousFerricOption As Integer                    ' new Droop option for amphiboles (Moy)
    
    MountNames As String                              ' for standard database only (comma delimited string containing standard mounts with this standard)
    
    UnknownIsStandardNumber As Integer                ' assume unknown is a standard, 0 = not a standard, non-zero = standard number
    
    EffectiveTakeOffs(1 To MAXCHAN%) As Single        ' efective take off angle for each spectrometer/crystal pair (from SCALERS.DAT)
End Type

Type TypeImage
    ImageSampleNumber As Integer    ' sample number row reference
    ImageNumber As Integer          ' image number (MAXIMAGES%)
    ImageChannelName As String
    ImageAnalogAverages As Integer
    ImageIx As Integer              ' x pixels
    ImageIy As Integer              ' y pixels
    ImageXmin As Single
    ImageXmax As Single
    ImageYmin As Single
    ImageYmax As Single
    ImageZmin As Long
    ImageZmax As Long
    ImageMag As Single
    ImageZ1 As Single         ' z axis stage positions for each corner (xmin, ymin) (default)
    ImageZ2 As Single         '  (xmin, ymax)
    ImageZ3 As Single         '  (xmax, ymax)
    ImageZ4 As Single         '  (xmax, ymin)
    
    ImageTakeoff As Single
    ImageKilovolts As Single
    ImageBeamCurrent As Single
    ImageBeamSize As Single
    ImageTitle As String
    ImageScanRotation As Single     ' added 12/12/2014
    ImageChannelNumber As Integer   ' added 11/13/2015
    ImageDisplayDPI As Single       ' added 12/03/2017
    
    ImageData() As Long             ' image intensity data (unnormalized)
End Type

Type TypeScan
    ScanToRow As Integer    ' points back to Sample table/RowOrder field
    ScanNumber As Long      ' not quite unique scan number
    ScanChannel As Integer  ' channel number (scannumber & scanchannel are unique)
    ScanType As Integer     ' 1 = PHA, 2 = Bias, 3 = Gain, 4 = peaking
    ScanMotor As Integer    ' spectrometer
    ScanElsyms As String
    ScanXrsyms As String
    ScanCrystal As String
    ScanCountTime As Single
    ScanFitCentroid As Single
    ScanFitThreshold As Single
    ScanFitPtoB As Single
    ScanFitCoeff1 As Single
    ScanFitCoeff2 As Single
    ScanFitCoeff3 As Single
    ScanFitDeviation As Single
    ScanPoints As Integer           ' calculated at read time
    ScanXdata() As Single           ' PHA or spec position
    ScanYdata() As Single           ' intensity
    
    ScanCurrentTakeOff As Single
    ScanCurrentKilovolts As Single
    ScanCurrentBeamCurrent As Single
    ScanCurrentBeamSize As Single
    
    ScanCurrentColumnConditionMethod As Integer     ' 0 = TKCS, 1 = condition string
    ScanCurrentColumnConditionString As String
    
    ScanCurrentMagnification As Single
    ScanCurrentBeamMode As Integer                  ' 0 = spot, 1 = scan, 2 = digital

    ScanCurrentBaseline As Single
    ScanCurrentWindow As Single
    ScanCurrentGain As Single
    ScanCurrentBias As Single
    ScanCurrentInteDiffMode As Integer
    ScanCurrentDeadTime As Single
    
    ScanCurrentPeakingStartSize As Single
    ScanCurrentPeakingStopSize As Single
    ScanCurrentPositionSample As String
    ScanCurrentStageX As Single
    ScanCurrentStageY As Single
    ScanCurrentStageZ As Single
    
    ScanROMPeakingType As Integer   ' 0 = internal, 1 = parabolic, 2 = maxima, 3 = gaussian, 4 = smart, 5 = smart, 6 = highest
    ScanDateTime As Variant         ' scan date and time
    
    ScanROMPeakingSet As Integer    ' 0 = final, 1 = coarse
    
    ScanPHAHardwareType As Integer  ' 0 = traditional PHA, 1 = MCA PHA
    
    ScanCurPeakPos As Single        ' v. 13.1.1
End Type

Type TypeMultiPoint
    MultiPointHiCountTimes(1 To MAXMULTI%) As Single   ' user specified off-peak count times
    MultiPointLoCountTimes(1 To MAXMULTI%) As Single

    MultiPointHiCounts(1 To MAXMULTI%) As Single       ' acquisition intensity data
    MultiPointLoCounts(1 To MAXMULTI%) As Single

    MultiPointHiManualFlag(1 To MAXMULTI%) As Integer  ' manual override flags for fitting (-1 = never use, 0 = automatic, 1 = always use)
    MultiPointLoManualFlag(1 To MAXMULTI%) As Integer
End Type

' Character constants
Global VbSpace As String * 1            ' space character
Global VbDquote As String * 1           ' double quote character
Global VbSquote As String * 1           ' single quote character
Global VbComma As String * 1            ' comma character
Global VbForwardSlash As String * 1     ' forward slash character
Global VbColon As String * 1            ' colon character

Global Const MAC_FILE_RECORD_LENGTH% = 2400
Global Const XRAY_FILE_RECORD_LENGTH% = 188

' Acquisition times for CalcImage pixels on a channel by channel basis
Global ConditionSampleDateTime(1 To MAXCHAN%) As Variant

' Real-Time parameters (stored in "Current" table .MDB file)
Global CurrentTableUpdatedOnly As Integer   ' flag for warning when loading saved sample setup
Global CurrentOnPeaks(1 To MAXCHAN%) As Single
Global CurrentHiPeaks(1 To MAXCHAN%) As Single
Global CurrentLoPeaks(1 To MAXCHAN%) As Single

Global CurrentWaveScanHiPeaks(1 To MAXCHAN%) As Single
Global CurrentWaveScanLoPeaks(1 To MAXCHAN%) As Single
Global CurrentWaveScanPoints(1 To MAXCHAN%) As Integer

Global CurrentPeakScanHiPeaks(1 To MAXCHAN%) As Single
Global CurrentPeakScanLoPeaks(1 To MAXCHAN%) As Single
Global CurrentPeakScanPoints(1 To MAXCHAN%) As Integer

Global CurrentQuickScanHiPeaks(1 To MAXCHAN%) As Single     ' new 10/6/07 ("quickscanpoints" points not used)
Global CurrentQuickScanLoPeaks(1 To MAXCHAN%) As Single
Global CurrentQuickScanSpeeds(1 To MAXCHAN%) As Single      ' new 05/23/2009

Global CurrentPeakingStartSizes(1 To MAXCHAN%) As Single
Global CurrentPeakingStopSizes(1 To MAXCHAN%) As Single
Global CurrentMinimumPeakToBackgrounds(1 To MAXCHAN%) As Single
Global CurrentMinimumPeakCounts(1 To MAXCHAN%) As Single
Global CurrentMaximumPeakAttempts(1 To MAXCHAN%) As Integer

Global CurrentPeakCountTimes(1 To MAXCHAN%) As Single
Global CurrentWaveCountTimes(1 To MAXCHAN%) As Single
Global CurrentQuickCountTimes(1 To MAXCHAN%) As Single

Global CurrentOnCountTimes(1 To MAXCHAN%) As Single
Global CurrentHiCountTimes(1 To MAXCHAN%) As Single
Global CurrentLoCountTimes(1 To MAXCHAN%) As Single
Global CurrentUnknownCountFactors(1 To MAXCHAN%) As Single  ' changed to single (4/24/02)
Global CurrentUnknownMaxCounts(1 To MAXCHAN%) As Long

Global CurrentOrderNumbers(1 To MAXCHAN%) As Integer
Global CurrentStdBackgroundTypes(1 To MAXCHAN%) As Integer
Global CurrentUnkBackgroundTypes(1 To MAXCHAN%) As Integer

Global CurrentBaselines(1 To MAXCHAN%) As Single
Global CurrentWindows(1 To MAXCHAN%) As Single
Global CurrentGains(1 To MAXCHAN%) As Single
Global CurrentBiases(1 To MAXCHAN%) As Single
Global CurrentInteDiffModes(1 To MAXCHAN%) As Integer
Global CurrentDeadTimes(1 To MAXCHAN%) As Single

Global CurrentSlitSizes(1 To MAXCHAN%) As String
Global CurrentSlitPositions(1 To MAXCHAN%) As String
Global CurrentDetectorModes(1 To MAXCHAN%) As String

Global CurrentIntegratedIntensitiesUseIntegratedFlags(1 To MAXCHAN%) As Integer
Global CurrentIntegratedIntensitiesInitialStepSizes(1 To MAXCHAN%) As Single
Global CurrentIntegratedIntensitiesMinimumStepSizes(1 To MAXCHAN%) As Single

Global CurrentPeakingBeforeAcquisitionElementFlags(1 To MAXCHAN%) As Integer

Global CurrentNthPointAcquisitionFlags(1 To MAXCHAN%) As Integer
Global CurrentNthPointAcquisitionIntervals(1 To MAXCHAN%) As Integer

Global CurrentMultiPointNumberofPointsAcquireHi(1 To MAXCHAN%) As Integer
Global CurrentMultiPointNumberofPointsAcquireLo(1 To MAXCHAN%) As Integer
Global CurrentMultiPointNumberofPointsIterateHi(1 To MAXCHAN%) As Integer
Global CurrentMultiPointNumberofPointsIterateLo(1 To MAXCHAN%) As Integer
Global CurrentMultiPointBackgroundFitType(1 To MAXCHAN%) As Integer
    
Global CurrentMultiPointAcquirePositionsHi() As Single        ' allocated in InitData: 1 To MAXCHAN%, 1 To MAXMULTI%
Global CurrentMultiPointAcquirePositionsLo() As Single        ' allocated in InitData: 1 To MAXCHAN%, 1 To MAXMULTI%
Global CurrentMultiPointAcquireLastCountTimesHi() As Single   ' allocated in InitData: 1 To MAXCHAN%, 1 To MAXMULTI%
Global CurrentMultiPointAcquireLastCountTimesLo() As Single   ' allocated in InitData: 1 To MAXCHAN%, 1 To MAXMULTI%

' File unit numbers used in Probe for EPMA
Global Const ProbeImageFileNumber% = 100   ' ProbeImageFile$ (*.BIM)
Global Const TipOfTheDayFileNumber% = 101  ' TipOfTheDayFile$ (TipOfTheDay.TXT)
Global Const Position1FileNumber% = 102    ' *.POS position import or export
Global Const Position2FileNumber% = 103    ' *.POS position import or export

Global Const MACFileNumber% = 106          ' MACFile$ (MAC binary data file)

Global Const XEdgeFileNumber% = 107        ' XedgeFile$ (xray edge energy binary file)
Global Const XLineFileNumber% = 108        ' XlineFile$ (xray line energy binary file)
Global Const XFlurFileNumber% = 109        ' XflurFile$ (xray fluorescence yield binary file)

Global Const EMPFileNumber% = 110          ' EmpMACFile$ or EmpAPFFile$ (empirical MAC/APF ASCII files)
Global Const EMPFacFileNumber% = 111       ' EmpFacFile$ (empirical Alpha Factors ASCII file)
Global Const EMPPHAFileNumber% = 112       ' EmpPHAFile$ (empirical PHA coefficient ASCII file)

Global Const XLineFileNumber2% = 113       ' XlineFile2$ (xray line energy binary file for additional x-rays)
Global Const XFlurFileNumber2% = 114       ' XflurFile2$ (xray fluorescence yield binary file for additional x-rays)

Global Const AbsorbFileNumber% = 115       ' AbsorbFile$ (coefficient data for ABSORB.BAS)

Global Const AFactorDataFileNumber% = 117           ' AFactorDataFile$ (alpha-factor ASCII file)
Global Const ProbeErrorLogFileNumber% = 118         ' ProbeErrorLogFile$ (PROBEWIN.ERR)
Global Const OutputReportFileNumber% = 119          ' OutputReportFile$ (*.TXT) analyze report output
Global Const RecalbELMFileNumber% = 120             ' .ELM file number
Global Const RecalbPHAFileNumber% = 121             ' .PHA file number
Global Const EMSASpectrumFileNumber% = 122          ' EMSA EDS file number
Global Const ProbeTextLogFileNumber% = 123          ' ProbeTextLogFile$ (PROBEWIN.TXT)
Global Const OutputDataFileNumber% = 124            ' OutputDataFile$ (*.OUT)

Global Const Temp1FileNumber% = 125                 ' temporary file I/O
Global Const Temp2FileNumber% = 126                 ' temporary file I/O
Global Const CustomOutputFileNumber% = 200          ' #200-255 Custom format analysis output files

' Database access flags
Global DatabaseExclusiveAccess As Integer        ' for generic exclusive access
Global StandardDatabaseExclusiveAccess As Integer
Global ProbeDatabaseExclusiveAccess As Integer
Global SetupDatabaseExclusiveAccess As Integer
Global UserDatabaseExclusiveAccess As Integer
Global PositionDatabaseExclusiveAccess As Integer
Global XrayDatabaseExclusiveAccess As Integer
Global MatrixDatabaseExclusiveAccess As Integer
Global BoundaryDatabaseExclusiveAccess As Integer
Global PureDatabaseExclusiveAccess As Integer

Global DatabaseNonExclusiveAccess As Integer        ' for generic non-exclusive access
Global StandardDatabaseNonExclusiveAccess As Integer
Global ProbeDatabaseNonExclusiveAccess As Integer
Global SetupDatabaseNonExclusiveAccess As Integer
Global UserDatabaseNonExclusiveAccess As Integer
Global PositionDatabaseNonExclusiveAccess As Integer
Global XrayDatabaseNonExclusiveAccess As Integer
Global MatrixDatabaseNonExclusiveAccess As Integer
Global BoundaryDatabaseNonExclusiveAccess As Integer
Global PureDatabaseNonExclusiveAccess As Integer

Global SystemPath As String
Global ProgramPath As String                ' works under VB6 IDE

Global ApplicationPath As String            ' does not work under VB6 IDE
Global ApplicationCommonAppData As String   ' all users
Global ApplicationAppData As String         ' roaming users

' Xray data files
Global XLineFile As String
Global XFlurFile As String
Global XLineFile2 As String     ' for additional x-ray lines
Global XFlurFile2 As String     ' for additional x-ray lines

Global XEdgeFile As String
Global ElementsFile As String
Global AbsorbFile As String

Global EmpMACFile As String
Global EmpAPFFile As String
Global EmpFACFile As String
Global EmpPHAFile As String

Global MACFile As String

Global ProbeWinINIFile As String
Global WindowINIFile As String
Global CrystalsFile As String

Global MaxMenuFileArray As Integer

' ASCII data files
Global OutputDataFile As String
Global OutputReportFile As String
Global AFactorDataFile As String
Global ProbeErrorLogFile As String
Global ProbeTextLogFile As String

' Access .MDB database files
Global ProbeDataFile As String
Global ProbeImageFile As String   ' image data (flat binary file)
Global OldProbeDataFile As String

Global StandardDataFile As String
Global CurrentSetupDataFile As String
Global SetupDataFile As String      ' standard intensity database
Global SetupDataFile2 As String     ' MAN standard intensity database
Global SetupDataFile3 As String     ' interference standard intensity database
Global XrayDataFile As String
Global PositionDataFile As String
Global UserDataFile As String
Global ProbeElmFile As String
Global ProbePHAFile As String

' Standard Database Index arrays
Global NumberOfAvailableStandards As Integer
Global StandardIndexNumbers(1 To MAXINDEX%) As Integer
Global StandardIndexNames(1 To MAXINDEX%) As String
Global StandardIndexDescriptions(1 To MAXINDEX%) As String
Global StandardIndexDensities(1 To MAXINDEX%) As Single
Global StandardIndexMaterialTypes(1 To MAXINDEX%) As String
Global StandardIndexMountNames(1 To MAXINDEX%) As String        ' comma delimited string containing standard mounts with this standard

' Global variables
Global FileViewer As String
Global DataFileVersionNumber As Single
Global ProbeDataFileVersionNumber As Single
Global ProgramVersionNumber As Single
Global ProgramVersionString As String

Global MDBUserName As String
Global MDBFileTitle As String
Global MDBFileDescription As String

Global MDBFileType As String
Global MDBFileCreated As String
Global MDBFileUpdated As String
Global MDBFileModified As String

Global CustomLabel1 As String
Global CustomLabel2 As String
Global CustomLabel3 As String

Global CustomText1 As String
Global CustomText2 As String
Global CustomText3 As String

Global UserStartDateTime As Variant
Global UserStopDateTime As Variant

' Real time globals
Global FaradayCupType As Integer
Global AbsorbedCurrentPresent As Integer
Global AbsorbedCurrentType As Integer

Global BeamOnFlag As Boolean
Global MoveStageWithoutBeamBlank As Integer

' PHA interface
Global PHAHardware As Integer
Global PHAHardwareType As Integer

Global PHAGainBias As Integer
Global PHAGainBiasType As Integer

Global PHAInteDiff As Integer
Global PHAInteDiffType As Integer

Global PHADeadTime As Integer
Global PHADeadTimeType As Integer

Global MinPHABaselineWindow As Single
Global MaxPHABaselineWindow As Single
Global MinPHAGainWindow As Single
Global MaxPHAGainWindow As Single
Global MaxPHABiasWindow As Single
Global MinScalerCountTime As Single
Global MaxScalerCountTime As Single

Global InterfaceType As Integer ' 0=Demo, 1=Unused, 2=JEOL 8900/8200/8500/8x30, 3=Unused, 4=Unused, 5=SX100/SXFive
Global RealTimeMode As Integer
Global LogWindowInterval As Single
Global RealTimeInterval As Single
Global RealTimeInterfaceBusy As Integer
Global EnterPositionsRelativeFlag As Integer
Global UpdatePeakWaveScanPositionsFlag As Integer

Global PositionImportExportFileType As Integer
Global DeadTimeCorrectionType As Integer
Global AutoFocusStyle As Integer
Global AutoFocusInterval As Integer
Global BiasChangeDelay As Single, DefaultBiasChangeDelay As Single
Global UseEmpiricalPHADefaults As Integer
Global KilovoltChangeDelay As Single
Global BeamCurrentChangeDelay As Single
Global BeamSizeChangeDelay As Single

Global FilamentStandbyPresent As Integer
Global FilamentStandbyType As Integer               ' 0 = reduce heat only, 1 = reduce heat and keV, 2 = reduce keV only, 3 = external script, 4 = load PCC file
Global FilamentStandbyExternalScript As String

Global OperatingVoltagePresent As Integer
Global OperatingVoltageType As Integer

Global BeamCurrentPresent As Integer
Global BeamCurrentType As Integer

Global BeamSizePresent As Integer
Global BeamSizeType As Integer

Global AutoFocusPresent As Integer
Global AutoFocusType As Integer             ' 0 = parabolic, 1 = gaussian, 2 = maximum value

Global ROMPeakingPresent As Integer
Global ROMPeakingString(0 To MAXROMPEAKTYPES%) As String
Global DefaultROMPeakingType As Integer     ' 0=internal, 1=parabolic, 2=maxima, 3=gaussian, 4 = smart, 5 = smart, 6 = highest intensity

Global EDSSpectraInterfacePresent As Integer            ' EDS spectrum interface
Global EDSSpectraInterfaceType As Integer               ' EDS spectrum interface type, 0 = Demo, 1 = JEOL MEC, 2 = Bruker, 3 = Oxford, 4 = Unused, 5 = Thermo NSS, 6 = JEOL

' Faraday cup parameters
Global FaradayWaitInTime As Single
Global FaradayWaitOutTime As Single

' Automation globals
Global AcquisitionOnMotorCrystal As Integer
Global AcquisitionOnCounterCount As Integer

Global AcquisitionOnSample As Integer
Global AcquisitionOnWavescan As Integer
Global AcquisitionOnPeakCenter As Integer
Global AcquisitionOnAutomate As Integer
Global tAcquisitionOnAutomate As Integer    ' used for quick standard time calculation

Global AcquisitionOnVolatile As Integer
Global AcquisitionOnQuickscan As Integer
Global AcquisitionOnImageInterface As Integer
Global AcquisitionOnBeamDeflect As Integer
Global AcquisitionOnAutoFocus As Integer
Global AcquisitionOnROMPeakPHA As Integer
Global AcquisitionOnConditions As Integer
Global AcquisitionOnROMScan(1 To MAXSPEC%) As Integer

Global AcquisitionOnEDS As Integer
Global AcquisitionOnCL As Integer

Global AutomateStep As Integer
Global AutomatePositionStep As Integer
Global AutomateListRowNumber As Integer
Global AutomateGridRowNumber As Integer
Global AutomateAutoFocusNumber As Integer

' Globals for FormVOLATILE
Global VolatileSampleName As String
Global VolatileCountIntervals As Integer
Global VolatileXIncrement As Integer
Global VolatileYIncrement As Integer
Global VolatileSelfCalibrationAcquisitionFlag As Integer
Global VolatileAssignedCalibrationAcquisitionFlag As Integer

' Globals for FormQUICK
Global QuickscanSampleName As String
Global QuickscanSpeed As Single

' Globals for FormDIGITIZE
Global NumberofUnknownPositionSamples As Integer
Global NumberofWavescanPositionSamples As Integer

Global NoMotorPositionBoundsChecking(1 To MAXMOT%) As Integer
Global NoMotorPositionLimitsChecking(1 To MAXMOT%) As Integer

' Updated values from Timer event
Global RealTimeMotorPositions(1 To MAXMOT%) As Single
Global RealTimeCrystalPositions(1 To MAXSPEC%) As String

Global RealTimeScalLabels(1 To MAXSPEC%) As String

' Wave/Peak scan globals
Global WaveMode As Integer  ' 1 = wavescan, 2 = peakscan
Global WavePeakCenterStart As Single
Global WavePeakCenterMotor As Integer
Global WavePeakCenterChannel As Integer
Global WavePeakCenterFlags(1 To MAXCHAN%) As Boolean
Global WavePeakSuccessFlags(1 To MAXCHAN%) As Boolean

Global WavescanXIncrementFlag As Integer
Global PeakingXIncrementFlag As Integer
Global UnknownXIncrementFlag As Integer

Global WavescanXIncrement As Single
Global WavescanXIncrementInterval As Single
Global WavescanYIncrement As Single
Global WavescanXIncrementPosition As Single
Global WavescanYIncrementPosition As Single

' Acquisition option flags
Global AutomateConfirmFlag As Integer
Global AutomateConfirmDelay As Single
Global AutomateNewSampleBasisFlag As Integer
Global AnalysisInProgress As Integer
Global AnalysisIsRunning As Integer
Global AllAnalysisUpdateNeeded As Integer
Global AllAFactorUpdateNeeded As Integer

Global SyncSpecMotionBeamBlankFlag As Integer   ' not implemented
Global StdUnkMeasureFaradayFlag As Integer
Global WaveScanMeasureFaradayFlag As Integer
Global AbsorbedCurrentMeasureFlag As Integer
Global LoadStdDataFromFileSetupFlag As Integer

Global AcquireEDSSpectraFlag As Integer

Global AutoAnalyzeFlag As Integer
Global UseAutomatedPHAControlFlag As Integer

Global DefaultBlankBeamFlag As Integer
Global ReturnToOnPeakFlag As Integer

Global QuickWaveScanAcquisitionFlag As Integer

' FormAUTOMATE globals
Global PeakOnAssignedStandardsFlag As Integer
Global UseQuickStandardsFlag As Integer
Global UseFilamentStandbyFlag As Integer
Global UseAutoFocusFlag As Integer

Global StandardPointsToAcquire As Integer
Global IncrementXForAdditionalPoints  As Integer
Global IncrementYForReStandardizations  As Integer

Global DefaultFiducialSetNumber As Integer

Global DefaultSampleSetupNumber As Integer
Global DefaultFileSetupName As String
Global DefaultFileSetupNumber As Integer
Global DefaultMultipleSetupNumber As Integer
Global DefaultMultipleSetupNumbers() As Integer
Global DefaultMultipleSetupNumberIndex As Integer

Global UpdatePositionSamplesAutomateList As Integer
Global UpdatePositionSamplesPositionList As Integer

' Flags set in GETTIM/GETOPT/PEAK
Global NominalBeam As Single
Global OriginalNominalBeam As Single
Global DefaultBeamAverages As Integer
Global AcquisitionOrderFlag As Integer
Global AcquisitionMotionFlag As Integer

Global SpecBackLashFlag As Integer
Global StageBacklashFlag As Integer
Global SpecBackLashType As Integer
Global StageBacklashType As Integer

Global StageStdBacklashFlag As Integer
Global StageUnkBacklashFlag As Integer
Global StageWavBacklashFlag As Integer

Global AutomatePeakCenterPreScanFlag As Integer
Global AutomatePeakCenterPostScanFlag As Integer
Global PeakCenterPreScanFlag As Integer
Global PeakCenterPostScanFlag As Integer
Global PeakCenterMethodFlag As Integer
Global PeakSkipPBCheck As Integer

' Microprobe configuration globals from PROBEWIN.INI
Global DefaultKiloVolts As Single
Global DefaultTakeOff As Single
Global DefaultBeamCurrent As Single
Global DefaultBeamSize As Single

Global DefaultDebugMode As Integer
Global DefaultOxideOrElemental As Integer

Global DefaultOnCountTime As Single
Global DefaultOffCountTime As Single
Global DefaultUnknownMaxCounts As Long
Global DefaultPeakingCountTime As Single
Global DefaultWavescanCountTime As Single
Global DefaultQuickscanCountTime As Single

Global DefaultPHACountTime As Single
Global DefaultPHAIntervals As Integer

Global DefaultMinimumKLMDisplay As Single
Global DefaultAbsorptionEdgeDisplay As Integer
Global DefaultGraphType As Integer
Global DefaultGraphTypeWav As Integer

Global DefaultXrayStart As Single
Global DefaultXrayStop As Single
Global DefaultRangeFraction As Single
Global DefaultLIFPeakWidth As Single
Global DefaultMinimumOverlap As Single
Global DefaultPHADiscrimination As Single

Global DefaultPeakCenterMethod As Integer   ' 0 = interval halving, 1 = parabolic, 2 = ROM, 3 = manual

Global NumberOfTunableSpecs As Integer
Global NumberOfStageMotors As Integer

' CRYSTALS.DAT globals
Global AllCrystalNames(1 To MAXCRYSTYPE%) As String
Global AllCrystal2ds(1 To MAXCRYSTYPE%) As Single
Global AllCrystalKs(1 To MAXCRYSTYPE%) As Single
Global AllCrystalElements(1 To MAXCRYSTYPE%) As String
Global AllCrystalXrays(1 To MAXCRYSTYPE%) As String

' SCALERS.DAT globals
Global ScalLabels(1 To MAXSPEC%) As String  ' must be numeric
Global ScalCrystalFlipFlags(1 To MAXSPEC%) As Integer
Global ScalCrystalFlipPositions(1 To MAXSPEC%) As Single
Global ScalNumberOfCrystals(1 To MAXSPEC%) As Integer
Global ScalCrystalNames(1 To MAXCRYS%, 1 To MAXSPEC%) As String
Global ScalOffPeakFactors(1 To MAXSPEC%) As Single

Global ScalWaveScanSizeFactors(1 To MAXSPEC%) As Single
Global ScalPeakScanSizeFactors(1 To MAXSPEC%) As Single
Global ScalWaveScanPoints(1 To MAXSPEC%) As Integer
Global ScalPeakScanPoints(1 To MAXSPEC%) As Integer

Global ScalLiFPeakingStartSizes(1 To MAXSPEC%) As Single
Global ScalLiFPeakingStopSizes(1 To MAXSPEC%) As Single
Global ScalMaximumPeakAttempts(1 To MAXSPEC%) As Integer
Global ScalMinimumPeakToBackgrounds(1 To MAXSPEC%) As Single
Global ScalMinimumPeakCounts(1 To MAXSPEC%) As Single

Global ScalBaseLines(1 To MAXCRYS%, 1 To MAXSPEC%) As Single        ' new dimensions for crystals
Global ScalWindows(1 To MAXCRYS%, 1 To MAXSPEC%) As Single
Global ScalGains(1 To MAXCRYS%, 1 To MAXSPEC%) As Single
Global ScalBiases(1 To MAXCRYS%, 1 To MAXSPEC%) As Single
Global ScalInteDiffModes(1 To MAXCRYS%, 1 To MAXSPEC%) As Integer   ' new dimensions for intediff and deadtime
Global ScalDeadTimes(1 To MAXCRYS%, 1 To MAXSPEC%) As Single

Global ScalInteDeadTimes(1 To MAXSPEC%) As Integer              ' Cameca integer hardware deadtimes only
Global ScalLargeArea(1 To MAXCRYS%, 1 To MAXSPEC%) As Integer   ' Cameca large area crystal flags only

Global ScalEffectiveTakeOffs(1 To MAXCRYS%, 1 To MAXSPEC%) As Single   ' new parameter for spectrometer/crystals

Global ScalBaseLineScaleFactors(1 To MAXSPEC%) As Single
Global ScalWindowScaleFactors(1 To MAXSPEC%) As Single
Global ScalGainScaleFactors(1 To MAXSPEC%) As Single
Global ScalBiasScaleFactors(1 To MAXSPEC%) As Single

Global ScalRolandCircleMMs(1 To MAXSPEC%) As Single
Global ScalCrystalFlipDelays(1 To MAXSPEC%) As Single
Global ScalSpecOffsetFactors(1 To MAXSPEC%) As Single

Global ScalBiasScanLows(1 To MAXSPEC%) As Single     ' added 12/10/05
Global ScalBiasScanHighs(1 To MAXSPEC%) As Single
Global ScalGainScanLows(1 To MAXSPEC%) As Single     ' added 12/12/05
Global ScalGainScanHighs(1 To MAXSPEC%) As Single
Global ScalScanBaselines(1 To MAXSPEC%) As Single
Global ScalScanWindows(1 To MAXSPEC%) As Single

' MOTORS.DAT globals
Global XMotor As Integer, YMotor As Integer, ZMotor As Integer
Global MotLabels(1 To MAXMOT%) As String
Global MotLoLimits(1 To MAXMOT%) As Single
Global MotHiLimits(1 To MAXMOT%) As Single
Global MotUnitsToAngstromMicrons(1 To MAXMOT%) As Single
Global MotBacklashFactors(1 To MAXMOT%) As Single
Global MotBacklashTolerances(1 To MAXMOT%) As Single
Global MotParkPositions(1 To MAXMOT%) As Single

' Symbol variables
Global Symlo(1 To MAXELM%) As String
Global Symup(1 To MAXELM%) As String
Global Edglo(1 To MAXEDG%) As String
Global Xraylo(1 To MAXRAY%) As String

Global Deflin(1 To MAXELM%) As String
Global Defcry(1 To MAXELM%) As String

Global AllAtomicNums(1 To MAXELM%) As Integer
Global AllCat(1 To MAXELM%) As Integer
Global AllOxd(1 To MAXELM%) As Integer
Global AllAtomicWts(1 To MAXELM%) As Single
Global AllAtomicCharges(1 To MAXELM%) As Single
Global AllAtomicDensities(1 To MAXELM%) As Single

' For DENSITY2.DAT
Global AllAtomicDensities2(1 To MAXELM%) As Single  ' liquid densites (at boiling point)
Global AllAtomicDensities3(1 To MAXELM%) As Single  ' solid densities (at melting point)
Global AllAtomicVolumes(1 To MAXELM%) As Single     ' atomic volumes in cm^3/mol

Global macez(1 To MAXEMP%) As Integer
Global macxl(1 To MAXEMP%) As Integer
Global macaz(1 To MAXEMP%) As Integer
Global macval(1 To MAXEMP%) As Single
Global macstr(1 To MAXEMP%) As String

Global macrenormfactor(1 To MAXEMP%) As Single
Global macrenormstandard(1 To MAXEMP%) As String

Global apfez(1 To MAXEMP%) As Integer
Global apfxl(1 To MAXEMP%) As Integer
Global apfaz(1 To MAXEMP%) As Integer
Global apfval(1 To MAXEMP%) As Single
Global apfstr(1 To MAXEMP%) As String

Global apfrenormfactor(1 To MAXEMP%) As Single
Global apfrenormstandard(1 To MAXEMP%) As String

' Print out and I/O
Global msg As String
Global icancel As Boolean
Global icancelauto As Boolean
Global icancelanal As Boolean
Global icancelload As Boolean
Global DebugMode As Integer
Global ExtendedFormat As Integer
Global SaveToDisk As Integer
Global SaveToText As Integer

Global LogWindowFontName As String
Global LogWindowFontSize As Integer
Global LogWindowFontBold As Integer
Global LogWindowFontItalic As Integer
Global LogWindowFontUnderline As Integer
Global LogWindowFontStrikeThru As Integer

Global AcquirePositionFontSize As Integer
Global AcquireCountFontSize As Integer

' Sample arrays
Global SampleNums(1 To MAXSAMPLE%) As Integer   ' sample numbers
Global SampleTyps(1 To MAXSAMPLE%) As Integer   ' sample types (1=st, 2=un, 3=wa)
Global SampleSets(1 To MAXSAMPLE%) As Integer   ' sample sets (always 1 for unknown)
Global SampleNams(1 To MAXSAMPLE%) As String    ' sample names
Global SampleDess(1 To MAXSAMPLE%) As String    ' sample descriptions
Global SampleDels(1 To MAXSAMPLE%) As Integer   ' sample deleted flags (all lines deleted = true)
Global SampleMags(1 To MAXSAMPLE%) As Single    ' sample magnifications (analytical)

Global NumberofSamples As Integer
Global NumberofStandards As Integer
Global NumberofUnknowns As Integer
Global NumberofWavescans As Integer
Global NumberofLines As Long            ' changed from integer 11/11/05

' Variable "constants"
Global MinSpecifiedValue As Single  ' minimum amount for force loading as specified concentration
Global StdMinimumValue As Single    ' minimum amount for use as assigned standard
Global MANMaximumValue As Single    ' maximum amount for use as MAN standard background

' Standard arrays
Global StandardNumbers(1 To MAXSTD%) As Integer
Global StandardNames(1 To MAXSTD%) As String
Global StandardDescriptions(1 To MAXSTD%) As String
Global StandardDensities(1 To MAXSTD%) As Single

Global StandardCoatingFlag(1 To MAXSTD%) As Integer    ' 0 = not coated, 1 = coated
Global StandardCoatingElement(1 To MAXSTD%) As Integer
Global StandardCoatingDensity(1 To MAXSTD%) As Single
Global StandardCoatingThickness(1 To MAXSTD%) As Single ' in angstroms

' Misc flags
Global GetElmFlag As Integer      ' procedure call flag for GETELM
Global EmpTypeFlag As Integer     ' procedure call flag for EMP
Global SetupFlag As Integer       ' procedure call flag for SETUP
Global SetupSampleFlag As Integer ' procedure call flag for SETUPSAM
Global SetupRunFlag As Integer    ' procedure call flag for SETUPRUN
Global NextSample As Integer      ' flag for "Next" button

' Calculation options
Global CorrectionFlag As Integer    ' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
Global MACTypeFlag As Integer
Global EmpiricalAlphaFlag As Integer

Global UseDriftFlag As Integer      ' standard drift correction flag
Global UseInterfFlag As Integer     ' interference correction flag
Global UseVolElFlag As Integer      ' volatile element correction flag
Global UseVolElType As Integer      ' volatile element correction type (0 = linear, 1 = qradratic)
Global UseMANAbsFlag As Integer     ' MAN absorption correction flag
Global UseMACFlag As Integer        ' empirical MACs flag
Global UseAPFFlag As Integer        ' empirical APFs flag
Global UseBlankCorFlag As Integer   ' blank trace correction flag

Global UseDetailedFlag As Integer   ' extra printout flag
Global UseAutomaticFormatForResultsFlag As Integer   ' autoformatted output
Global UseAutomaticFormatForResultsType As Integer   ' autoformatted output (0=maximum decimals, 1=significant decimals)
Global PrintAnalyzedAndSpecifiedOnSameLineFlag As Integer

Global UseOffPeakElementsForMANFlag As Integer          ' see FormMAN
Global UseMANForOffPeakElementsFlag As Integer          ' see FormMAN
Global UseBeamDriftCorrectionFlag As Integer            ' see FormANALYSIS
Global UseDeadtimeCorrectionFlag As Integer             ' see FormANALYSIS
Global UseOxygenFromHalogensCorrectionFlag As Integer   ' see FormANALYSIS
Global UseChargeBalanceCalculationFlag As Integer       ' see FormANALYSIS

' ZAF selection strings
Global corstring(0 To MAXCORRECTION%) As String
Global empstring(1 To 2) As String
Global zafstring(0 To MAXZAF%) As String        ' dimension from zero for backward compatibility
Global zafstring2(0 To MAXZAF%) As String       ' dimension from zero for backward compatibility

Global mipstring(1 To 9) As String
Global bscstring(1 To 5) As String
Global phistring(1 To 7) As String
Global stpstring(1 To 6) As String
Global bksstring(0 To 10) As String
Global absstring(1 To 15) As String
Global flustring(1 To 5) As String

Global macstring(1 To MAXMACTYPE%) As String
Global macstring2(1 To MAXMACTYPE%) As String

' ZAF globals
Global iabs As Integer, ibsc As Integer, ibks As Integer
Global imip As Integer, iphi As Integer, izaf As Integer
Global istp As Integer, iflu As Integer, ielc As Integer               ' ielc% is not currently utilized

' MAN drift arrays
Global MANAssignsDriftCounts() As Single    ' (1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData
Global MANAssignsDateTimes() As Double      ' (1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData
Global MANAssignsSets() As Integer          ' (1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData
Global MANAssignsSampleRows() As Integer    ' (1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData

Global MANAssignsCountTimes() As Single     ' (1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData
Global MANAssignsBeamCurrents() As Single   ' (1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) allocated in InitData

' Stage BitMap parameters
Global StageBitMapCount As Integer
Global StageBitMapFile(1 To MAXBITMAP%) As String
Global StageBitMapXmin(1 To MAXBITMAP%) As Single
Global StageBitMapXmax(1 To MAXBITMAP%) As Single
Global StageBitMapYmin(1 To MAXBITMAP%) As Single
Global StageBitMapYmax(1 To MAXBITMAP%) As Single
Global StageBitMapIndex As Integer

Global JoyStickXPolarity As Integer
Global JoyStickYPolarity As Integer
Global JoyStickZPolarity As Integer

Global RegistrationName As String
Global RegistrationInstitution As String

Global PeakROMCentroidPosition(1 To MAXSPEC%) As Single     ' to determine generic ROM success
Global EDSThinWindowPresent As Integer
Global NoMotorPositionLimitsCheckingFlag As Integer

' New globals for 32 bit version
Global LogWindowBufferSize As Long    ' size in bytes

Global MinimumFaradayCurrent As Single
Global FaradayCurrentUnits As String

' Data folders
Global UserDataDirectory As String
Global OriginalUserDataDirectory As String
Global StandardPOSFileDirectory As String
Global CalcZAFDATFileDirectory As String
Global ColumnPCCFileDirectory As String
Global SurferDataDirectory As String
Global GrapherDataDirectory As String
Global DemoImagesDirectory As String

Global DemoImagesDirectoryJEOL As String
Global DemoImagesDirectoryCameca As String

Global CalculateElectronandXrayRangesFlag As Integer

Global ExtendedMenuFlag As Integer
Global IgnoreZAFandAlphaFactorWarnings As Integer
Global UseAlternatingOnAndOffPeakAcquisitionFlag As Integer

' Email
Global EmailNotificationOfErrorsFlag As Integer
Global SMTPServerAddress As String
Global SMTPAddressTo As String
Global SMTPAddressFrom As String
Global SMTPUserName As String
Global SMTPUserPassword As String

' Image acquisition
Global ImageInterfacePresent As Integer
Global ImageInterfaceType As Integer
Global ImageInterfaceNameChan0 As String    ' special for imported PrbImg file
Global ImageInterfaceNameChan1 As String
Global ImageInterfaceNameChan2 As String
Global ImageInterfaceNameChan3 As String
Global ImageInterfaceImageIxIy As Single    ' beam scan ratio of Ix/Iy

Global ImageInterfaceCalNumberOfBeamCalibrations As Integer                 ' for beam deflection calibration array from INI file
Global ImageInterfaceCalKeVArray(1 To MAXBEAMCALIBRATIONS%) As Single       ' for beam deflection calibration array from INI file
Global ImageInterfaceCalMagArray(1 To MAXBEAMCALIBRATIONS%) As Single       ' for beam deflection calibration array from INI file
Global ImageInterfaceCalXMicronsArray(1 To MAXBEAMCALIBRATIONS%) As Single  ' for beam deflection calibration array from INI file
Global ImageInterfaceCalYMicronsArray(1 To MAXBEAMCALIBRATIONS%) As Single  ' for beam deflection calibration array from INI file
Global ImageInterfaceCalScanRotationArray(1 To MAXBEAMCALIBRATIONS%) As Single  ' for beam deflection calibration array from INI file

Global ImageInterfaceBeamXPolarity As Integer       ' acquisition flag to fix beam polarity problems
Global ImageInterfaceBeamYPolarity As Integer       ' acquisition flag to fix beam polarity problems
Global ImageInterfaceStageXPolarity As Integer      ' acquisition flag to fix stage polarity problems
Global ImageInterfaceStageYPolarity As Integer      ' acquisition flag to fix stage polarity problems

Global ImageInterfaceDisplayXPolarity As Integer    ' acquisition flag to fix image display problems (obsolete?)
Global ImageInterfaceDisplayYPolarity As Integer    ' acquisition flag to fix image display problems (obsolete?)

Global UseBeamDeflectionFlag As Integer
Global UseEDSSampleCountTimeFlag As Integer
Global EDSSampleCountTime As Single
Global EDSSpecifiedCountTime As Single
Global EDSUnknownCountFactor As Single

' EDS (direct socket interface)
Global EDS_IPAddress As String
Global EDS_ServicePort As String

Global EDS_ServerName As String     ' Bruker EDS remote client interface
Global EDS_LoginName As String
Global EDS_LoginPassword As String

' WDS (JEOL 8900/8200/8500/8230/8530 TCP/IP direct socket interface)
Global WDS_IPAddress As String          ' used by 8200 and 8900
Global WDS_IPAddress2 As String         ' used by 8200 only for EOS
Global WDS_ServicePort As Integer       ' used by 8200 and 8900
Global WDS_ServicePort2 As Integer      ' used by 8200 only for EOS

Global MagnificationPresent As Integer
Global MagnificationType As Integer

Global OperatingVoltageTolerance As Single
Global BeamCurrentTolerance As Single
Global BeamSizeTolerance As Single

Global PeakCenterModifiedOffPeakFlag(1 To MAXCHAN%) As Integer
Global PeakCenterModifiedPeakScanFlag(1 To MAXCHAN%) As Integer
Global PeakCenterModifiedWavescanFlag(1 To MAXCHAN%) As Integer

Global ColumnConditionPresent As Integer
Global ColumnConditionType As Integer               ' interface type (not used)
Global DefaultColumnConditionMethod As Integer      ' 0 = use (TKCS), 1 = use column condition string
Global DefaultColumnConditionString As String       ' actual string

Global FaradayAlwaysOnTopFlag As Integer

Global ImageInterfacePolarityChan1 As Integer
Global ImageInterfacePolarityChan2 As Integer
Global ImageInterfacePolarityChan3 As Integer

Global ColumnConditionChangeDelay As Single

Global FaradayCurrentFormat As String
Global FaradayBeamCurrentSafeThreshold As Single

Global CombineMultipleSampleSetupsFlag As Integer
Global SurferOutputVersionNumber As Integer         ' 6 or (7 or 8 or 9, etc.)
Global SelPrintStartDocFlag As Integer

Global ProcessInterval As Single
Global RealTimePauseAutomation As Integer
Global UseZeroPointCalibrationCurveFlag As Integer
Global VerboseMode As Integer

' Plot and PlotWave parameters (and multi-point bgd plot)
Global ErrorbarSigmaIndex As Integer
Global ErrorbarSigmaNumber As Integer
Global ErrorbarSpacingIndex As Integer
Global ErrorbarSpacingNumber As Integer

Global ScanRotationPresent As Integer
Global DefaultScanRotation As Single

' For specimen mounted faraday cups
Global FaradayStagePresent As Integer
Global FaradayStagePositions(1 To MAXAXES%) As Single   ' X, Y, Z

' Detector globals
Global DetectorsFile As String
Global DetSlitSizesNumber(1 To MAXSPEC%) As Integer
Global DetSlitSizes(1 To MAXDET%, 1 To MAXSPEC%) As String
Global DetSlitPositionsNumber(1 To MAXSPEC%) As Integer
Global DetSlitPositions(1 To MAXDET%, 1 To MAXSPEC%) As String
Global DetDetectorModesNumber(1 To MAXSPEC%) As Integer
Global DetDetectorModes(1 To MAXDET%, 1 To MAXSPEC%) As String

Global DetectorSlitSizePresent As Integer
Global DetectorSlitSizeType As Integer
Global DetectorSlitPositionPresent As Integer
Global DetectorSlitPositionType As Integer
Global DetectorModePresent As Integer
Global DetectorModeType As Integer

Global RealTimeDetectorSlitSizes(1 To MAXSPEC%) As Integer          ' array pointer
Global RealTimeDetectorSlitPositions(1 To MAXSPEC%) As Integer      ' array pointer
Global RealTimeDetectorModes(1 To MAXSPEC%) As Integer              ' array pointer

Global DetSlitSizeExchangeFlags(1 To MAXSPEC%) As Integer
Global DetSlitPositionExchangeFlags(1 To MAXSPEC%) As Integer
Global DetDetectorModeExchangeFlags(1 To MAXSPEC%) As Integer

Global DetSlitSizeExchangePositions(1 To MAXSPEC%) As Single
Global DetSlitPositionExchangePositions(1 To MAXSPEC%) As Single
Global DetDetectorModeExchangePositions(1 To MAXSPEC%) As Single

Global DetSlitSizeExchangeRowlands(1 To MAXSPEC%) As Single         ' obsolete
Global DetSlitPositionExchangeRowlands(1 To MAXSPEC%) As Single     ' obsolete
Global DetDetectorModeExchangeRowlands(1 To MAXSPEC%) As Single     ' obsolete

Global DetSlitSizeDefaultIndexes(1 To MAXSPEC%) As Integer
Global DetSlitPositionDefaultIndexes(1 To MAXSPEC%) As Integer
Global DetDetectorModeDefaultIndexes(1 To MAXSPEC%) As Integer

' Multiple peak calibration global (see CalibratePeakCenterFlag and CalibratePeakCenterFiles())
Global UseMultiplePeakCalibrationOffsetFlag As Integer
Global CalibratePeakCenterFiles(0 To MAXRAY_OLD% - 1) As String

' Wav file global
Global UseWavFileAfterAutomationString As String

' Move all stage motors hardware flag
Global MoveAllStageMotorsHardwarePresent As Integer
Global Jeol8900PreAcquireString As String
Global Jeol8900PostAcquireString As String

Global ConfirmAllPositionsInSample As Integer
Global ForceAnalyticalColumnCondition As Integer

' Multiple peak coefficients (0 to 5 = Ka, Kb, La, Lb, Ma or Mb)
Global MultiplePeakCoefficient1(0 To MAXRAY% - 2, 1 To MAXCRYS%, 1 To MAXSPEC%) As Single
Global MultiplePeakCoefficient2(0 To MAXRAY% - 2, 1 To MAXCRYS%, 1 To MAXSPEC%) As Single
Global MultiplePeakCoefficient3(0 To MAXRAY% - 2, 1 To MAXCRYS%, 1 To MAXSPEC%) As Single

Global UseAPFOption As Integer
Global ExcelMethodOption As Integer     ' note different options for Probewin and CalcZAF

Global UseAggregateIntensitiesFlag As Integer
Global UseForceSizeFlag As Integer
Global UseForceColumnConditionFlag As Integer

Global CheckFaradayStateFlag As Integer     ' used only in Faraday, Monitor, GunAligh and vacuum, etc.

Global DefaultReplicates As Integer
Global ReplicatesStep As Integer

Global PlotGraphWithSectors As Integer
Global DisplayCountIntensitiesUnnormalizedFlag As Integer

Global JeolEOSInterfaceType As Long     ' 1 = 8200, 2 = 8900, 3 = 8230/8530

' New particle and thin film flags
Global iptc As Integer                  ' new particle and thin film flag
Global PTCModel As Integer              ' make into array (1 to MAXMODELS%) later
Global PTCDiameter As Single            ' make into array (1 to MAXDIAMS%) later
Global PTCDensity As Single
Global PTCThicknessFactor As Single
Global PTCNumericalIntegrationStep As Single
Global PTCDoNotNormalizeSpecifiedFlag As Boolean

Global ptcstring(1 To MAXMODELS%) As String

Global tPrevInstance As Integer ' temp variable for multiple instances
Global DoAnalysisReport As Integer

' New JEOL system parameters
Global JEOLVelocity(1 To MAXMOT%) As Long
Global JEOLBacklash(1 To MAXMOT%) As Long

' New variables for Monitor app
Global ScanComboLabels(1 To MAXMONITOR%) As String
Global ScanComboCommands(1 To MAXMONITOR%) As String
Global ScanComboNumberOf(1 To MAXMONITOR%) As Integer
Global ScanComboNames(1 To MAXMONITOR%, 1 To MAXMONITORLIST%) As String
Global ScanComboParameters(1 To MAXMONITOR%, 1 To MAXMONITORLIST%) As String

Global AlwaysPollFaradayCupStateFlag As Boolean         ' change to boolean 12-20-2018
Global DriverLoggingLevel As Long                       ' 0 - disabled, 1 - basic logging, 2 - detailed
Global ThermalFieldEmissionPresentFlag As Integer       ' 0 = not present, <> 0 present

Global JeolCondenserCoarseCalibrationSettingLow(1 To MAXAPERTURES%) As Integer
Global JeolCondenserCoarseCalibrationSettingMedium(1 To MAXAPERTURES%) As Integer
Global JeolCondenserCoarseCalibrationSettingHigh(1 To MAXAPERTURES%) As Integer
Global JeolCondenserFineCalibrationSetting As Integer

Global JeolCondenserCoarseCalibrationMode As Integer    ' 0 = internal measure calibration, read calibration from INI file
Global JeolCondenserCoarseCalibrationBeamLow(1 To MAXAPERTURES%) As Single
Global JeolCondenserCoarseCalibrationBeamMedium(1 To MAXAPERTURES%) As Single
Global JeolCondenserCoarseCalibrationBeamHigh(1 To MAXAPERTURES%) As Single
Global JeolCondenserNumberOfApertures As Integer        ' default = 1 (one value per line)

Global UseParticleCorrectionFlag As Integer
Global BeamCurrentToleranceSet As Single    ' tolerance for setting beam current (8200 and 8900 only)
Global PHAAdjustPresent As Integer

Global XrayColor(1 To MAXRAY% - 1) As Long
Global XrayColor2(1 To MAXRAY% - 1) As Integer
Global SampleExchangePositions(1 To MAXAXES%) As Single

Global Jeolhandle As Long
Global JeolMonitorInterval As Long      ' JEOL monitor packet interval (in msec)

Global UseQuickStandardsMode As Integer     ' 0 = normal, 1 = majors
Global UseQuickStandardsMinimum As Single     ' weight percent

' Auto-focus INI values
Global AutoFocusOffset As Single            ' Z axis offset for autofocus adjustment (in stage units)
Global AutoFocusMaxDeviation As Single      ' maximum allowed deviation for optical data fit for autofocus adjustment (in stage units)
Global AutoFocusThresholdFraction As Single, AutoFocusMinimumPtoB As Single
Global AutoFocusRangeFineScan As Single, AutoFocusRangeCoarseScan As Single
Global AutoFocusPointsFineScan As Long, AutoFocusPointsCoarseScan As Long           ' 50-1000 points
Global AutoFocusTimeFineScan As Integer, AutoFocusTimeCoarseScan As Integer         ' 1-500 msec

' Auto-focus measurements (for FormSCAN real time display of centroid and fit)
Global AutoFocusROMPeakMotor As Integer         ' 0 = auto-focus, or 1 - MAXSPEC% (only used by TestType)
Global AutoFocusROMPeakMaxPoints As Long        ' maximum number of points in data array (last index)
Global AutoFocusROMPeakData() As Single         ' (0 to MAXSPEC%, 1 to MAXAUTOFOCUSSCANS%, 1 to 2, 1 to npts&) autofocus z-axis and optical intensity or ROM peak spec position and x-ray intensity data

Global AutoFocusROMPeakPoints(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%) As Long     ' number of points in data array
Global AutoFocusROMPeakFitLastMode(0 To MAXSPEC%) As Integer  ' last mode run (1 = fine, 2 = coarse, 3 = 2nd fine)
Global AutoFocusROMPeakThreshold(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%) As Single
Global AutoFocusROMPeakMeasuredPtoB(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%) As Single
Global AutoFocusROMPeakFitCoeff(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%, 1 To MAXCOEFF%) As Single
Global AutoFocusROMPeakCentroid(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%) As Single
Global AutoFocusROMPeakDeviation(0 To MAXSPEC%, 1 To MAXAUTOFOCUSSCANS%) As Single

' Image acquisition globals
Global NumberOfImages As Integer    ' number of images in run
Global ImageSmps(1 To MAXIMAGES%) As Integer
Global ImageNums(1 To MAXIMAGES%) As Integer
Global ImageNams(1 To MAXIMAGES%) As String
Global ImageADs(1 To MAXIMAGES%) As Integer
Global ImageIxs(1 To MAXIMAGES%) As Integer
Global ImageIys(1 To MAXIMAGES%) As Integer
Global ImageMags(1 To MAXIMAGES%) As Single
Global ImageTitles(1 To MAXIMAGES%) As String

Global ImageXMins(1 To MAXIMAGES%) As Single
Global ImageXMaxs(1 To MAXIMAGES%) As Single
Global ImageYMins(1 To MAXIMAGES%) As Single
Global ImageYMaxs(1 To MAXIMAGES%) As Single

Global DefaultImageChannelNumber As Integer     ' image acquisition channel
Global DefaultImagePaletteNumber As Integer     ' image acquisition palette
Global DefaultImageAnalogAverages As Integer    ' image A/D conversions
Global DefaultImageIx As Integer                ' image pixel sizes
Global DefaultImageIy As Integer
Global DefaultImageAnalogUnits As String

Global ImageFlag As Integer                 ' 0 = called from Acquire, 1 = called from Digitize
Global ImagePaletteNumber As Integer        ' 0 = gray, 1 = thermal, 2 = rainbow, 3 = blue-red, 4 = custom
Global ImagePaletteArray(0 To BIT8&) As Long
Global UseImageAutomateModeOnStds As Integer
Global UseImageAutomateModeOnUnks As Integer
Global UseImageAutomateModeOnWavs As Integer
Global UseImageAutomateModes As Integer         ' 1 = before, 2 = after, 3 = both, 4 = after confirm only

Global AcquireFirstSampleStarted As Integer     ' flag for before and after image acquisition automation

Global DefaultBiasScanCountTime As Single
Global DefaultBiasScanIntervals As Integer
Global DefaultGainScanCountTime As Single
Global DefaultGainScanIntervals As Integer

Global UseWideROMPeakScanAlwaysFlag As Integer          ' for John Fournelle
Global UseCurrentConditionsOnStartUpFlag As Integer     ' for Chi Ma
Global UseCurrentConditionsAlwaysFlag As Integer        ' for Paul Carpenter

Global UseROMBasedSpectrometerScanFlag As Integer
Global DisplayPHAParameterDialogPriorFlag As Integer
Global DisplayPHAParameterDialogAfterFlag As Integer

Global DoNotDisplayStandardImagesDuringDigitizationFlag As Integer
Global NumberOfScans As Long    ' number of scans in run

Global DefaultBeamMode As Integer               ' 0 = analog spot, 1 = analog scan, 2 = digital spot
Global BeamModeString(0 To 2) As String
Global DefaultMagnification As Single           ' current magnification
Global DefaultMagnificationDefault As Single    ' default magnification
Global DefaultMagnificationAnalytical As Single ' analytical magnification
Global DefaultMagnificationImaging As Single    ' imaging magnification

Global DefaultBeamCenterXYZ(1 To MAXAXES%) As Single
Global DefaultBeamDeflectionXYZ(1 To MAXAXES%) As Single

Global BeamModePresent As Integer
Global BeamModeType As Integer

Global MinMagWindow As Single, MaxMagWindow As Single
Global DefaultAcquireString As String

Global DefaultMaximumOrder As Integer
Global RomanNum(1 To MAXKLMORDER%) As String * MAXKLMORDERCHAR%
Global DefaultKLMSpecificElement As Integer

Global UseSharedMonitorDataFlag As Integer
Global MonitorProcessStartedFlag As Integer

Global RealTimePHABaselines(1 To MAXSPEC%) As Single
Global RealTimePHAWindows(1 To MAXSPEC%) As Single
Global RealTimePHAGains(1 To MAXSPEC%) As Single
Global RealTimePHABiases(1 To MAXSPEC%) As Single
Global RealTimePHAModes(1 To MAXSPEC%) As Integer

Global DefaultPeakCenterSkipPBCheck As Integer

Global ROMPeakingParabolicThresholdFraction As Single   ' new version 6.58
Global ROMPeakingMaximaThresholdFraction As Single
Global ROMPeakingGaussianThresholdFraction As Single
Global ROMPeakingMaxDeviation As Single

Global AutomatedPHAParameterDialogPriorFlag As Integer  ' stored in version 6.61
Global AutomatedPHAParameterDialogAfterFlag As Integer
Global AutomatedPHAParameterDialogTypeFlag As Integer   ' stored in version 8.32

Global ImageDisplaySizeInCentimeters As Single

Global AnalysisCheckForSamePeakPositions As Integer  ' new for v. 6.62
Global AnalysisCheckForSamePHASettings As Integer

Global Light_Reflected_Transmitted As Long    ' 0 = reflected, 1 = transmitted

Global NumberofForbiddenElements As Integer, DefaultNumberofForbiddenElements As Integer
Global ForbiddenElements(1 To MAXFORBIDDEN%) As Integer

' New Cameca system parameters
Global SX100Velocity(1 To MAXMOT%) As Long
Global LimitToLimit As Single       ' for spectrometer motion
Global LimitToLimit2 As Single      ' for stage motion

Global AutomationTotalTime As Variant
Global AutomationStartTime As Variant
Global ProbeImageStartTime As Variant
Global ProbeImageElapsedTime As Variant

Global UseOnlyDigitizedStandardPositionsFlag As Boolean
Global DefaultReflectedLightIntensity As Integer
Global DefaultTransmittedLightIntensity As Integer

Global DefaultMatchStandardDatabase As String
Global DefaultVacuumUnitsType As Integer        ' 0 = Pascals, 1 = Torr, 2 = mBar

Global ImageClicked As Boolean          ' FormIMAGE and FormIMAGECALIBRATE flag

Global ImageAutoBrightnessContrastSEGain As Integer
Global ImageAutoBrightnessContrastSEOffset As Integer
Global ImageAutoBrightnessContrastBSEGain As Integer
Global ImageAutoBrightnessContrastBSEOffset As Integer

Global DisableSpectrometerNumber As Integer     ' spectrometer disable flag

Global AnalBlankCorrectionPercents(1 To MAXCHAN%) As Single
Global ImageAlternateScaleBarUnits As Integer   ' (0 = none, 1 = nm, 2 = um, 3 = mm, 4 = cm, 5 = meters, 6 = microinches, 7 = milliinches, 8 = inches)
Global CalculateAllMatrixCorrections As Boolean
Global ForceNegativeKratiosToZeroFlag As Integer

Global DefaultStandardCoatingFlag As Integer    ' 0 = not coated, 1 = coated
Global DefaultStandardCoatingElement As Integer
Global DefaultStandardCoatingDensity As Single
Global DefaultStandardCoatingThickness As Single    ' in angstroms

Global DefaultSampleCoatingFlag As Integer    ' 0 = not coated, 1 = coated
Global DefaultSampleCoatingElement As Integer
Global DefaultSampleCoatingDensity As Single
Global DefaultSampleCoatingThickness As Single    ' in angstroms

Global UseConductiveCoatingCorrectionForElectronAbsorption As Boolean
Global UseConductiveCoatingCorrectionForXrayTransmission As Boolean

Global DisableFullQuantInterferenceCorrectionFlag As Integer
Global DisableMatrixCorrectionInterferenceCorrectionFlag As Integer

Global PENEPMA_Root As String
Global PENDBASE_Path As String
Global PENEPMA_Path As String
Global PENEPMA_PAR_Path As String   ' for network shared PAR folder

' Thin film globals
Global TF_AllOrAverageFlag As Integer, TF_SkipOrSetZeroValuesFlag As Integer
Global TF_OutputType As Integer, TF_DoNotExportStandardCompositions As Integer

Global TF_CoatStdFlag As Integer, TF_CoatStdElm As Integer    ' coating flag, element
Global TF_CoatStdDensity As Single, TF_CoatStdThickness As Single   ' density, thickness

Global TF_CoatUnkFlag As Integer, TF_CoatUnkElm As Integer    ' coating flag, element
Global TF_CoatUnkDensity As Single, TF_CoatUnkThickness As Single   ' density, thickness

' Sample description
Global TF_SampleDescriptionFlag As Integer, TF_HomogeneousOrReplicateFlag As Integer

' Homogeneous layer
Global TF_HomogeneousLayerDensity As Single, TF_HomogeneousLayerThickness As Single
Global TF_HomogeneousLayerSiliconFlag As Integer

' Replicate layers
Global TF_ReplicateLayerNum As Integer

' Substrate element
Global TF_SubstrateElm As Integer, TF_ReplicateUseKnownThickness As Integer
Global TF_SubstrateAsOxide As Integer, TF_SubstrateOption As Integer, TF_SubstrateStandard As Integer

Global TotalAcquisitionTime As Variant
Global CurrentAcquisitionStartTime As Variant
Global CurrentAcquisitionStopTime As Variant

Global SpectrometerROMScanMode As Integer   ' SX100/SXFive only, 0 = absolute scan, 1 = relative scan

Global UseAutomationReStandardization As Boolean
Global AutomationReStandardizationOn As Boolean
Global AutomationReStandardizationStart As Double
Global AutomationReStandardizationInterval As Double    ' in days

Global ZAFEquationMode As Integer       ' (only used in CalcZAF)
Global CalcZAFMode As Integer           ' (only used in CalcZAF (ZAFPrintStd))

' 0 = beam deflection, 1 = random point
' 2 = traverse start, 3 = traverse stop, 4 = traverse digitize
' 5 = grid start, 6 = grid stop, 7 = grid digitize
' 8 = move stage
Global DigitizeMode As Integer
Global DigitizeImageSkipBeamModes As Boolean

Global DefaultAperture As Integer   ' (default = aperture 1)

Global DigitizeAutoIncrementFlag As Integer
Global DigitizeAutoIncrementNumber As Integer
Global DigitizeAutoDigitizeFlag As Integer

Global JeolCoarseCondenserCalibrationDelay As Single

Global DefaultImageShiftX As Single    ' change from integer for SX100/SXFive (10-29-2011)
Global DefaultImageShiftY As Single    ' change from integer for SX100/SXFive (10-29-2011)

Global PHAFirstTimeDelay As Single      ' in seconds when PHA is first set (for large bias change issues)

Global ROMPeakingString2(0 To 2) As String
Global CurrentROMPeakingSet(1 To MAXSPEC%) As Integer     ' 0 = fine, 1 = coarse, 2 = 2nd fine

Global AutomationProgressReportTime As Variant
'Global Const AutomationProgressReportInterval As Variant = 0.001    ' in days (86 seconds, for testing only)
Global Const AutomationProgressReportInterval As Variant = 0.333   ' in days (8 hours)

Global ImportIncrementXIncrement As Single, ImportIncrementYIncrement As Single   ' in microns
Global ImportIncrementXFactor As Integer, ImportIncrementYFactor As Integer
Global ImportIncrementXTotal As Single, ImportIncrementYTotal As Single

Global AutoIncrementDelimiterString As String

Global SX100MinimumSpeeds(1 To MAXMOT%) As Long

Global UseLastUnknownAsWavescanSetupFlag As Integer

Global DoAnalysisOutputFlag As Boolean      ' to suppress output for blank analysis

Global UserSpecifiedOutputSampleNameFlag As Boolean     ' user specified custom output flags
Global UserSpecifiedOutputLineNumberFlag As Boolean
Global UserSpecifiedOutputWeightPercentFlag As Boolean
Global UserSpecifiedOutputOxidePercentFlag As Boolean
Global UserSpecifiedOutputAtomicPercentFlag As Boolean
Global UserSpecifiedOutputTotalFlag As Boolean
Global UserSpecifiedOutputDetectionLimitsFlag As Boolean
Global UserSpecifiedOutputPercentErrorFlag As Boolean
Global UserSpecifiedOutputStageXFlag As Boolean
Global UserSpecifiedOutputStageYFlag As Boolean
Global UserSpecifiedOutputStageZFlag As Boolean
Global UserSpecifiedOutputRelativeDistanceFlag As Boolean
Global UserSpecifiedOutputOnPeakTimeFlag As Boolean
Global UserSpecifiedOutputHiPeakTimeFlag As Boolean
Global UserSpecifiedOutputLoPeakTimeFlag As Boolean
Global UserSpecifiedOutputOnPeakCountsFlag As Boolean
Global UserSpecifiedOutputOffPeakCountsFlag As Boolean
Global UserSpecifiedOutputNetPeakCountsFlag As Boolean
Global UserSpecifiedOutputKrawFlag As Boolean
Global UserSpecifiedOutputDateTimeFlag As Boolean

Global FilamentWarmUpInterval As Single     ' filament warmup interval delay in seconds

Global SampleSyms(1 To MAXSAMPLETYPES%) As String
Global InterfSyms(1 To MAXINTF%) As String

' Global monitor variables for updating StageMap window control
Global MonitorStateBeamMode As Integer          ' 0 = spot, 1 = scan, 2 = digital
Global MonitorStateLightMode As Integer         ' 0 = reflected, 1 = transmitted, 2 = both off
Global MonitorStateMagnification As Single
Global MonitorStateReflected As Integer         ' 0 = off, 1 = on
Global MonitorStateTransmitted As Integer       ' 0 = off, 1 = on

Global AcquireVolatileSelfStandardIntensitiesFlag As Boolean   ' true = acquire TDI data on unknown and also standard samples

Global WaveScanMeasureFaradayNthPoint As Integer
Global UseCountOverwriteIntensityDataFlag As Boolean

Global BgdStrings(0 To MAXOFFBGDTYPES%) As String   ' long strings for off-peak backgrounds (lower case)
Global BgStrings(0 To MAXOFFBGDTYPES%) As String    ' short strings for off-peak backgrounds (upper case)
Global BglStrings(0 To MAXOFFBGDTYPES%) As String   ' long strings for off-peak backgrounds (upper case)

Global BgdTypeStrings(0 To 2) As String             ' background type strings (0 = off-peak, 1 = MAN, 2 = Multi-Point bgd)
Global BeamModeStrings(0 To 2) As String            ' beam mode strings (0 = spot, 1 = scan, 2 = digital)

Global DefaultMultiPointNumberofPointsAcquireHi As Integer
Global DefaultMultiPointNumberofPointsAcquireLo As Integer
Global DefaultMultiPointNumberofPointsIterateHi As Integer
Global DefaultMultiPointNumberofPointsIterateLo As Integer

Global a08 As String
Global A10 As String
Global a12 As String
Global a14 As String
Global a16 As String
Global a18 As String
Global a22 As String
Global a24 As String
Global a32 As String
Global a64 As String

Global ProbeforEPMAQuickStartGuide As String
Global UseFluorescenceByBetaLinesFlag As Boolean

Global GeologicalSortOrderFlag As Integer
Global TurnOffSEDetectorBeforeAcquisitionFlag As Integer
Global TimeStampMode As Boolean

Global UseUnknownCountTimeForInterferenceStandardFlag As Integer
Global UserSpecifiedOutputZAFFlag As Boolean
Global UserSpecifiedOutputMACFlag As Boolean
Global UserSpecifiedOutputKratioFlag As Boolean

Global ProbeforEPMAFAQ As String
Global ProbeForEPMAConstantKRatios As String

Global UserSpecifiedOutputStdAssignsFlag As Boolean
Global InterfaceTypeStored As Integer        ' stored instrument interface type for post-processing
Global JEOLEOSInterfaceTypeStored As Long        ' stored instrument interface type for post-processing
Global EDSSpectraInterfaceTypeStored As Integer        ' stored EDS interface type for post-processing
Global ImageInterfaceTypeStored As Integer        ' stored image interface type for post-processing

Global InterfaceString(0 To MAXINTERFACE%) As String
Global InterfaceStringEDS(0 To MAXINTERFACE_EDS%) As String
Global InterfaceStringImage(0 To MAXINTERFACE_IMAGE%) As String

' Used for realtime TDI acquisition and also in demo mode to simulate volatile correction
Global VolatileSelfFaradayCupOutNow As Variant      ' time of day the faradaycup is removed
Global VolatileSelfFaradayCupStartNow(1 To MAXCHAN%) As Variant    ' time of day the count integration begins

Global UseVolElTimeWeightingFlag As Boolean
Global VolElTimeWeightingFactor As Integer

Global ImagePlotAspectControl As Integer
Global ImagePlotStagePositions As Integer
Global ImagePlotDataBar As Integer
Global ImagePlotLineNumbers As Integer
Global ImagePlotLineNumbersShort As Integer

Global ImageSkipDeletedPoints As Integer
Global ImageSkipDuplicatePoints As Integer
Global ImagePlotSampleName As Integer
Global ImagePlotSampleNameType As Integer

Global ImagePlotNthPoint As Integer
Global ImagePlotScaleBarPosition As Integer
Global ImagePlotSampleOrAllPositions As Integer

Global ImagePlotScaleBar As Integer
Global ImagePlotScaleBarMode As Integer

Global MinTotalValue As Single

Global UserSpecifiedOutputSampleNumberFlag As Boolean
Global UserSpecifiedOutputSampleConditionsFlag As Boolean

Global ForceNegativeInterferenceIntensitiesToZeroFlag As Boolean

Global MonthSyms(1 To 12) As String                  ' alphabetic month strings
Global MineralStrings(0 To MAXMINTYPES%) As String   ' five mineral end-member strings (including zero for none)

Global ThermoNSSLocalRemoteMode As Integer  '(0 = NSS and PFE on same computer, 1 = Thermo on remote computer)
Global UseDoNotSetConditionsFlag As Boolean

Global UserSpecifiedOutputFormulaFlag As Boolean
Global MonitorFontSize As Integer
Global JEOLSecurityNumber As Long

Global UseCurrentBeamBlankStateOnStartUpAndTerminationFlag As Boolean
Global ShowAllPeakingOptionsFlag As Boolean

Global UseRightMouseClickToDigitizeFlag As Boolean
Global UseChemicalAgeCalculationFlag As Boolean

Global ForceSetPHAParametersFlag As Boolean
Global AutomationOverheadPerAnalysis As Single          ' device dependent seconds per analysis point (fudge factor)

Global UseMANParametersFlag As Boolean

Global MotorVelocityChangeFlag(1 To MAXMOT%) As Long
Global MotorVelocityChangeSpeed(1 To MAXMOT%) As Single

Global IntegratedIntensityUseSmoothingFlag As Boolean
Global IntegratedIntensitySmoothingPointsPerSide As Integer ' must be an even number starting with 2

' Custom colors
Global VbDarkBlue As Long

Global DefaultDoNotRescaleKLMFlag As Boolean

Global ReflectedLightPresent As Boolean
Global TransmittedLightPresent As Boolean

Global DoNotUseFastQuantFlag As Boolean

Global ImageAnalogUnitsShortStrings(0 To MAXINTERFACE_IMAGE%) As String
Global ImageAnalogUnitsLongStrings(0 To MAXINTERFACE_IMAGE%) As String
Global ImageAnalogUnitsToolTipStrings(0 To MAXINTERFACE_IMAGE%) As String

Global KioskIOStatusString As String
Global KioskKilovolts As Single
Global KioskBeamCurrent As Single
Global KioskBeamSize As Single

Global FormPOSITIONisLoaded As Boolean
Global FormSTAGEMAPisLoaded As Boolean

Global UserSpecifiedOutputTotalPercentFlag As Boolean
Global UserSpecifiedOutputTotalOxygenFlag As Boolean
Global UserSpecifiedOutputTotalCationsFlag As Boolean
Global UserSpecifiedOutputCalculatedOxygenFlag As Boolean
Global UserSpecifiedOutputExcessOxygenFlag As Boolean
Global UserSpecifiedOutputZbarFlag As Boolean
Global UserSpecifiedOutputAtomicWeightFlag As Boolean
Global UserSpecifiedOutputOxygenFromHalogensFlag As Boolean
Global UserSpecifiedOutputHalogenCorrectedOxygenFlag As Boolean
Global UserSpecifiedOutputChargeBalanceFlag As Boolean
Global UserSpecifiedOutputFeChargeFlag As Boolean

Global InstrumentAcknowledgementString As String

Global AnalysisInSilentModeFlag As Boolean

Global MatrixMDBFile As String
Global PureMDBFile As String
Global BoundaryMDBFile As String

Global BinaryRanges(1 To MAXBINARY%) As Single  ' for matrix calculations

Global UsePenepmaKratiosFlag As Integer     ' 1 = no, 2 = yes
Global RealTimeBeamCurrentNumberofTimesSet As Long
Global RealTimeBeamCurrentNumberofTimesCalled As Long
Global LastBeamCurrentMeasured As Single

Global CalcImageQuantFlag As Boolean

Global OnPeakTimeFractionFlag As Boolean
Global OnPeakTimeFractionValue As Single

Global GRIDBB_BAS_File As String
Global GRIDCC_BAS_File As String

Global SLICEXY_BAS_File As String
Global POLYXY_BAS_File As String
Global MODALXY_BAS_File As String

Global STRIPXY1_BAS_File As String
Global STRIPXY2_BAS_File As String
Global STRIPXY3_BAS_File As String

Global UseSecondaryBoundaryFluorescenceCorrectionFlag As Boolean

Global UsePenepmaKratiosLimitFlag As Boolean    ' only fit up to limit concentrations for emitting element
Global PenepmaKratiosLimitValue As Single    ' fitting limit value for concentrations for emitting element

Global JEOLEIKSVersionNumber As Single     ' 2009 = 3, 2011 = 4, 2012 = 5

Global PenepmaMinimumElectronEnergy As Single      ' use for MSIMPA value in Penfluor.inp

Global SampleImportExportFlag As Boolean        ' for Julien Allaz

Global HysteresisPresentFlag As Boolean        ' for SX100/SXFIVE at the moment

Global UserImagesDirectory As String
Global OriginalUserImagesDirectory As String

Global UserEDSDirectory As String
Global UserCLDirectory As String
Global UserEBSDDirectory As String

' New variables for GRDInfo.ini
Global Default_X_Polarity As Integer
Global Default_Y_Polarity As Integer
Global Default_Stage_Units As String

Global CustomMDBFile As String

Global TempBMPFileName As String

Global GrapherAppDirectory As String
Global SurferAppDirectory As String

Global SurferPlotsPerPage As Integer
Global SurferPlotsPerPagePolygon As Integer
Global SurferPageSecondsDelay As Integer

Global UserSpecifiedOutputSpaceBeforeFlag As Boolean
Global UserSpecifiedOutputAverageFlag As Boolean
Global UserSpecifiedOutputStandardDeviationFlag As Boolean
Global UserSpecifiedOutputStandardErrorFlag As Boolean
Global UserSpecifiedOutputMinimumFlag As Boolean
Global UserSpecifiedOutputMaximumFlag As Boolean
Global UserSpecifiedOutputSpaceAfterFlag As Boolean

Global UserSpecifiedOutputUnkIntfCorsFlag As Boolean
Global UserSpecifiedOutputUnkMANAbsCorsFlag As Boolean
Global UserSpecifiedOutputUnkAPFCorsFlag As Boolean
Global UserSpecifiedOutputUnkVolElCorsFlag As Boolean
Global UserSpecifiedOutputUnkVolElDevsFlag As Boolean

Global UserSpecifiedOutputSampleDescriptionFlag As Boolean
Global UserSpecifiedOutputEndMembersFlag As Boolean
Global UserSpecifiedOutputOxideMolePercentFlag As Boolean

Global UseDefaultFocusFlag As Boolean

Global GettingStartedManual As String
Global AdvancedTopicsManual As String
Global UserReferenceManual As String

Global SX100MoveSpectroMilliSecDelayBefore As Long
Global SX100MoveSpectroMilliSecDelayAfter As Long

Global SX100MoveStageMilliSecDelayBefore As Long
Global SX100MoveStageMilliSecDelayAfter As Long

Global SX100ScanSpectroMilliSecDelayBefore As Long
Global SX100ScanSpectroMilliSecDelayAfter As Long

Global SX100FlipCrystalMilliSecDelayBefore As Long
Global SX100FlipCrystalMilliSecDelayAfter As Long

Global CalcImageProjectFile As String       ' main project file (full path) for CalcImage (needed for PictureSnap)

Global ProbewinHelpFile As String
Global CalcImageHelpFile As String
Global RemoteHelpFile As String
Global MatrixHelpFile As String

Global UserSpecifiedOutputStandards As Boolean
Global UserSpecifiedOutputUnknowns As Boolean

Global MinimumOverVoltageType As Integer    ' 0 = 2%, 1 = 10%, 2 = 20%

Global ProbeSoftwareInternetBrowseMethod As Integer  ' 0 = WWW, 1 = DVD

Global ImageShiftMinimumMag As Single

Global UserSpecifiedOutputStandardPublishedValuesFlag As Boolean
Global UserSpecifiedOutputStandardPercentVariancesFlag As Boolean
Global UserSpecifiedOutputStandardAlgebraicDifferencesFlag As Boolean

Global PHAHardwareTypeString(0 To 1) As String  ' 0 = trad. PHA, 1 = MCA PHA

Global ImageShiftPresent As Integer
Global ImageShiftType As Integer

Global vstring(0 To 2) As String

Global PHAMultiChannelMin As Single
Global PHAMultiChannelMax As Single

Global UserSpecifiedOutputTotalAtomsFlag As Boolean
Global UserSpecifiedOutputRelativeLineNumberFlag As Boolean

Global MosaicSizeX As Single     ' in stage units
Global MosaicSizeY As Single     ' in stage units
Global MosaicKilovolts As Single
Global MosaicMagnification As Single
Global MosaicImageIncrement As Long

Global TipOfTheDayFile As String
Global TipOfTheDayShowFromMenu As Boolean

Global UsePositionSampleMagKeVForAutomatedImaging As Boolean

Global OriginalUserEDSDirectory As String
Global OriginalUserCLDirectory As String
Global OriginalUserEBSDDirectory As String

Global OriginalCalcZAFDATDirectory As String
Global OriginalColumnPCCFileDirectory As String
Global OriginalSurferDataDirectory As String
Global OriginalGrapherDataDirectory As String
Global OriginalDemoImagesDirectory As String

Global EDSInterfaceInsertRetractPresent As Boolean
Global EDSInterfaceMaxEnergyThroughputPresent As Boolean

Global MaxEnergyArraySize As Integer
Global MaxEnergyArrayValue(1 To MAX_ENERGY_ARRAY_SIZE%) As Single
Global MaxThroughputArraySize As Integer
Global MaxThroughputArrayValue(1 To MAX_THROUGHPUT_ARRAY_SIZE%) As Single

Global EDSInterfaceMCSInputsPresent As Boolean

Global StrataGEMVersion As Single

Global ImageRGB1_R As String
Global ImageRGB1_G As String
Global ImageRGB1_B As String

Global DefaultEDSDeadtimePercent As Single

Global PHAFirstTimeDelaySet(1 To MAXSPEC%) As Boolean

Global DecontaminationTimeFlag As Boolean
Global DecontaminationTime As Single

Global CalcImageSurferOutputTemplateFlag As Integer   ' for CalcImage presentation output template (0 = default, 1 = custom1, 2 = custom2, etc.)
Global CalcImageSurferSliceTemplateFlag As Integer   ' for CalcImage presentation output template (0 = default, 1 = custom1, 2 = custom2)
Global CalcImageSurferPolygonTemplateFlag As Integer   ' for CalcImage presentation output template (0 = default, 1 = custom1, 2 = custom2)
Global CalcImageSurferStripTemplateFlag As Integer   ' for CalcImage presentation output template (0 = default, 1 = custom1, 2 = custom2)

Global UserSpecifiedOutputBeamCurrentFlag As Boolean
Global UserSpecifiedOutputAbsorbedCurrentFlag As Boolean
Global UserSpecifiedOutputBeamCurrent2Flag As Boolean
Global UserSpecifiedOutputAbsorbedCurrent2Flag As Boolean

Global ImageUseImageConditionsInDataBar As Integer

Global CLSpectraInterfacePresent As Boolean
Global CLSpectraInterfaceType As Integer                    ' 0 = demo, 1 = Ocean Optics, 2 = Gatan, 3 = Newport, 4 = not used yet
Global CLSpectraInterfaceTypeStored As Integer              ' 0 = demo, 1 = Ocean Optics, 2 = Gatan, 3 = Newport, 4 = not used yet
Global CLInterfaceInsertRetractPresent As Boolean

Global InterfaceStringCL(0 To MAXINTERFACE_CL%) As String
Global InterfaceStringCLUnitsX(0 To MAXINTERFACE_CL%) As String

Global AcquireCLSpectraFlag As Boolean
Global CLSpecifiedCountTime As Single
Global CLUnknownCountFactor As Single
Global CLDarkSpectraCountTimeFraction As Single

Global CLAcquisitionCountTime As Single     ' not stored in MDB

Global ProbeImageAcquisitionFile As String
Global ProbeImageSampleSetupNumber As Integer

Global UseConfirmDuringAutomationFlag As Boolean
Global RunProbeImageFlag As Boolean

Global EDSIntensityOption As Integer
Global CLIntensityOption As Integer

Global RunStandardsAfterProbeImageFlag As Boolean

Global UseStageReproducibilityCorrectionFlag As Boolean

Global ImageSizeIndex As Integer
Global ImageChannelNumber As Integer

Global ImageData_TDI_SampleName As String
Global ImageData_TDI_Kilovolts As Single
Global ImageData_TDI_nTDI As Integer     ' for CalcImage TDI pixel arrays (number of TDI intervals/images)
Global ImageData_TDI_Time As Single      ' for CalcImage TDI pixel arrays (TDI interval time in seconds)
Global ImageData_TDI_Beam1 As Single     ' for CalcImage TDI pixel arrays (TDI beam current- start)
Global ImageData_TDI_Beam2 As Single     ' for CalcImage TDI pixel arrays (TDI beam current- end)
Global ImageData_TDI_Ix As Long          ' for CalcImage TDI pixel arrays (always the same for all quant images)
Global ImageData_TDI_Iy As Long          ' for CalcImage TDI pixel arrays (always the same for all quant images)
Global ImageData_TDI_Start As Double     ' for CalcImage TDI pixel arrays (TDI start time)
Global ImageData_TDI_Stop As Double      ' for CalcImage TDI pixel arrays (TDI stop time)

Global ImageData_TDI() As Single         ' for CalcImage TDI pixel arrays (dimensioned in CalcImage) (1 to ImageData_TDI_Ix, 1 to ImageData_TDI_Iy, 1 to LastElm, 1 to nTDI)
Global ImageTime_TDI() As Double         ' for CalcImage TDI pixel arrays (dimensioned in CalcImage) (1 to ImageData_TDI_Ix, 1 to ImageData_TDI_Iy, 1 to LastElm, 1 to nTDI)

Global CLSpectrumAcquisitionOverhead As Single          ' overhead factor for light CL spectra

Global MDB_Template As String

Global OffPeakMarkerLabelFlag As Integer
Global ShowPeakMarkers As Integer, ShowDateTimeStamp As Integer, ShowGridLines As Integer
Global KLMOption As Integer, SGNumPointsPerSide As Integer

Global SlopePolyChanged As Integer  ' flag for changed background model (for all background types)
Global SlopePolyNumberofSamples As Integer
Global SlopePolyChannel As Integer

Global SlopePolySymbol As String
Global SlopePolyXray As String
Global SlopePolyCrystal As String
Global SlopePolyMotor As Integer

Global SlopePolyOffPeakType As Integer
Global SlopePolySampleRows() As Integer
Global ExponentialBase As Single
Global SlopeCoeff() As Single
Global PolyCoeff() As Single
Global PolyNominalBeam As Single
Global PolynomialIndex As Integer

Global ModelPeakingFitType As Integer

Global GraphWavescanType As Integer         ' 1 = spectrometer, 2 = angstroms, 3 = keV

Global CalcImageScanTypeFlag As Integer             ' 0 = beam scan, 1 = stage scan
Global CalcImageOrientationTypeFlag As Integer      ' 0 = Cameca stage/beam scan or JEOL beam scan (upper left/lower right), 1 = JEOL stage scan (upper right/lower left)

Global ThermoNSSVersionNumber As Single

Global DisplayFullScanRangeForAcquisitionFlag As Boolean

Global ManualPHAElementChannel As Integer   ' 0 = none, > 0 = channel to manually acquire PHA

' New globals for integrated intensity background fit (see global constant MAX_INTEGRATED_BGD_FIT% above)
Global IntegratedBackgroundFitType As Integer
Global IntegratedBackgroundFitPointsLow As Integer
Global IntegratedBackgroundFitPointsHigh As Integer

Global MoveStageToleranceX As Single    ' tolerance for initiating a stage move
Global MoveStageToleranceY As Single
Global MoveStageToleranceZ As Single

Global CalcImageAnalogSignalFlags(1 To 3) As Integer    ' leave as integer for CIP file input
Global CalcImageAnalogSignalLabels(1 To 3) As String

Global InstrumentFacility As String

Global UsePenepmaSimulationForDemoMode As Boolean

Global OriginalBeamCurrent As Single

Global UserSpecifiedOutputDetectionLimitsOxideFlag As Boolean

Global Penepma12UseKeVRoundingFlag As Boolean       ' flag to round keV to nearest integer (default = true)

Global InsertFaradayDuringStageJogFlag As Boolean

Global XtalFlipDuration As Single              ' Bragg crystal flip time

Global UseInterpolatedOffPeaksForMANFitFlag As Boolean                  ' flag to indicate user wants to utilize off-peak corrected standard intensities for MAN curves
Global UseInterpolatedOffPeaksForMANFitMode As Boolean                  ' temporary flag to indicate loading of MAN fit curves
Global UseInterpolatedIntensitiesEvenIfElementIsPresentFlag As Boolean

Global GrapherOutputVersionNumber As Integer         ' Scripter.exe app location moved from \Scripter to Surfer app folder

Global AnalyticalTotalMinimum As Single
Global AnalyticalTotalMaximum As Single

Global UseLineDrawingModeFlag As Boolean
Global UseRectangleDrawingModeFlag As Boolean

Global LAB6FieldEmissionPresentFlag As Integer       ' 0 = not present, <> 0 present

Global SkipPeakingJustDoPHAFlag As Integer

Global LoadFormulasFromStandardDatabaseFlag As Boolean
Global LoadCalculateOxygenFromStandardDatabaseFlag As Boolean

Global SetBeamModeAfterAcquisition As Integer

Global MaxInterfValue As Single

Global PhiRhoZPlotPoints As Long
Global PhiRhoZPlotSets As Long
Global PhiRhoZPlotX() As Single         ' new globals for phi-rho-z plot values
Global PhiRhoZPlotY1() As Single        ' generated intensities
Global PhiRhoZPlotY2() As Single        ' emitted intensities

Global DoNotUseEDSInterfaceForNetIntensitiesFlag As Integer         ' new flag to turn off EDS interface for obtaing net intensities

Global NthPointAcquisitionFlag As Boolean   ' for actual acquisition of Nth point data
Global NthPointCalculationFlag As Boolean   ' for test calculation of Nth Point data
Global DefaultNthPointAcquisitionInterval As Integer

' New flags for sample based Nth point acquisition
Global NthPointAcquisitionDefaultFlag As Boolean   ' for storing default Nth point acquisition flag
Global NthPointAcquisitionTypeFlag As Integer   ' 0 = acquire Nth points on both stds and unks, 1 = acquire Nth points on stds only, 2 = acquire Nth points on unks only

Global CalculatePhiRhoZPlotCurves As Boolean        ' used in CalcZAF

Global UserSpecifiedOutputChemAgeFlag As Boolean

Global SkipOutputEDSIntensitiesDuringAutomation As Boolean

Global UseZFractionZbarCalculationsFlag As Boolean         ' MAN Zbar calculations for continuum
Global ZFractionZbarCalculationsExponent As Single         ' MAN Zbar calculations for continuum

Global UseContinuumAbsCalculationsFlag As Boolean

Global UserSpecifiedOutputFerrousFerricFlag As Boolean

Global UserSpecifiedOutputMachineReadableFlag As Boolean

Global UseEDSStoredNetIntensitiesFlag As Boolean

Global MultiPointBackgroundFitTypeStrings(0 To 2) As String
Global MultiPointBackgroundFitTypeStrings2(0 To 2) As String

Global JEOLUnfreezeAfterFlag As Integer

Global CamecaSXFiveTactisUSBVideoFlag As Integer

Global DisableStageMoveAll As Integer               ' special flag to disable move stage and get status stage

Global SpecifiedAPFMaximumLineEnergy As Single

Global DefaultJEOLUnfreezeDelay As Integer

Global JEOLSpectrometerOrientationType As Integer
Global SpectrometerOrientations(0 To MAXSPEC%) As Integer   ' zero index for EDS spectrometer
    
Global AutomatedImageAcquisitionMagChangeMilliSecDelay As Long

Global FerricFerrousMethodStrings(0 To 6) As String

Global JEOLMoveSpectroMilliSecDelayAfter As Long

Global RunMultipleSetupsOneAtATimeFlag As Boolean
Global DoYIncrementOnMultipleSampleSetupsFlag As Integer

Global DeadTimeCorrectionTypeString(1 To 5) As String

Global JEOLMECLoggingFlag As Integer
Global JEOLMoveStageMilliSecDelayAfter As Long

Global DisableLightControl As Boolean       ' false = normal light control, true = disable reflected and transmitted light control

Global tempEDSSpectraInterfacePresent As Integer   ' temporary flag for EDSSpectraInterfacePresent

Global JEOLCountSpectroMilliSecDelayAfter As Long

Global SkipPolygonizationModeling As Boolean

Global UseOxygenFromSulfurCorrectionFlag As Integer
Global UserSpecifiedOutputOxygenFromSulfurFlag As Boolean
Global UserSpecifiedOutputSulfurCorrectedOxygenFlag As Boolean

Global MeasureAbsorbedFaradayCurrentOnlyOncePerSampleFlag As Boolean
Global DoNotMeasure2ndFaradayAbsorbedCurrentsFlag As Boolean

Global UseEffectiveTakeOffAnglesFlag  As Boolean

Global ZFractionBackscatterExponent As Single       ' for backscatter calculations (zero for variable exponent)

Global SecondaryDistanceMethodString(0 To 3) As String

Global JEOLFlipCrystalMilliSecDelayAfter As Long
